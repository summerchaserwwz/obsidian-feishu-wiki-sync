/**
 * 同步引擎（核心）
 *
 * 功能：
 * - 单篇笔记同步（新建 / 更新）
 * - 批量文件夹同步（保留目录结构）
 * - 统一错误处理和重试
 *
 * 流程：
 * 1. 读取文件内容，去除 frontmatter
 * 2. 预处理 Obsidian 语法为标准 Markdown
 * 3. 飞书 convert API 将 Markdown 转为 Block
 * 4. 查找或创建知识库节点
 * 5. descendant API 覆盖写入文档内容
 * 6. 更新 frontmatter 同步元数据
 */

import { App, TFile, TFolder, Notice } from "obsidian";
import { FeishuWikiSettings } from "../settings/settings-types";
import { WikiApi } from "../feishu/wiki-api";
import { DocApi } from "../feishu/doc-api";
import { DriveApi } from "../feishu/drive-api";
import { WikiNode, DocxBlock } from "../feishu/types";
import { preprocessObsidian, stripFrontmatter } from "../converter/markdown-to-blocks";
import { ImageResolver } from "../converter/image-resolver";
import { SyncStateManager, SyncMeta } from "./sync-state";

/** 单次同步结果 */
export interface SyncResult {
  success: boolean;
  fileName: string;
  skipped?: boolean;
  error?: string;
}

export class SyncEngine {
  private imageResolver: ImageResolver;
  private syncState: SyncStateManager;

  constructor(
    private app: App,
    private settings: FeishuWikiSettings,
    private wikiApi: WikiApi,
    private docApi: DocApi,
    driveApi: DriveApi
  ) {
    this.imageResolver = new ImageResolver(app.vault, driveApi);
    this.syncState = new SyncStateManager(app);
  }

  /**
   * 同步单篇笔记到飞书知识库
   *
   * 输入参数：
   * - file: 要同步的 Obsidian 文件
   * - targetSpaceId: 目标知识空间 ID（不传则使用设置中的默认值）
   * - targetParentNodeToken: 目标父节点 token（不传则使用设置中的默认值）
   *
   * 返回值：SyncResult
   */
  /** 文件夹路径 → 知识库节点 token 的缓存（会话内复用，避免重复创建） */
  private folderNodeCache = new Map<string, string>();

  async syncFile(
    file: TFile,
    targetSpaceId?: string,
    targetParentNodeToken?: string,
    syncRootPath?: string
  ): Promise<SyncResult> {
    const spaceId = targetSpaceId ?? this.settings.defaultSpaceId;
    const baseParentToken = targetParentNodeToken ?? this.settings.defaultParentNodeToken;

    if (!spaceId) {
      return {
        success: false,
        fileName: file.name,
        error: "未配置默认知识空间，请在插件设置中选择",
      };
    }

    try {
      if (!this.matchesFrontmatterFilter(file)) {
        return {
          success: true,
          skipped: true,
          fileName: file.name,
          error: "不符合 frontmatter 过滤条件，已跳过",
        };
      }

      // 读取文件内容
      const rawContent = await this.app.vault.read(file);
      const content = stripFrontmatter(rawContent);
      const title = file.basename;

      // 查找或创建知识库节点（文档）
      const existingMeta = this.syncState.readSyncMeta(file);
      let node: WikiNode;

      if (existingMeta?.nodeToken && existingMeta.spaceId === spaceId) {
        if (this.settings.updateStrategy === "skip-if-exists") {
          return {
            success: true,
            skipped: true,
            fileName: file.name,
            error: "已存在同步记录，按当前更新策略跳过",
          };
        }

        // 更新现有节点
        node = {
          space_id: spaceId,
          node_token: existingMeta.nodeToken,
          obj_token: existingMeta.docToken,
          obj_type: "docx",
          parent_node_token: baseParentToken,
          node_type: "origin",
          title,
          has_child: false,
        };
      } else {
        // 新建节点：从 syncFolder 调用时父节点已由 ensureFolderNode 解析，直接使用；
        // 单文件同步时需要根据文件路径确保父目录节点链存在
        const actualParentToken = syncRootPath
          ? baseParentToken
          : await this.ensurePathNodes(spaceId, file.path, baseParentToken);
        node = await this.wikiApi.createNode(spaceId, title, actualParentToken || undefined);
      }

      // 预处理 Obsidian 语法 → 标准 Markdown，再用飞书 convert API 转换并写入
      const processedContent = preprocessObsidian(content);
      if (processedContent.trim()) {
        const { blockIdRelations, imageUrlMap } = await this.docApi.overwriteDocumentFromMarkdown(node.obj_token, processedContent);

        // 处理图片：convert API 返回的 image block 需要单独上传素材
        await this.processImagesAfterConvert(imageUrlMap, blockIdRelations, node.obj_token, file.path);
      } else {
        // 空内容：清空文档
        await this.docApi.overwriteDocument(node.obj_token, []);
      }

      // 更新 frontmatter 同步元数据
      const doc = await this.docApi.getDocument(node.obj_token);
      const meta: SyncMeta = {
        spaceId,
        nodeToken: node.node_token,
        docToken: node.obj_token,
        lastSync: new Date().toISOString(),
        docRevision: doc.revision_id,
      };
      await this.syncState.writeSyncMeta(file, meta);

      return { success: true, fileName: file.name };
    } catch (err) {
      const error = err instanceof Error ? err.message : String(err);
      return { success: false, fileName: file.name, error };
    }
  }

  /**
   * 判断文件是否符合设置中的 frontmatter 过滤规则。
   *
   * 支持两种写法：
   * - `publish`：frontmatter 中存在 publish 字段即可同步
   * - `publish: true`：frontmatter 字段值必须匹配 true
   */
  private matchesFrontmatterFilter(file: TFile): boolean {
    const filter = this.settings.frontmatterFilter.trim();
    if (!filter) return true;

    const frontmatter = this.app.metadataCache.getFileCache(file)?.frontmatter;
    if (!frontmatter) return false;

    const separatorIndex = filter.indexOf(":");
    if (separatorIndex === -1) {
      return Object.prototype.hasOwnProperty.call(frontmatter, filter);
    }

    const key = filter.slice(0, separatorIndex).trim();
    const expected = filter.slice(separatorIndex + 1).trim();
    if (!key) return true;

    const actual = frontmatter[key];
    if (actual === undefined || actual === null) return false;

    return String(actual).trim().toLowerCase() === expected.toLowerCase();
  }

  /**
   * 处理 convert API 返回的图片块
   *
   * 流程：
   * 1. 从 imageUrlMap 获取临时 block ID 和对应的图片 URL
   * 2. 通过 blockIdRelations 找到真实 block ID
   * 3. 下载图片并上传到飞书（用 block ID 作 parent_node）
   * 4. 调用 replaceImage 更新 block
   */
  private async processImagesAfterConvert(
    imageUrlMap: { block_id: string; image_url: string }[],
    blockIdRelations: { temporary_block_id: string; block_id: string }[],
    documentId: string,
    sourceFilePath: string
  ): Promise<void> {
    if (imageUrlMap.length === 0) return;

    const relationMap = new Map(blockIdRelations.map((r) => [r.temporary_block_id, r.block_id]));

    for (const img of imageUrlMap) {
      const realBlockId = relationMap.get(img.block_id) ?? img.block_id;

      try {
        const fileToken = await this.imageResolver.resolveAndUpload(img.image_url, realBlockId, sourceFilePath);
        if (fileToken) {
          await this.docApi.replaceImage(documentId, realBlockId, fileToken);
        }
      } catch (err) {
        console.warn(`[FeishuSync] 图片处理失败 "${img.image_url}":`, err);
      }
    }
  }

  /**
   * 根据文件的 vault 路径，确保飞书知识库中对应的文件夹节点链存在
   *
   * 仅用于单文件同步场景；从 syncFolder 调用时父节点已由 ensureFolderNode 解析。
   *
   * 例如 file.path = "03-KNOWLEDGE/AI/prompt.md"
   * → 确保 "03-KNOWLEDGE" 节点存在（父为 baseParentToken）
   * → 确保 "AI" 节点存在（父为 "03-KNOWLEDGE" 节点）
   * → 返回 "AI" 节点的 token 作为文档的父节点
   */
  private async ensurePathNodes(
    spaceId: string,
    filePath: string,
    baseParentToken: string
  ): Promise<string> {
    const parts = filePath.split("/");
    // 去掉最后一个（文件名），只处理目录部分
    const dirParts = parts.slice(0, -1);

    if (dirParts.length === 0) {
      // 文件在 vault 根目录，直接放到配置的根节点下
      return baseParentToken;
    }

    let parentToken = baseParentToken;

    for (let i = 0; i < dirParts.length; i++) {
      const cacheKey = `${spaceId}:${dirParts.slice(0, i + 1).join("/")}`;

      if (this.folderNodeCache.has(cacheKey)) {
        parentToken = this.folderNodeCache.get(cacheKey)!;
        continue;
      }

      // 在飞书知识库中查找或创建对应的文件夹节点
      const folderNode = await this.wikiApi.findOrCreateNode(
        spaceId,
        dirParts[i],
        parentToken || undefined
      );
      this.folderNodeCache.set(cacheKey, folderNode.node_token);
      parentToken = folderNode.node_token;
    }

    return parentToken;
  }

  /**
   * 批量同步文件夹，保留目录结构
   *
   * 输入参数：
   * - folder: 要同步的文件夹
   * - targetSpaceId: 目标知识空间 ID
   * - targetParentNodeToken: 目标父节点 token
   * - onProgress: 进度回调（已完成数, 总数, 当前文件名）
   *
   * 返回值：SyncResult 数组
   */
  async syncFolder(
    folder: TFolder,
    targetSpaceId: string,
    targetParentNodeToken: string,
    onProgress?: (done: number, total: number, currentFile: string) => void,
    shouldCancel?: () => boolean
  ): Promise<SyncResult[]> {
    // 收集所有要同步的 .md 文件
    const files = this.collectMarkdownFiles(folder);
    const excludeFolders = this.settings.excludeFolders
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);

    const filteredFiles = files.filter((f) => {
      // 检查是否在排除目录中
      return !excludeFolders.some((ex) => f.path.includes(`/${ex}/`) || f.path.startsWith(`${ex}/`));
    });

    const results: SyncResult[] = [];
    /** 文件夹路径 → 知识库节点 token 的缓存（避免重复创建） */
    const folderNodeCache = new Map<string, string>();

    // 先在飞书创建根文件夹节点（使用文件夹名作为标题）
    const rootNode = await this.wikiApi.findOrCreateNode(
      targetSpaceId,
      folder.name,
      targetParentNodeToken || undefined
    );
    folderNodeCache.set(folder.path, rootNode.node_token);

    for (let i = 0; i < filteredFiles.length; i++) {
      if (shouldCancel?.()) {
        break;
      }

      const file = filteredFiles[i];
      onProgress?.(i, filteredFiles.length, file.name);

      // 确保父目录节点已创建
      const parentNodeToken = await this.ensureFolderNode(
        file,
        folder,
        targetSpaceId,
        folderNodeCache
      );

      const result = await this.syncFile(file, targetSpaceId, parentNodeToken, folder.path);
      results.push(result);
    }

    onProgress?.(filteredFiles.length, filteredFiles.length, "");
    return results;
  }

  /**
   * 确保文件所在的每一层父目录在飞书知识库中都有对应节点
   *
   * 返回值：文件直接父目录的节点 token
   */
  private async ensureFolderNode(
    file: TFile,
    syncRoot: TFolder,
    spaceId: string,
    cache: Map<string, string>
  ): Promise<string> {
    const relativePath = file.path.slice(syncRoot.path.length + 1);
    const parts = relativePath.split("/");
    // 去掉文件名，只保留目录部分
    const dirParts = parts.slice(0, -1);

    let parentToken = cache.get(syncRoot.path) ?? "";

    for (let depth = 0; depth < dirParts.length; depth++) {
      const currentPath = `${syncRoot.path}/${dirParts.slice(0, depth + 1).join("/")}`;

      if (!cache.has(currentPath)) {
        const node = await this.wikiApi.findOrCreateNode(spaceId, dirParts[depth], parentToken || undefined);
        cache.set(currentPath, node.node_token);
      }

      parentToken = cache.get(currentPath)!;
    }

    return parentToken;
  }

  /**
   * 递归收集文件夹内所有 .md 文件
   */
  private collectMarkdownFiles(folder: TFolder): TFile[] {
    const files: TFile[] = [];
    for (const child of folder.children) {
      if (child instanceof TFile && child.extension === "md") {
        files.push(child);
      } else if (child instanceof TFolder) {
        files.push(...this.collectMarkdownFiles(child));
      }
    }
    return files;
  }

  /**
   * 处理图片：上传所有 pendingImages，将 token 填入对应 block
   */
  private async processImages(
    blocks: DocxBlock[],
    pendingImages: { alt: string; src: string; blockIndex: number }[],
    documentId: string,
    sourceFilePath: string
  ): Promise<DocxBlock[]> {
    if (pendingImages.length === 0) return blocks;

    const processedBlocks = [...blocks];

    for (const pending of pendingImages) {
      const isRemoteImage = pending.src.startsWith("http://") || pending.src.startsWith("https://");
      if (!this.settings.uploadLocalImages && !isRemoteImage) {
        processedBlocks[pending.blockIndex] = this.createImageFallbackBlock(pending);
        continue;
      }

      const fileToken = await this.imageResolver.resolveAndUpload(
        pending.src,
        documentId,
        sourceFilePath
      );

      if (fileToken) {
        processedBlocks[pending.blockIndex] = {
          block_type: 27,
          image: { token: fileToken },
        };
      } else {
        // 上传失败：替换为显示图片路径的文本（不丢失信息）
        processedBlocks[pending.blockIndex] = this.createImageFallbackBlock(pending);
      }
    }

    return processedBlocks;
  }

  /** 图片未上传时写入文本占位，避免生成空 token 的非法图片块。 */
  private createImageFallbackBlock(pending: { alt: string; src: string }): DocxBlock {
    return {
      block_type: 2,
      text: {
        elements: [
          {
            text_run: {
              content: `[图片: ${pending.alt || pending.src}]`,
              text_element_style: { italic: true },
            },
          },
        ],
      },
    };
  }
}
