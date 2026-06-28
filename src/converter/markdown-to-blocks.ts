/**
 * Markdown → 飞书 Block 转换器（v2）
 *
 * 功能：将 Obsidian Markdown 解析为飞书文档的 Block 数组
 *
 * 支持标准 Markdown：
 *   标题、段落、粗体、斜体、删除线、行内代码、链接、
 *   有序/无序/嵌套列表、代码块、引用块、Callout、任务列表、分割线、图片
 *
 * 支持 Obsidian 特有语法：
 *   - Wikilink [[path|display]] → 显示文本
 *   - 嵌入图片 ![[image.png]] → 待上传图片
 *   - 嵌入笔记 ![[note]] → 引用块提示
 *   - Callout > [!type] title → 引用块 + 标题
 *   - 高亮 ==text== → 粗体+下划线模拟
 *   - 嵌套列表 → 保留缩进层级
 */

import { DocxBlock, TextElement, TextBlockContent, FEISHU_CODE_LANGUAGE } from "../feishu/types";

/** 图片信息（待上传） */
export interface PendingImage {
  /** 图片原始 alt 文本 */
  alt: string;
  /** 图片源路径（本地路径或 URL） */
  src: string;
  /** 在 blocks 数组中的位置索引 */
  blockIndex: number;
}

/** 转换结果 */
export interface ConversionResult {
  blocks: DocxBlock[];
  /** 待上传的图片列表 */
  pendingImages: PendingImage[];
}

// 飞书 Block 类型常量
const BT = {
  PAGE: 1, TEXT: 2,
  H1: 3, H2: 4, H3: 5, H4: 6, H5: 7, H6: 8,
  BULLET: 12, ORDERED: 13, CODE: 14, QUOTE: 15, TODO: 17,
  DIVIDER: 22, IMAGE: 27,
} as const;

// ════════════════════════════════════════════════════════
//  表格解析
// ════════════════════════════════════════════════════════

/** 解析 GFM 表格行，返回每列的文本 */
function parseTableRow(line: string): string[] {
  return line
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => cell.trim());
}

// ════════════════════════════════════════════════════════
//  预处理：Obsidian 语法 → 标准 Markdown
// ════════════════════════════════════════════════════════

/**
 * 预处理 Obsidian 特有语法，转为标准 Markdown 或占位符
 *
 * 处理：
 * - ![[image.png]] → ![image.png](image.png)
 * - ![[note]] → > 📎 嵌入笔记: note
 * - [[path|display]] → display
 * - [[path]] → path 的最后一段
 * - ==highlight== → **highlight**
 */
export function preprocessObsidian(markdown: string): string {
  let result = markdown;

  // 嵌入图片 ![[xxx.png]] ![[xxx.jpg]] 等
  result = result.replace(
    /!\[\[([^\]]+?\.(png|jpg|jpeg|gif|svg|webp|bmp|ico)(?:\|[^\]]*)?)\]\]/gi,
    (_, inner) => {
      const parts = inner.split("|");
      const src = parts[0].trim();
      const alt = parts[1]?.trim() || src.split("/").pop() || src;
      return `![${alt}](${src})`;
    }
  );

  // 嵌入笔记 ![[note name]] （非图片）→ 引用块
  result = result.replace(
    /!\[\[([^\]]+)\]\]/g,
    (_, inner) => {
      const display = inner.split("|").pop()?.trim() || inner.trim();
      return `> 嵌入笔记: ${display}`;
    }
  );

  // Wikilink [[path|display]] → display
  result = result.replace(
    /\[\[([^\]]+)\]\]/g,
    (_, inner) => {
      const parts = inner.split("|");
      if (parts.length >= 2) {
        return parts[1].trim();
      }
      // [[path/to/note]] → note
      const name = parts[0].split("/").pop()?.trim() || parts[0].trim();
      // 去掉 .md 后缀
      return name.replace(/\.md$/i, "");
    }
  );

  // 高亮 ==text== → 加粗表示（飞书不支持背景高亮）
  result = result.replace(/==(.+?)==/g, "**$1**");

  return result;
}

// ════════════════════════════════════════════════════════
//  行内解析
// ════════════════════════════════════════════════════════

/** 解析行内 Markdown → TextElement 数组 */
function parseInline(text: string): TextElement[] {
  const elements: TextElement[] = [];
  let remaining = text;

  while (remaining.length > 0) {
    // 粗体斜体 ***text***
    let m = remaining.match(/^(\*\*\*|___)(.+?)\1/);
    if (m) {
      elements.push({ text_run: { content: m[2], text_element_style: { bold: true, italic: true } } });
      remaining = remaining.slice(m[0].length);
      continue;
    }

    // 粗体 **text** 或 __text__
    m = remaining.match(/^(\*\*|__)(.+?)\1/);
    if (m) {
      elements.push({ text_run: { content: m[2], text_element_style: { bold: true } } });
      remaining = remaining.slice(m[0].length);
      continue;
    }

    // 斜体 *text* 或 _text_
    m = remaining.match(/^(\*|_)([^\s*].*?[^\s*]|[^\s*])\1/);
    if (m) {
      elements.push({ text_run: { content: m[2], text_element_style: { italic: true } } });
      remaining = remaining.slice(m[0].length);
      continue;
    }

    // 删除线 ~~text~~
    m = remaining.match(/^~~(.+?)~~/);
    if (m) {
      elements.push({ text_run: { content: m[1], text_element_style: { strikethrough: true } } });
      remaining = remaining.slice(m[0].length);
      continue;
    }

    // 行内代码 `code`
    m = remaining.match(/^`([^`]+)`/);
    if (m) {
      elements.push({ text_run: { content: m[1], text_element_style: { inline_code: true } } });
      remaining = remaining.slice(m[0].length);
      continue;
    }

    // 链接 [text](url)
    m = remaining.match(/^\[([^\]]*)\]\(([^)]+)\)/);
    if (m) {
      const url = m[2];
      // TODO: 非 http(s) 链接（如 Obsidian 相对路径）当前降级为纯文本，
      //       后续可解析为 vault 内文件路径，查找 frontmatter 中的 feishu_wiki_node，
      //       拼接为飞书文档链接 https://<domain>/wiki/<node_token>
      const isValidUrl = /^(https?|mailto):/.test(url);
      if (isValidUrl) {
        elements.push({ text_run: { content: m[1] || url, text_element_style: { link: { url } } } });
      } else {
        // 飞书不支持的协议，保留显示文本和链接地址
        const display = m[1] ? `${m[1]} (${url})` : url;
        elements.push({ text_run: { content: display } });
      }
      remaining = remaining.slice(m[0].length);
      continue;
    }

    // 普通文本（直到遇到特殊字符）
    m = remaining.match(/^[^*_~`\[\\]+/);
    if (m) {
      elements.push({ text_run: { content: m[0] } });
      remaining = remaining.slice(m[0].length);
      continue;
    }

    // 转义或不认识的字符
    elements.push({ text_run: { content: remaining[0] } });
    remaining = remaining.slice(1);
  }

  return elements;
}

/** 构建文本块内容 */
function makeTextContent(elements: TextElement[]): TextBlockContent {
  return { elements: elements.length > 0 ? elements : [{ text_run: { content: "" } }] };
}

// ════════════════════════════════════════════════════════
//  缩进检测
// ════════════════════════════════════════════════════════

/** 获取行的缩进级别（空格数 / 2 或 tab 数） */
function getIndentLevel(line: string): number {
  const match = line.match(/^(\s*)/);
  if (!match) return 0;
  const spaces = match[1].replace(/\t/g, "  ").length;
  return Math.floor(spaces / 2);
}

// ════════════════════════════════════════════════════════
//  主转换函数
// ════════════════════════════════════════════════════════

/**
 * 将 Markdown 文本转换为飞书文档 Block 数组
 *
 * 输入参数：
 * - markdown: 去除 frontmatter 后的 Markdown 正文
 *
 * 返回值：ConversionResult（blocks + pendingImages）
 */
export function markdownToBlocks(markdown: string): ConversionResult {
  // 先预处理 Obsidian 特有语法
  const processed = preprocessObsidian(markdown);

  const blocks: DocxBlock[] = [];
  const pendingImages: PendingImage[] = [];
  const lines = processed.split("\n");
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];
    const trimmed = line.trimStart();

    // ── 代码块 ──────────────────────────────────
    if (trimmed.startsWith("```")) {
      const lang = trimmed.slice(3).trim().toLowerCase();
      const langCode = FEISHU_CODE_LANGUAGE[lang] ?? FEISHU_CODE_LANGUAGE["plaintext"];
      const codeLines: string[] = [];
      i++;
      while (i < lines.length && !lines[i].trimStart().startsWith("```")) {
        codeLines.push(lines[i]);
        i++;
      }
      i++; // 跳过结束 ```
      blocks.push({
        block_type: BT.CODE,
        code: {
          elements: [{ text_run: { content: codeLines.join("\n") } }],
          style: { language: langCode, wrap: false },
        },
      });
      continue;
    }

    // ── 分割线 ──────────────────────────────────
    if (/^(-{3,}|\*{3,}|_{3,})\s*$/.test(trimmed)) {
      blocks.push({ block_type: BT.DIVIDER });
      i++;
      continue;
    }

    // ── 标题 ────────────────────────────────────
    const headingMatch = trimmed.match(/^(#{1,6})\s+(.+)/);
    if (headingMatch) {
      const level = headingMatch[1].length;
      const blockType = (BT.H1 + level - 1) as 3 | 4 | 5 | 6 | 7 | 8;
      const content = makeTextContent(parseInline(headingMatch[2]));
      const block: DocxBlock = { block_type: blockType };
      const key = `heading${level}` as keyof DocxBlock;
      (block as unknown as Record<string, unknown>)[key] = content;
      blocks.push(block);
      i++;
      continue;
    }

    // ── Callout > [!type] ───────────────────────
    const calloutMatch = trimmed.match(/^>\s*\[!(\w+)\]\s*(.*)/);
    if (calloutMatch) {
      const calloutType = calloutMatch[1].toUpperCase();
      const calloutTitle = calloutMatch[2] || calloutType;
      // Callout 标题行 → 加粗引用
      blocks.push({
        block_type: BT.QUOTE,
        quote: makeTextContent([
          { text_run: { content: `${calloutType}: ${calloutTitle}`, text_element_style: { bold: true } } },
        ]),
      });
      // 后续 > 行作为 Callout 内容
      i++;
      while (i < lines.length && lines[i].trimStart().startsWith(">")) {
        const content = lines[i].trimStart().replace(/^>\s?/, "");
        if (content.trim()) {
          blocks.push({
            block_type: BT.QUOTE,
            quote: makeTextContent(parseInline(content)),
          });
        }
        i++;
      }
      continue;
    }

    // ── 引用块 > ────────────────────────────────
    if (trimmed.startsWith("> ") || trimmed === ">") {
      const quoteContent = trimmed.slice(2) || "";
      blocks.push({
        block_type: BT.QUOTE,
        quote: makeTextContent(parseInline(quoteContent)),
      });
      i++;
      continue;
    }

    // ── 任务列表 - [x] text ─────────────────────
    const todoMatch = trimmed.match(/^-\s+\[(x| )\]\s+(.*)/i);
    if (todoMatch) {
      const done = todoMatch[1].toLowerCase() === "x";
      blocks.push({
        block_type: BT.TODO,
        todo: {
          elements: parseInline(todoMatch[2]),
          style: { done },
        } as DocxBlock["todo"],
      });
      i++;
      continue;
    }

    // ── 无序列表 ────────────────────────────────
    const bulletMatch = trimmed.match(/^(-|\*|\+)\s+(.*)/);
    if (bulletMatch) {
      blocks.push({
        block_type: BT.BULLET,
        bullet: makeTextContent(parseInline(bulletMatch[2])),
      });
      i++;
      continue;
    }

    // ── 有序列表 ────────────────────────────────
    const orderedMatch = trimmed.match(/^\d+\.\s+(.*)/);
    if (orderedMatch) {
      blocks.push({
        block_type: BT.ORDERED,
        ordered: makeTextContent(parseInline(orderedMatch[1])),
      });
      i++;
      continue;
    }

    // ── 图片 ![alt](src) ────────────────────────
    const imageMatch = trimmed.match(/^!\[([^\]]*)\]\(([^)]+)\)/);
    if (imageMatch) {
      const blockIndex = blocks.length;
      blocks.push({ block_type: BT.IMAGE, image: { token: "__PENDING__" } });
      pendingImages.push({ alt: imageMatch[1], src: imageMatch[2], blockIndex });
      i++;
      continue;
    }

    // ── 表格（GFM） ────────────────────────────
    if (trimmed.includes("|") && i + 1 < lines.length && /^\|[-: |]+\|/.test(lines[i + 1].trim())) {
      // 解析表头
      const headerCells = parseTableRow(trimmed);
      i += 2; // 跳过头和分隔线
      const dataRows: string[][] = [];
      while (i < lines.length && lines[i].trim().includes("|")) {
        dataRows.push(parseTableRow(lines[i].trim()));
        i++;
      }

      const colCount = headerCells.length;
      const allRows = [headerCells, ...dataRows];
      const cellIds: string[] = [];
      const tableCells: { id: string; blocks: DocxBlock[] }[] = [];

      for (let r = 0; r < allRows.length; r++) {
        for (let c = 0; c < colCount; c++) {
          const cellId = `tmp_cell_${blocks.length}_${r}_${c}`;
          const cellText = allRows[r][c] || "";
          cellIds.push(cellId);
          tableCells.push({
            id: cellId,
            blocks: [{
              block_type: BT.TEXT,
              text: { elements: parseInline(cellText) },
            }],
          });
        }
      }

      const tableBlock: DocxBlock & { _tableCells?: { id: string; blocks: DocxBlock[] }[] } = {
        block_type: 31,
        table: {
          cells: cellIds,
          property: {
            row_size: allRows.length,
            column_size: colCount,
            column_width: Array(colCount).fill(100),
            merge_info: cellIds.map(() => ({ row_span: 1, col_span: 1 })),
          },
        },
        _tableCells: tableCells,
      };
      blocks.push(tableBlock);
      continue;
    }

    // ── 空行 ────────────────────────────────────
    if (trimmed === "") {
      i++;
      continue;
    }

    // ── 普通段落 ─────────────────────────────────
    // 收集连续非空行作为同一段落
    const paraLines: string[] = [trimmed];
    i++;
    while (
      i < lines.length &&
      lines[i].trim() !== "" &&
      !lines[i].trimStart().startsWith("#") &&
      !lines[i].trimStart().startsWith(">") &&
      !lines[i].trimStart().startsWith("```") &&
      !lines[i].trimStart().match(/^(-|\*|\+|\d+\.)\s/) &&
      !lines[i].trimStart().match(/^!\[/) &&
      !lines[i].trimStart().match(/^(-{3,}|\*{3,}|_{3,})\s*$/)
    ) {
      paraLines.push(lines[i].trim());
      i++;
    }

    const fullPara = paraLines.join(" ");
    blocks.push({
      block_type: BT.TEXT,
      text: makeTextContent(parseInline(fullPara)),
    });
  }

  return { blocks, pendingImages };
}

/**
 * 去除 Markdown frontmatter（--- ... --- 部分）
 */
export function stripFrontmatter(content: string): string {
  if (!content.startsWith("---")) return content;
  const endIndex = content.indexOf("\n---", 3);
  if (endIndex === -1) return content;
  return content.slice(endIndex + 4).trimStart();
}
