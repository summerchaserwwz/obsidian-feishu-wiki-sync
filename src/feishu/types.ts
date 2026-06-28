/**
 * 飞书 Open API 相关类型定义
 *
 * 覆盖认证、知识库（Wiki）、文档（Docx Block）、云盘（Drive）模块
 */

// ── 通用 ────────────────────────────────────────────────────────────────────

/** 飞书 API 通用响应包装 */
export interface FeishuResponse<T = unknown> {
  code: number;
  msg: string;
  data: T;
}

// ── 认证 ────────────────────────────────────────────────────────────────────

export interface TenantAccessTokenResponse {
  code: number;
  msg: string;
  tenant_access_token: string;
  expire: number;
}

// ── Wiki 知识库 ──────────────────────────────────────────────────────────────

/** 知识空间 */
export interface WikiSpace {
  space_id: string;
  name: string;
  description?: string;
  type: string;
}

/** 知识库节点 */
export interface WikiNode {
  space_id: string;
  node_token: string;
  /** 节点对应的实体 token（文档/表格等的实际 ID） */
  obj_token: string;
  obj_type: "docx" | "sheet" | "mindnote" | "bitable" | "slides" | "file";
  parent_node_token: string;
  node_type: "origin" | "shortcut";
  title: string;
  has_child: boolean;
}

/** 知识库节点列表响应 */
export interface WikiNodeListResponse {
  items: WikiNode[];
  page_token?: string;
  has_more: boolean;
}

/** 创建知识库节点的请求体 */
export interface CreateWikiNodeRequest {
  obj_type: "docx";
  node_type: "origin";
  title: string;
  parent_node_token?: string;
}

/** 创建知识库节点的响应 */
export interface CreateWikiNodeResponse {
  node: WikiNode;
}

// ── Docx 文档 ────────────────────────────────────────────────────────────────

/** 文档信息 */
export interface DocxDocument {
  document_id: string;
  revision_id: number;
  title: string;
}

/** 文本元素的样式 */
export interface TextElementStyle {
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  underline?: boolean;
  inline_code?: boolean;
  link?: { url: string };
}

/** 文本运行（一段连续样式文本） */
export interface TextRun {
  content: string;
  text_element_style?: TextElementStyle;
}

/** 文本元素（可以是 text_run，也可以是其他类型如 mention） */
export interface TextElement {
  text_run?: TextRun;
}

/** 文本块（段落/标题/列表项等通用内容） */
export interface TextBlockContent {
  elements: TextElement[];
  style?: {
    align?: "left" | "center" | "right";
    collapse?: boolean;
  };
}

/** 代码块内容 */
export interface CodeBlockContent {
  elements: TextElement[];
  style: {
    language: number; // 飞书代码语言枚举值
    wrap: boolean;
  };
}

/** 图片块内容 */
export interface ImageBlockContent {
  token: string;
  width?: number;
  height?: number;
}

/** 表格属性 */
export interface TableBlockProperty {
  row_size: number;
  column_size: number;
  column_width: number[];
  merge_info: { row_span: number; col_span: number }[];
}

/** 表格块内容（block_type: 31） */
export interface TableBlockContent {
  cells: string[];
  property: TableBlockProperty;
}

/** 表格单元格内容（用于嵌套块创建） */
export interface TableCellData {
  id: string;
  blocks: DocxBlock[];
}

/** 飞书文档 Block */
export interface DocxBlock {
  block_id?: string;
  block_type: number;
  parent_id?: string;
  children?: string[];
  // 各类块的内容（根据 block_type 填其中一个）
  text?: TextBlockContent;           // type 2 (paragraph)
  heading1?: TextBlockContent;       // type 3
  heading2?: TextBlockContent;       // type 4
  heading3?: TextBlockContent;       // type 5
  heading4?: TextBlockContent;       // type 6
  heading5?: TextBlockContent;       // type 7
  heading6?: TextBlockContent;       // type 8
  bullet?: TextBlockContent;         // type 12
  ordered?: TextBlockContent;        // type 13
  code?: CodeBlockContent;           // type 14
  quote?: TextBlockContent;          // type 15
  todo?: TextBlockContent & { style?: { done: boolean } }; // type 17
  image?: ImageBlockContent;         // type 27
  table?: TableBlockContent;         // type 31
  // 内部使用：表格单元格数据（不发送给飞书，仅用于构建嵌套块）
  _tableCells?: TableCellData[];
}

/** 创建子块的请求体 */
export interface CreateBlockChildrenRequest {
  children: DocxBlock[];
  index: number;
  document_revision_id?: number;
}

// ── Drive 云盘 ───────────────────────────────────────────────────────────────

/** 媒体上传响应 */
export interface MediaUploadResponse {
  file_token: string;
}

// ── 代码语言枚举（飞书定义） ────────────────────────────────────────────────

export const FEISHU_CODE_LANGUAGE: Record<string, number> = {
  plaintext: 1,
  abap: 2,
  ada: 3,
  apache: 4,
  apex: 5,
  "c++": 6,
  c: 7,
  cobol: 8,
  "c#": 9,
  css: 10,
  dart: 11,
  delphi: 12,
  django: 13,
  dockerfile: 14,
  erlang: 15,
  fortran: 16,
  "f#": 17,
  go: 18,
  groovy: 19,
  html: 20,
  html5: 21,
  java: 22,
  javascript: 23,
  json: 24,
  julia: 25,
  kotlin: 26,
  latex: 27,
  less: 28,
  lisp: 29,
  logo: 30,
  lua: 31,
  makefile: 32,
  markdown: 33,
  matlab: 34,
  mermaid: 35,
  nginx: 36,
  "objective-c": 37,
  openedgeabl: 38,
  pascal: 39,
  perl: 40,
  php: 41,
  prolog: 42,
  protobuf: 43,
  python: 44,
  r: 45,
  ruby: 46,
  rust: 47,
  scala: 48,
  scss: 49,
  shell: 50,
  sql: 51,
  swift: 52,
  typescript: 53,
  vbscript: 54,
  visual_basic: 55,
  xml: 56,
  yaml: 57,
};
