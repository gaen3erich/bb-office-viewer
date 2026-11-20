// Office Viewer — Word .docx structural editing.
//
// `docx-preview` renders a document for reading only. These helpers let the
// viewer also edit it: they read a .docx ZIP, extract the body paragraphs into
// a small model ({ style, text, runs }) and later rewrite the text (and run
// formatting) back into word/document.xml, leaving tables, images and every
// other part of the package untouched.
//
// This is deliberately a "structure, not WYSIWYG" editor: paragraphs are edited
// as blocks with a style and run formatting, not as free-form page layout.
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

export type DocxRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Text color as a hex string without a leading '#' (e.g. "FF0000"). */
  color?: string;
  /** Background / highlight color as a hex string without a leading '#' (e.g. "FFFF00"). */
  highlight?: string;
  /** Font family name (e.g. "Arial", "Times New Roman"). */
  font?: string;
  /** Font size in points (e.g. 12). */
  size?: number;
};

export type DocxParagraphStyle = "normal" | "title" | "subtitle" | "heading1" | "heading2" | "heading3" | "heading4" | "quote" | "listBullet" | "listNumber";

export type DocxParagraph = {
  style: DocxParagraphStyle;
  text: string;
  runs: DocxRun[];
  bullet?: boolean;
  level?: number;
  align?: "left" | "center" | "right" | "justify" | "both" | null;
};

export type DocxModel = { paragraphs: DocxParagraph[] };

const W = "w";
const VALID_STYLES: Record<DocxParagraphStyle, string> = {
  normal: "Normal",
  title: "Title",
  subtitle: "Subtitle",
  heading1: "Heading1",
  heading2: "Heading2",
  heading3: "Heading3",
  heading4: "Heading4",
  quote: "Quote",
  listBullet: "ListBullet",
  listNumber: "ListNumber",
};

/** Normalize a raw wordprocessing style id into our set of supported styles. */
function styleFromId(id: string | null | undefined): DocxParagraphStyle {
  const value = (id ?? "").toLowerCase();
  if (value === "title") return "title";
  if (value === "subtitle") return "subtitle";
  if (value.startsWith("heading")) {
    const level = parseInt(value.slice(7), 10);
    if (level >= 1 && level <= 4) return ("heading" + level) as DocxParagraphStyle;
    return "normal";
  }
  if (value.includes("quote")) return "quote";
  if (value.includes("bullet")) return "listBullet";
  if (value.includes("number")) return "listNumber";
  return "normal";
}

function directChildren(node: any, local: string): any[] {
  if (!node) return [];
  return (Array.from(node.childNodes ?? []) as any[]).filter((child) => child.localName === local);
}

function firstChild(node: any, local: string): any | null {
  return directChildren(node, local)[0] ?? null;
}

function textOfElement(node: any): string {
  return node?.textContent ?? "";
}

/** Normalize a colour value to an upper-case hex string without a leading '#'. Returns undefined for empty values. */
function normalizeColor(value: string | null | undefined): string | undefined {
  const v = (value ?? "").replace(/^#/, "").trim().toUpperCase();
  return v ? v : undefined;
}

/** Word highlight colour names mapped to hex, the only values w:highlight accepts. */
const HIGHLIGHT_HEX: Record<string, string> = {
  yellow: "FFFF00", green: "00FF00", cyan: "00FFFF", magenta: "FF00FF", blue: "0000FF", red: "FF0000",
  darkblue: "000080", darkcyan: "008080", darkgreen: "008000", darkmagenta: "800080",
  darkred: "800000", darkyellow: "808000", darkgray: "808080", lightgray: "D3D3D3", black: "000000", white: "FFFFFF",
};

/** Convert a Word highlight name to a hex colour when known; otherwise pass the value through. */
function highlightHex(value: string | null | undefined): string | undefined {
  const v = (value ?? "").trim().toLowerCase();
  if (!v) return undefined;
  return HIGHLIGHT_HEX[v] ?? normalizeColor(v);
}

function extractRun(node: any): DocxRun {
  const textParts: string[] = [];
  for (const child of Array.from(node.childNodes ?? []) as any[]) {
    if (child.localName === "t") textParts.push(textOfElement(child));
    else if (child.localName === "tab") textParts.push("\t");
    else if (child.localName === "br" || child.localName === "cr") textParts.push("\n");
    else if (child.localName === "noBreakHyphen") textParts.push("-");
  }
  const rPr = firstChild(node, "rPr");
  const has = (local: string) => !!rPr && directChildren(rPr, local).length > 0;
  const run: DocxRun = {
    text: textParts.join(""),
    bold: has("b") || has("bCs"),
    italic: has("i") || has("iCs"),
    underline: has("u"),
    strike: has("strike") || has("dstrike"),
  };
  if (rPr) {
    const color = normalizeColor(firstChild(rPr, "color")?.getAttribute("w:val"));
    if (color) run.color = color;
    const fonts = firstChild(rPr, "rFonts");
    const font = fonts?.getAttribute("w:ascii") ?? fonts?.getAttribute("w:hAnsi") ?? null;
    if (font) run.font = font;
    const sz = firstChild(rPr, "sz")?.getAttribute("w:val");
    const size = sz == null ? null : Math.round(parseInt(sz, 10) / 2);
    if (size) run.size = size;
    const shdFill = normalizeColor(firstChild(rPr, "shd")?.getAttribute("w:fill"));
    const highlight = highlightHex(firstChild(rPr, "highlight")?.getAttribute("w:val"));
    const background = shdFill ?? highlight;
    if (background) run.highlight = background;
  }
  return run;
}

function extractParagraph(node: any): DocxParagraph {
  const pPr = firstChild(node, "pPr");
  let style: DocxParagraphStyle = "normal";
  let bullet = false;
  let level = 0;
  let align: DocxParagraph["align"] = null;
  if (pPr) {
    const pStyle = firstChild(pPr, "pStyle");
    style = styleFromId(pStyle?.getAttribute("w:val"));
    const numPr = firstChild(pPr, "numPr");
    bullet = !!numPr;
    const ilvl = firstChild(numPr, "ilvl");
    const ilvlVal = ilvl?.getAttribute("w:val");
    level = ilvlVal == null ? 0 : parseInt(ilvlVal, 10) || 0;
    const jc = firstChild(pPr, "jc");
    const jcVal = jc?.getAttribute("w:val");
    align = (["left", "center", "right", "justify", "both"] as const).includes(jcVal) ? (jcVal as DocxParagraph["align"]) : null;
  }
  const runs = directChildren(node, "r").map(extractRun);
  const text = runs.map((run) => run.text).join("");
  return { style, text, runs, bullet, level, align };
}

function findBody(doc: any): any | null {
  const root = doc.documentElement;
  if (!root) return null;
  return directChildren(root, "body")[0] ?? null;
}

const xmlParser = new DOMParser();
const xmlSerializer = new XMLSerializer();

/** Extract all top-level body paragraphs of a .docx file into an editable model. */
export function extractDocxModel(bytes: Uint8Array): DocxModel {
  const files = unzipSync(bytes);
  const xml = strFromU8(files["word/document.xml"] ?? new Uint8Array());
  const doc = xmlParser.parseFromString(xml, "application/xml");
  const body = findBody(doc);
  const paragraphs = body ? directChildren(body, "p").map(extractParagraph) : [];
  return { paragraphs };
}

/** Apply an edited paragraph model to the package, preserving every other part. */
export function applyDocxModel(bytes: Uint8Array, paragraphs: DocxParagraph[]): Uint8Array {
  const files = unzipSync(bytes);
  const xml = strFromU8(files["word/document.xml"] ?? new Uint8Array());
  const doc = xmlParser.parseFromString(xml, "application/xml");
  const body = findBody(doc);
  const nodes = body ? directChildren(body, "p") : [];
  for (let i = 0; i < nodes.length && i < paragraphs.length; i++) {
    applyToParagraph(nodes[i], paragraphs[i]);
  }
  files["word/document.xml"] = strToU8(xmlSerializer.serializeToString(doc));
  return zipSync(files);
}

function formatSignature(run: DocxRun | undefined): string {
  return [run?.bold ?? false, run?.italic ?? false, run?.underline ?? false, run?.strike ?? false, run?.color ?? "", run?.highlight ?? "", run?.font ?? "", run?.size ?? 0].join("|");
}

function runsDiffer(paraRuns: DocxRun[], beforeRuns: DocxRun[]): boolean {
  if (paraRuns.length === 0) return false;
  // The format toolbar always rewrites the whole paragraph, so comparing the first
  // run's formatting is sufficient and avoids flattening mixed formatting on no-op.
  return formatSignature(paraRuns[0]) !== formatSignature(beforeRuns[0]);
}

function applyToParagraph(node: any, para: DocxParagraph): void {
  const before = extractParagraph(node);
  const textChanged = para.text !== before.text;
  const styleChanged = para.style !== before.style;
  const formatChanged = runsDiffer(para.runs, before.runs);
  // Nothing changed: leave the original runs (and any embedded fields/hyperlinks) intact.
  if (!textChanged && !styleChanged && !formatChanged) return;

  // Ensure a pPr exists and apply the requested style.
  const pPr = firstChild(node, "pPr") ?? createElement(node.ownerDocument, "w:pPr", W);
  applyStyleToPPr(pPr, para.style);
  if (pPr.parentNode !== node) node.insertBefore(pPr, node.firstChild);

  // Replace the runs. Formatting-only changes (from the WYSIWYG toolbar) keep the
  // individual runs so existing mixed formatting is preserved; text/style edits are
  // collapsed to a single run carrying the requested formatting.
  const runsToWrite = formatChanged && !textChanged ? para.runs : [para.runs[0] ?? { ...before.runs[0], text: para.text }];
  for (const runNode of directChildren(node, "r")) node.removeChild(runNode);
  const ref = pPr.nextSibling ?? null;
  for (const source of runsToWrite) {
    if (!source) continue;
    const run = createRun(node.ownerDocument, source.text ?? para.text, {
      bold: source.bold,
      italic: source.italic,
      underline: source.underline,
      strike: source.strike,
      color: source.color,
      highlight: source.highlight,
      font: source.font,
      size: source.size,
    });
    node.insertBefore(run, ref);
  }
}

function applyStyleToPPr(pPr: any, style: DocxParagraphStyle): void {
  const existing = firstChild(pPr, "pStyle");
  if (style === "normal") {
    if (existing) pPr.removeChild(existing);
    return;
  }
  const id = VALID_STYLES[style];
  if (existing) {
    existing.setAttribute("w:val", id);
  } else {
    const pStyle = createElement(pPr.ownerDocument, "w:pStyle", W);
    pStyle.setAttribute("w:val", id);
    pPr.insertBefore(pStyle, pPr.firstChild ?? null);
  }
}

function createRun(
  doc: any,
  text: string,
  props: { bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; color?: string; highlight?: string; font?: string; size?: number },
): any {
  const run = createElement(doc, "w:r", W);
  const rPr = createElement(doc, "w:rPr", W);
  if (props.bold) { rPr.appendChild(createElement(doc, "w:b", W)); rPr.appendChild(createElement(doc, "w:bCs", W)); }
  if (props.italic) { rPr.appendChild(createElement(doc, "w:i", W)); rPr.appendChild(createElement(doc, "w:iCs", W)); }
  if (props.underline) { const u = createElement(doc, "w:u", W); u.setAttribute("w:val", "single"); rPr.appendChild(u); }
  if (props.strike) { rPr.appendChild(createElement(doc, "w:strike", W)); }
  if (props.color) { const c = createElement(doc, "w:color", W); c.setAttribute("w:val", props.color); rPr.appendChild(c); }
  if (props.highlight) { const shd = createElement(doc, "w:shd", W); shd.setAttribute("w:val", "clear"); shd.setAttribute("w:color", "auto"); shd.setAttribute("w:fill", props.highlight); rPr.appendChild(shd); }
  if (props.font) { const f = createElement(doc, "w:rFonts", W); f.setAttribute("w:ascii", props.font); f.setAttribute("w:hAnsi", props.font); f.setAttribute("w:cs", props.font); rPr.appendChild(f); }
  if (props.size) {
    const s = createElement(doc, "w:sz", W); s.setAttribute("w:val", String(props.size * 2)); rPr.appendChild(s);
    const sc = createElement(doc, "w:szCs", W); sc.setAttribute("w:val", String(props.size * 2)); rPr.appendChild(sc);
  }
  if (rPr.childNodes.length > 0) run.appendChild(rPr);
  const t = createElement(doc, "w:t", W);
  t.setAttribute("xml:space", "preserve");
  t.textContent = text;
  run.appendChild(t);
  return run;
}

function createElement(doc: any, name: string, prefix: string): any {
  const qualified = name.includes(":") ? name : `${prefix}:${name}`;
  return doc.createElement(qualified);
}
