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
  return Array.from(node.childNodes ?? []).filter((child) => child.localName === local);
}

function firstChild(node: any, local: string): any | null {
  return directChildren(node, local)[0] ?? null;
}

function textOfElement(node: any): string {
  return node?.textContent ?? "";
}

function extractRun(node: any): DocxRun {
  const textParts: string[] = [];
  for (const child of Array.from(node.childNodes ?? [])) {
    if (child.localName === "t") textParts.push(textOfElement(child));
    else if (child.localName === "tab") textParts.push("\t");
    else if (child.localName === "br" || child.localName === "cr") textParts.push("\n");
    else if (child.localName === "noBreakHyphen") textParts.push("-");
  }
  const rPr = firstChild(node, "rPr");
  const has = (local: string) => !!rPr && directChildren(rPr, local).length > 0;
  return {
    text: textParts.join(""),
    bold: has("b") || has("bCs"),
    italic: has("i") || has("iCs"),
    underline: has("u"),
    strike: has("strike") || has("dstrike"),
  };
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

function applyToParagraph(node: any, para: DocxParagraph): void {
  const before = extractParagraph(node);
  // Text unchanged: keep the original runs (and any embedded fields/hyperlinks) intact.
  if (para.text === before.text && para.style === before.style) return;

  // Ensure a pPr exists and apply the requested style.
  const pPr = firstChild(node, "pPr") ?? createElement(node.ownerDocument, "w:pPr", W);
  applyStyleToPPr(pPr, para.style);
  if (pPr.parentNode !== node) node.insertBefore(pPr, node.firstChild);

  // Replace the runs with a single run carrying the requested formatting.
  for (const runNode of directChildren(node, "r")) node.removeChild(runNode);

  const first = para.runs[0] ?? before.runs[0];
  const run = createRun(node.ownerDocument, para.text, {
    bold: first?.bold,
    italic: first?.italic,
    underline: first?.underline,
    strike: first?.strike,
  });
  node.insertBefore(run, pPr.nextSibling ?? null);
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

function createRun(doc: any, text: string, props: { bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean }): any {
  const run = createElement(doc, "w:r", W);
  const rPr = createElement(doc, "w:rPr", W);
  if (props.bold) { rPr.appendChild(createElement(doc, "w:b", W)); rPr.appendChild(createElement(doc, "w:bCs", W)); }
  if (props.italic) { rPr.appendChild(createElement(doc, "w:i", W)); rPr.appendChild(createElement(doc, "w:iCs", W)); }
  if (props.underline) { const u = createElement(doc, "w:u", W); u.setAttribute("w:val", "single"); rPr.appendChild(u); }
  if (props.strike) { rPr.appendChild(createElement(doc, "w:strike", W)); }
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
