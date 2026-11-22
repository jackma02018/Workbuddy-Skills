/**
 * 发票信息提取（batch-invoice-processor 技能 · 文本层通道）
 *
 * 处理来源：
 *   .xml 增值税电子发票（结构化 XML）
 *   .ofd 增值税电子发票（ZIP 文本层）
 *   .pdf 含文本层的电子发票（零依赖抽取 + ToUnicode CMap 中文还原）
 *   .txt / 其它 纯文本
 *   .png/.jpg/.jpeg/.bmp/.webp 纯图片 / 图片型 PDF -> 标记为 needs_vision，交给本平台多模态读图（Read 工具）提取
 *
 * 依赖加载策略：优先本地 require；找不到时回退到 WorkBuddy 受管 Node 工作区（jszip / fast-xml-parser）。
 * 在无这些依赖的干净环境，XML/OFD 会降级为 needs_vision，PDF/文本仍可用。
 *
 * 用法：
 *   node extract.js <文件或目录> [--json 提取结果.json]
 * 输出：标准发票对象数组（见下方 canonical 字段），直接喂给 organize.js。
 */

const fs = require("fs");
const path = require("path");
const os = require("os");
const zlib = require("zlib");

// ---------------- 依赖加载（含受管 workspace 回退） ----------------
function loadDep(name) {
  try { return require(name); } catch (e) {}
  const searchPaths = [
    path.join(os.homedir(), ".workbuddy", "binaries", "node", "workspace", "node_modules"),
    path.join(__dirname, "..", "node_modules"),
    path.join(__dirname, "node_modules"),
  ];
  for (const base of searchPaths) {
    try { return require(path.join(base, name)); } catch (e) {}
  }
  return null;
}
const XMLParser = loadDep("fast-xml-parser");
const JSZip = loadDep("jszip");

// ---------------- 字段清洗 ----------------
function normalizeStr(s) { return (s || "").replace(/\s+/g, " ").trim(); }

// 中文大写金额 -> 数字（CID 字体导致小写金额无法解析时的兜底）
function uppercaseCnToNumber(s) {
  const digitMap = { "零": 0, "〇": 0, "一": 1, "壹": 1, "二": 2, "贰": 2, "两": 2, "三": 3, "叁": 3, "四": 4, "肆": 4, "五": 5, "伍": 5, "六": 6, "陆": 6, "七": 7, "柒": 7, "八": 8, "捌": 8, "九": 9, "玖": 9 };
  const unitMap = { "十": 10, "拾": 10, "百": 100, "佰": 100, "千": 1000, "仟": 1000, "万": 10000, "萬": 10000, "亿": 100000000 };
  const yuanIdx = s.search(/[圆元]/);
  const intPart = yuanIdx >= 0 ? s.slice(0, yuanIdx) : s;
  let jiao = 0, fen = 0;
  const jm = s.match(/([零〇一壹二贰两三四伍六柒捌玖])角/);
  const fm = s.match(/([零〇一壹二贰两三四伍六柒捌玖])分/);
  if (jm) jiao = digitMap[jm[1]];
  if (fm) fen = digitMap[fm[1]];
  let total = 0, section = 0, current = 0;
  for (const ch of intPart) {
    if (digitMap[ch] !== undefined) current = digitMap[ch];
    else if (unitMap[ch] !== undefined) {
      if (unitMap[ch] >= 10000) { section = (section + current) * unitMap[ch]; current = 0; }
      else { total += (current || 1) * unitMap[ch]; current = 0; }
    }
  }
  total += section + current;
  const val = total + jiao * 0.1 + fen * 0.01;
  return Math.round(val * 100) / 100;
}

// 从文本抽取标准字段（PDF/OFD/文本通用）
function pullFieldsFromText(text) {
  text = text.replace(/[\x00-\x08\x0b-\x1f\x7f]/g, " ");
  const result = {};
  let match;

  match = text.match(/发票代码[：:\s]*(\d{10,12})/);
  if (match) result.invoiceCode = match[1];

  match = text.match(/发票号码[：:\s]*(\d{8,20})/);
  if (match) result.invoiceNumber = match[1];

  match = text.match(/开票日期[：:\s]*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
  if (match) result.issueDate = `${match[1]}-${String(match[2]).padStart(2, "0")}-${String(match[3]).padStart(2, "0")}`;
  if (!result.issueDate) {
    match = text.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (match) result.issueDate = `${match[1]}-${String(match[2]).padStart(2, "0")}-${String(match[3]).padStart(2, "0")}`;
  }

  const entityNames = [...text.matchAll(/名称[：:\s]*([\s\S]*?)(?=统一社会|纳税人识别号|密码区|代码)/g)].map((x) => normalizeStr(x[1]));
  if (entityNames[0]) result.buyerName = entityNames[0];
  if (entityNames[1]) result.sellerName = entityNames[1];

  match = text.match(/购买方[\s\S]{0,120}?纳税人识别号[：:\s]*([A-Z0-9]{15,20})/);
  if (match) result.buyerTaxNo = match[1];
  match = text.match(/销售方[\s\S]{0,120}?纳税人识别号[：:\s]*([A-Z0-9]{15,20})/);
  if (match) result.sellerTaxNo = match[1];

  match = text.match(/价税合计[\s\S]{0,160}?小写[：:）)\s]*[¥￥]?\s*([\d,]+\.?\d*)/);
  if (match) result.total = match[1].replace(/,/g, "");

  match = text.match(/价税合计[\s\S]{0,160}?大写[\s\S]{0,40}?([零〇一壹二贰三叁四肆五伍六陆七柒八捌九玖十拾百佰千仟万萬亿圆元角整]+)/);
  if (!match) match = text.match(/价税[\s\S]{0,80}?大写[\s\S]{0,40}?([零〇一壹二贰三叁四肆五伍六陆七柒八捌九玖十拾百佰千仟万萬亿圆元角整]+)/);
  if (match && !result.total) {
    result.totalInWords = match[1];
    const num = uppercaseCnToNumber(result.totalInWords);
    if (num > 0) result.total = num.toFixed(2);
  }

  match = text.match(/金额[：:]?\s*(?:（小写）|\(小写\))?\s*[¥￥]?\s*([\d,]+\.\d{2})/);
  if (match) result.amount = match[1].replace(/,/g, "");

  match = text.match(/(?<!价)税额[：:]?\s*(?:（小写）|\(小写\))?\s*[¥￥]?\s*([\d,]+\.\d{2})/);
  if (match) result.tax = match[1].replace(/,/g, "");

  match = text.match(/校验码[：:\s]*([0-9A-Za-z]{15,32})/);
  if (match) result.checkCode = match[1];

  // 税率（形如 13%、9%、6%）
  match = text.match(/税率[：:]\s*(\d{1,2})%/);
  if (match) result.taxRate = match[1] + "%";

  return result;
}

// ---------------- XML 解析 ----------------
function firstDefined(obj, keys) {
  for (const k of keys) {
    if (obj && obj[k] !== undefined && obj[k] !== null && obj[k] !== "") return obj[k];
  }
  return undefined;
}
function extractValue(v) {
  if (v === undefined || v === null) return "";
  if (typeof v === "object") return v.__cdata !== undefined ? String(v.__cdata) : (v["#text"] !== undefined ? String(v["#text"]) : JSON.stringify(v));
  return String(v);
}
function parseInvoiceXml(xmlStr) {
  if (!XMLParser) return null; // 依赖缺失时降级
  const parser = new XMLParser({
    ignoreAttributes: false, attributeNamePrefix: "@_", trimValues: true,
    cdataPropName: "__cdata", parseTagValue: false,
  });
  const doc = parser.parse(xmlStr);
  const root = doc.fpkj || doc.RESPONSE || doc || {};
  return {
    invoiceCode: extractValue(firstDefined(root, ["fpdm", "发票代码", "invoiceCode"])),
    invoiceNumber: extractValue(firstDefined(root, ["fphm", "发票号码", "invoiceNumber"])),
    issueDate: extractValue(firstDefined(root, ["kprq", "开票日期", "issueDate"])).slice(0, 10),
    buyerName: extractValue(firstDefined(root, ["gfmc", "购买方名称", "buyerName"])),
    buyerTaxNo: extractValue(firstDefined(root, ["gfsh", "购买方税号", "buyerTaxNo"])),
    sellerName: extractValue(firstDefined(root, ["xfmc", "销售方名称", "sellerName"])),
    sellerTaxNo: extractValue(firstDefined(root, ["xfsh", "销售方税号", "sellerTaxNo"])),
    amount: extractValue(firstDefined(root, ["je", "金额", "amount"])),
    tax: extractValue(firstDefined(root, ["se", "税额", "tax"])),
    total: extractValue(firstDefined(root, ["jshj", "价税合计", "total"])),
    taxRate: extractValue(firstDefined(root, ["sl", "税率", "taxRate"])),
    checkCode: extractValue(firstDefined(root, ["jqbm", "校验码", "checkCode"])),
  };
}

// ---------------- OFD 解析 ----------------
async function extractOfd(buffer) {
  if (!JSZip) return null;
  const zip = await JSZip.loadAsync(buffer);
  const xmlPaths = Object.keys(zip.files).filter((p) => p.toLowerCase().endsWith(".xml"));
  const texts = await Promise.all(xmlPaths.map(async (p) => String(await zip.files[p].async("string"))));
  let text = texts.join("\n");
  const textCodes = [...text.matchAll(/<TextCode[^>]*>([\s\S]*?)<\/TextCode>/g)].map((m) => m[1]);
  if (textCodes.length) text = textCodes.join("\n");
  return pullFieldsFromText(text);
}

// ---------------- PDF 解析（零依赖） ----------------
function decompressBuffer(buf) {
  try { return zlib.inflateSync(buf); } catch (e) { try { return zlib.gunzipSync(buf); } catch (e2) { return null; } }
}
function decodePdfLiteral(s) {
  return s
    .replace(/\\([0-7]{1,3})/g, (m, o) => String.fromCharCode(parseInt(o, 8) & 0xff))
    .replace(/\\([()\\])/g, (m, c) => c)
    .replace(/\\r/g, "\r").replace(/\\n/g, "\n").replace(/\\t/g, "\t");
}
function decodeHexString(hex) {
  const clean = hex.replace(/\s+/g, "");
  if (clean.length % 2 !== 0) return "";
  const bytes = Buffer.from(clean, "hex");
  if (bytes.length >= 2) {
    let beOk = true;
    for (let i = 0; i < bytes.length; i += 2) if (bytes.readUInt16BE(i) === 0) { beOk = false; break; }
    if (beOk) return bytes.swap16().toString("utf16le");
  }
  return bytes.toString("latin1");
}
function buildCMap(text) {
  const charMap = new Map();
  const bc = text.match(/beginbfchar([\s\S]*?)endbfchar/);
  if (bc) {
    const lines = bc[1].match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g) || [];
    for (const ln of lines) {
      const mm = ln.match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/);
      charMap.set(parseInt(mm[1], 16), String.fromCodePoint(parseInt(mm[2], 16)));
    }
  }
  const br = text.match(/beginbfrange([\s\S]*?)endbfrange/);
  if (br) {
    const blocks = br[1].match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]+>|\[[\s\S]*?\])/g) || [];
    for (const blk of blocks) {
      const mm = blk.match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(.+)$/s);
      const s = parseInt(mm[1], 16), e = parseInt(mm[2], 16), third = mm[3].trim();
      if (third.startsWith("<")) {
        const u = parseInt(third.replace(/[<>]/g, ""), 16);
        for (let c = s; c <= e; c++) charMap.set(c, String.fromCodePoint(u + (c - s)));
      } else {
        const arr = (third.match(/<([0-9A-Fa-f]+)>/g) || []).map((h) => parseInt(h.replace(/[<>]/g, ""), 16));
        let i = 0; for (let c = s; c <= e; c++) charMap.set(c, String.fromCodePoint(arr[i++] ?? 0));
      }
    }
  }
  return charMap;
}
function getStreamContent(str, objNum) {
  const re = new RegExp(objNum + "\\s+\\d+\\s+obj([\\s\\S]*?)endobj", "g");
  let m;
  while ((m = re.exec(str))) {
    const body = m[1];
    const sIdx = body.indexOf("stream");
    if (sIdx < 0) continue;
    let s = body.slice(sIdx + 6).replace(/^\r?\n/, "");
    const eIdx = s.indexOf("endstream");
    if (eIdx >= 0) s = s.slice(0, eIdx);
    let buf = Buffer.from(s, "latin1");
    if (/FlateDecode/i.test(body)) { const inf = decompressBuffer(buf); if (inf) buf = inf; else return null; }
    return buf.toString("latin1");
  }
  return null;
}
function resolveFontToUnicode(str, fobj) {
  const re = new RegExp(fobj + "\\s+\\d+\\s+obj([\\s\\S]*?)endobj", "g");
  let m;
  while ((m = re.exec(str))) {
    const mm = m[1].match(/\/ToUnicode\s+(\d+)\s+(\d+)\s+R/);
    if (mm) return +mm[1];
  }
  return null;
}
function resolveToken(tok, cmap) {
  if (tok.startsWith("<")) {
    const clean = tok.slice(1, -1).replace(/\s+/g, "");
    if (clean.length % 2 !== 0) return "";
    const bytes = Buffer.from(clean, "hex");
    if (cmap) { let s = ""; for (let i = 0; i < bytes.length; i += 2) s += (cmap.get(bytes.readUInt16BE(i)) ?? "?"); return s; }
    return decodeHexString(tok);
  }
  return decodePdfLiteral(tok.slice(1, -1));
}
function extractTextWithFonts(txt, fontToUni, combined) {
  let out = "", curFont = null;
  const opRe = /(?:\/([A-Za-z0-9]+)\s+[\d.]+\s*Tf)|(?:(<[0-9A-Fa-f\s]*>|\((?:\\.|[^()\\])*\))\s*(Tj|'|"))|(?:\[([\s\S]*?)\]\s*TJ)/g;
  let m;
  while ((m = opRe.exec(txt))) {
    if (m[1]) { curFont = m[1]; continue; }
    const cmap = (curFont && fontToUni[curFont]) || combined || null;
    if (m[4] != null) {
      const sRe = /(<[0-9A-Fa-f\s]*>|\((?:\\.|[^()\\])*\))/g; let mm;
      while ((mm = sRe.exec(m[4]))) out += resolveToken(mm[1], cmap);
    } else if (m[2] != null) out += resolveToken(m[2], cmap);
  }
  return out;
}
function parsePdfText(buffer) {
  const data = Buffer.from(buffer);
  const str = data.toString("latin1");
  const toUniByObj = {};
  const tuRe = /\/ToUnicode\s+(\d+)\s+(\d+)\s+R/g; let tm;
  while ((tm = tuRe.exec(str))) {
    const onum = +tm[1];
    if (!toUniByObj[onum]) { const ct = getStreamContent(str, onum); if (ct != null) toUniByObj[onum] = buildCMap(ct); }
  }
  const fontToUni = {}, combined = new Map();
  for (const k in toUniByObj) for (const [cid, u] of toUniByObj[k]) combined.set(cid, u);
  const fontBlocks = str.match(/\/Font\s*<<([\s\S]*?)>>/g) || [];
  for (const blk of fontBlocks) {
    const refs = blk.match(/\/([A-Za-z0-9]+)\s+(\d+)\s+(\d+)\s+R/g) || [];
    for (const r of refs) {
      const mm = r.match(/\/([A-Za-z0-9]+)\s+(\d+)\s+(\d+)\s+R/);
      const cmapObj = resolveFontToUnicode(str, +mm[2]);
      if (cmapObj != null && toUniByObj[cmapObj]) fontToUni[mm[1]] = toUniByObj[cmapObj];
    }
  }
  const texts = [];
  const streamRe = /stream\r?\n([\s\S]*?)\r?\n?endstream/g; let m;
  while ((m = streamRe.exec(str))) {
    let sb = Buffer.from(m[1], "latin1");
    const dictStart = str.lastIndexOf("<<", m.index), dict = str.slice(dictStart, m.index);
    const arr = dict.match(/\/Filter\s*\[\s*([^\]]*)\s*\]/);
    let flist = arr ? arr[1].match(/\/([A-Za-z0-9]+)/g).map((s) => s.slice(1)) : [];
    if (!flist.length) { const single = dict.match(/\/Filter\s*\/([A-Za-z0-9]+)/); if (single) flist = [single[1]]; }
    let ok = true;
    for (const f of flist) {
      if (f === "FlateDecode" || f === "Fl") { const inf = decompressBuffer(sb); if (!inf) { ok = false; break; } sb = inf; }
      else { ok = false; break; }
    }
    if (!ok) continue;
    texts.push(extractTextWithFonts(sb.toString("latin1"), fontToUni, combined));
  }
  return texts.join("\n");
}

// ---------------- 统一输出（canonical 字段） ----------------
function toCanonicalRecord(file, type, fields, rawText) {
  const strVal = (v) => (v === undefined || v === null || v === "") ? "" : String(v).trim();
  return {
    source_file: path.basename(file),
    page: "",
    invoice_code: strVal(fields.invoiceCode),
    invoice_number: strVal(fields.invoiceNumber),
    invoice_date: strVal(fields.issueDate),
    invoice_type: strVal(fields.invoiceType),
    buyer_name: strVal(fields.buyerName),
    buyer_tax_id: strVal(fields.buyerTaxNo),
    seller_name: strVal(fields.sellerName),
    seller_tax_id: strVal(fields.sellerTaxNo),
    amount_ex_tax: strVal(fields.amount),
    tax_amount: strVal(fields.tax),
    total_amount: strVal(fields.total),
    tax_rate: strVal(fields.taxRate),
    check_code: strVal(fields.checkCode),
    remarks: strVal(fields.remarks),
    total_in_words: strVal(fields.totalInWords),
    reimbursement_person: "",
    expense_category: "",
    confidence: "",
    needs_vision: false,
    raw_text: rawText || "",
    error: "",
  };
}

function computeConfidence(r) {
  const keys = ["invoice_code", "invoice_number", "invoice_date", "buyer_name", "seller_name", "total_amount"];
  const found = keys.filter((k) => r[k]).length;
  return found >= 5 ? "高" : found >= 3 ? "中" : found >= 1 ? "低" : "无";
}

async function extractFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const buf = fs.readFileSync(filePath);
  let rec;
  if (ext === ".xml") {
    const raw = buf.toString("utf-8");
    const f = parseInvoiceXml(raw) || {};
    if (!f.invoiceNumber && !f.total) {
      // 结构化解析失败，转文本兜底
      const tf = pullFieldsFromText(raw);
      Object.assign(f, tf);
    }
    rec = toCanonicalRecord(filePath, "xml", f, "");
  } else if (ext === ".ofd") {
    const f = await extractOfd(buf);
    if (!f) { rec = toCanonicalRecord(filePath, "ofd", {}, ""); rec.needs_vision = true; rec.error = "缺少 jszip 依赖，OFD 无法解析"; }
    else rec = toCanonicalRecord(filePath, "ofd", f, "");
  } else if (ext === ".pdf") {
    const text = parsePdfText(buf);
    const f = pullFieldsFromText(text);
    rec = toCanonicalRecord(filePath, "pdf", f, text);
    if (!text || text.trim().length < 20) { rec.needs_vision = true; rec.confidence = "无"; rec.error = "PDF 无文本层，需多模态读图提取"; }
  } else if ([".png", ".jpg", ".jpeg", ".bmp", ".webp"].includes(ext)) {
    rec = toCanonicalRecord(filePath, "image", {}, "");
    rec.needs_vision = true; rec.confidence = "无"; rec.error = "图片文件，需多模态读图提取";
  } else {
    const text = buf.toString("utf-8");
    const f = pullFieldsFromText(text);
    rec = toCanonicalRecord(filePath, "text", f, text);
  }
  if (!rec.confidence) rec.confidence = computeConfidence(rec);
  return rec;
}

// ---------------- 主流程 ----------------
async function main() {
  const args = process.argv.slice(2);
  if (!args.length) { console.error("用法: node extract.js <文件或目录> [--json 提取结果.json]"); process.exit(1); }
  const target = args[0];
  const jsonOut = args.includes("--json") ? args[args.indexOf("--json") + 1] : null;
  let files = [];
  const INVOICE_EXT = /\.(pdf|ofd|xml|png|jpe?g|bmp|webp|txt)$/i;
  const SKIP_NAME = /^(结果|汇总|提取|\.task_ready|README|readme|instruction|Dockerfile)/i;
  if (fs.statSync(target).isDirectory()) {
    files = fs.readdirSync(target)
      .filter((f) => !SKIP_NAME.test(f) && INVOICE_EXT.test(f))
      .map((f) => path.join(target, f));
  } else files = [target];

  const results = [];
  let visionCount = 0;
  for (const f of files) {
    try {
      const r = await extractFile(f);
      results.push(r);
      if (r.needs_vision) visionCount++;
      console.log(`✔ ${r.source_file} [${r.page || r.source_file}] 置信度=${r.confidence}${r.needs_vision ? " (需读图)" : ""} 价税合计=${r.total_amount || "-"}`);
    } catch (e) {
      const rec = toCanonicalRecord(f, path.extname(f).slice(1), {}, "");
      rec.confidence = "无";
      rec.error = String(e.message || e);
      rec.needs_vision = true;
      results.push(rec);
      console.error(`✘ ${path.basename(f)} 异常: ${e.message}`);
    }
  }
  if (jsonOut) fs.writeFileSync(jsonOut, JSON.stringify(results, null, 2));
  console.log(`\n共处理 ${results.length} 个文件${visionCount ? `，其中 ${visionCount} 个需多模态读图提取` : ""}。` + (jsonOut ? ` 结果已写入 ${jsonOut}` : ""));
}

if (require.main === module) main().catch((e) => { console.error(e); process.exit(1); });
module.exports = { extractFile, parsePdfText, pullFieldsFromText };
