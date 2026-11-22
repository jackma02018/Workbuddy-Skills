/**
 * 发票整理聚合（batch-invoice-processor 技能）
 *
 * 输入：一张发票对象数组（canonical 字段，由 extract.js 产出，或由本平台多模态读图补全）。
 * 处理：费用归类、金额核验、去重、异常检测、按人员维度汇总。
 * 输出：
 *   - Excel 汇总表（3 个工作簿：发票明细 / 按人员统计 / 处理日志），依赖 jszip 时生成；
 *   - CSV 备用文件（始终生成，Excel 打不开时可用）；
 *   - Markdown 摘要文本（打印到控制台）。
 *
 * 用法：
 *   node organize.js --input 提取结果.json --out 发票汇总_2026-08-01.xlsx [--csv 同名.csv] [--md 同名.md]
 */

const fs = require("fs");
const path = require("path");
const os = require("os");

function loadDep(name) {
  try { return require(name); } catch (e) {}
  const searchPaths = [
    path.join(os.homedir(), ".workbuddy", "binaries", "node", "workspace", "node_modules"),
    path.join(__dirname, "..", "node_modules"),
    path.join(__dirname, "node_modules"),
  ];
  for (const base of searchPaths) { try { return require(path.join(base, name)); } catch (e) {} }
  return null;
}
const JSZip = loadDep("jszip");

// ---------------- 工具函数 ----------------
function parseNum(s) {
  if (s === undefined || s === null || s === "") return null;
  const t = String(s).replace(/[¥￥,\s]/g, "").replace(/[^0-9.\-]/g, "");
  if (t === "") return null;
  const n = parseFloat(t);
  return isNaN(n) ? null : n;
}
function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// ---------------- 费用归类 ----------------
const CATEGORY_RULES = [
  ["城际交通", /火车|高铁|动车|客运|汽车票|航空|机票|铁路|航班|行程单|通行费/],
  ["市内交通", /出租|打车|滴滴|轻轨|地铁|公交|网约车|共享单车/],
  ["住宿餐饮", /酒店|宾馆|住宿|餐饮|饭店|餐费|外卖|美食|食堂/],
  ["办公用品", /办公|文具|纸张|耗材|打印|复印|文化用品/],
  ["通信费", /通信|电信|移动|联通|话费|宽带|手机|网络/],
  ["物业水电", /物业|水电|燃气|暖气|供暖|房租|租赁/],
  ["维修保养", /维修|保养|检测|安装|配件|养护/],
  ["咨询服务", /咨询|服务|顾问|审计|评估|设计|策划|中介/],
  ["软件及IT", /软件|系统|IT|信息技术|云计算|数据|开发|算力/],
  ["广告宣传", /广告|宣传|推广|媒体|印刷|制作|营销/],
  ["培训教育", /培训|教育|课程|考试|学费|学校|学院/],
  ["会议展览", /会议|会展|论坛|展览|研讨/],
  ["劳务费", /劳务|人工|工资|薪酬|社保|公积金|兼职/],
  ["工程物资", /工程|建设|施工|材料|设备|建材|钢材|机械/],
];
function classifyExpense(rec) {
  if (rec.expense_category) return rec.expense_category;
  const text = `${rec.seller_name || ""} ${rec.remarks || ""} ${rec.invoice_type || ""}`;
  for (const [cat, re] of CATEGORY_RULES) if (re.test(text)) return cat;
  return "其他";
}

// ---------------- 金额核验 ----------------
const SKIP_CHECK_TYPES = /铁路电子客票|公路客票|通行费/;
function verifyAmount(rec) {
  if (SKIP_CHECK_TYPES.test(rec.invoice_type)) return "免校验";
  const a = parseNum(rec.amount_ex_tax), t = parseNum(rec.tax_amount), tot = parseNum(rec.total_amount);
  if (a == null || t == null || tot == null) return "字段缺失";
  return Math.abs(a + t - tot) <= 0.01 ? "一致" : "不一致";
}

// ---------------- 文件名金额核对 ----------------
function checkFileNameAmount(rec) {
  const m = (rec.source_file || "").match(/(\d+(?:\.\d{1,2})?)\s*元/);
  if (!m) return "无金额可核对";
  const fn = parseFloat(m[1]);
  const tot = parseNum(rec.total_amount);
  if (tot == null) return "无金额可核对";
  return Math.abs(fn - tot) <= 0.01 ? "一致" : `不一致(文件${fn}/实${tot})`;
}

// ---------------- 主处理 ----------------
// 仅接受发票类源文件；跳过评测遥测 / README / 说明等非票文件
const INVOICE_SOURCE_RE = /\.(pdf|png|jpe?g|webp|gif|txt|csv|xlsx?)$/i;
const INVOICE_SOURCE_SKIP_RE = /(?:^|[\\/])(?:\.task_ready|README|readme|instruction|Dockerfile)(?:$|\.)/i;

function isInvoiceSourceFile(name) {
  if (!name || typeof name !== "string") return false;
  const base = path.basename(name);
  if (INVOICE_SOURCE_SKIP_RE.test(base)) return false;
  if (/\.md$/i.test(base) && !/发票|invoice/i.test(base)) return false;
  return INVOICE_SOURCE_RE.test(base);
}

function processRecords(records) {
  const dedupMap = new Map();      // 去重键 -> 第一次出现的序号
  const sellerFreq = {};
  const codeNumSets = {};           // 同发票代码下号码集合，用于连号检测
  const out = [];
  let idx = 0;
  for (const r of records) {
    if (!isInvoiceSourceFile(r.source_file || "")) {
      console.warn(`跳过非发票源文件: ${r.source_file || "(空)"}`);
      continue;
    }
    idx++;
    const rec = Object.assign({}, r);
    rec.expense_category = classifyExpense(rec);
    rec.reimbursement_month = rec.invoice_date ? rec.invoice_date.slice(0, 7) : "";

    const a = parseNum(rec.amount_ex_tax), t = parseNum(rec.tax_amount), tot = parseNum(rec.total_amount);
    const incomplete = !rec.invoice_number && !rec.total_amount && rec.needs_vision;

    // 去重
    let dupKey = null;
    if (rec.invoice_code && rec.invoice_number) dupKey = rec.invoice_code + "|" + rec.invoice_number;
    else if (rec.invoice_number && String(rec.invoice_number).length >= 20) dupKey = "N|" + rec.invoice_number;
    let isDup = false;
    if (dupKey) {
      if (dedupMap.has(dupKey)) { isDup = true; rec._dupOf = dedupMap.get(dupKey); }
      else dedupMap.set(dupKey, idx);
    }

    // 销售方频次
    if (rec.seller_name) sellerFreq[rec.seller_name] = (sellerFreq[rec.seller_name] || 0) + 1;
    // 连号
    if (rec.invoice_code && rec.invoice_number && /^\d+$/.test(rec.invoice_number)) {
      (codeNumSets[rec.invoice_code] = codeNumSets[rec.invoice_code] || []).push(parseInt(rec.invoice_number, 10));
    }

    // 异常
    const issues = [];
    if (incomplete) issues.push("需读图提取(未补全)");
    else {
      const ac = verifyAmount(rec);
      if (ac === "不一致") issues.push("金额不一致");
      if (ac === "字段缺失") issues.push("金额字段缺失");
      if (isDup) issues.push(`疑似重复(与#${rec._dupOf})`);
      if (rec.tax_rate && rec.tax_rate.replace("%", "") === "0" && !/免税/.test(rec.invoice_type)) issues.push("税率为0");
      if (tot != null && tot > 100000) issues.push("金额异常大");
      if (rec.invoice_date) {
        const d = new Date(rec.invoice_date); const now = new Date();
        const oneYearAgo = new Date(); oneYearAgo.setFullYear(now.getFullYear() - 1);
        if (d > now) issues.push("日期在未来");
        else if (d < oneYearAgo) issues.push("日期超1年");
      }
    }
    if (rec.seller_name && sellerFreq[rec.seller_name] >= 5) issues.push("同销售方≥5张");

    rec._is_duplicate = isDup ? "是" : "否";
    rec._amount_check = incomplete ? "—" : verifyAmount(rec);
    rec._name_check = incomplete ? "—" : checkFileNameAmount(rec);
    rec._anomaly = issues.join("；");
    rec._incomplete = incomplete;
    out.push(rec);
  }

  // 连号检测（同一代码下号码连续≥3）
  for (const code in codeNumSets) {
    const nums = codeNumSets[code].slice().sort((a, b) => a - b);
    let run = 1;
    for (let i = 1; i < nums.length; i++) {
      if (nums[i] === nums[i - 1] + 1) run++; else run = 1;
      if (run >= 3) {
        for (const rec of out) {
          if (rec.invoice_code === code && parseInt(rec.invoice_number, 10) >= nums[i] - run + 1 && parseInt(rec.invoice_number, 10) <= nums[i]) {
            if (!rec._anomaly.includes("连号发票")) rec._anomaly = (rec._anomaly ? rec._anomaly + "；" : "") + "连号发票";
          }
        }
      }
    }
  }
  return out;
}

// ---------------- 按人员统计 ----------------
function aggregateByPerson(records) {
  const map = {};
  for (const r of records) {
    const p = r.reimbursement_person || "（未填报销人）";
    const g = map[p] || (map[p] = { person: p, count: 0, amount: 0, tax: 0, total: 0 });
    g.count++;
    g.amount += parseNum(r.amount_ex_tax) || 0;
    g.tax += parseNum(r.tax_amount) || 0;
    g.total += parseNum(r.total_amount) || 0;
  }
  const rows = Object.values(map).map((g) => ({
    person: g.person, count: g.count,
    amount: +g.amount.toFixed(2), tax: +g.tax.toFixed(2), total: +g.total.toFixed(2),
  }));
  const sum = rows.reduce((a, r) => ({ count: a.count + r.count, amount: a.amount + r.amount, tax: a.tax + r.tax, total: a.total + r.total }), { count: 0, amount: 0, tax: 0, total: 0 });
  return { rows, sum: { count: sum.count, amount: +sum.amount.toFixed(2), tax: +sum.tax.toFixed(2), total: +sum.total.toFixed(2) } };
}

// ---------------- CSV ----------------
function buildCsv(records) {
  const cols = ["序号", "报销人", "发票代码", "发票号码", "开票日期", "发票类型", "购买方名称", "购买方税号", "销售方名称", "销售方税号", "费用类型", "金额(不含税)", "税率", "税额", "价税合计", "所属部门", "所属项目", "报销月份", "是否重复", "金额校验", "文件名金额核对", "异常标记", "备注", "校验码", "文件名"];
  const esc = (v) => { const s = v == null ? "" : String(v); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const lines = [cols.join(",")];
  records.forEach((r, i) => {
    lines.push([
      i + 1, r.reimbursement_person, r.invoice_code, r.invoice_number, r.invoice_date, r.invoice_type,
      r.buyer_name, r.buyer_tax_id, r.seller_name, r.seller_tax_id, r.expense_category,
      r.amount_ex_tax, r.tax_rate, r.tax_amount, r.total_amount, r.department, r.project, r.reimbursement_month,
      r._is_duplicate, r._amount_check, r._name_check, r._anomaly, r.remarks, r.check_code, r.source_file,
    ].map(esc).join(","));
  });
  return lines.join("\r\n");
}

// ---------------- Markdown 汇总 ----------------
function buildMarkdown(records, stats) {
  const persons = stats.rows;
  let total = 0;
  records.forEach((r) => { const t = parseNum(r.total_amount); if (t) total += t; });
  total = +total.toFixed(2);
  let md = `# 发票整理汇总\n\n`;
  md += `**共 ${records.length} 张发票，价税合计 ¥${total.toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}**\n\n`;
  md += `## 按人员统计\n\n| 报销人 | 张数 | 价税合计 |\n| --- | --- | --- |\n`;
  persons.forEach((p) => { md += `| ${p.person} | ${p.count} | ¥${p.total.toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} |\n`; });
  md += `| **合计** | **${stats.sum.count}** | **¥${stats.sum.total.toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}** |\n\n`;

  const dups = records.filter((r) => r._is_duplicate === "是");
  const anoms = records.filter((r) => r._anomaly && !r._incomplete);
  const vision = records.filter((r) => r._incomplete);
  md += `## 处理情况\n`;
  if (dups.length) md += `- 重复发票 ${dups.length} 张：${dups.map((r) => `#${records.indexOf(r) + 1} ${r.invoice_number}`).join("、")}\n`;
  if (anoms.length) md += `- 异常 ${anoms.length} 张：${anoms.map((r) => `#${records.indexOf(r) + 1}(${r._anomaly})`).join("；")}\n`;
  if (vision.length) md += `- 需多模态读图提取 ${vision.length} 张（未补全，建议用 Read 工具读取后补全字段）\n`;
  if (!dups.length && !anoms.length && !vision.length) md += `- 无重复、无异常、无需读图补全。\n`;
  return md;
}

// ---------------- Excel（jszip） ----------------
function colLetter(n) { let s = ""; while (n >= 0) { s = String.fromCharCode((n % 26) + 65) + s; n = Math.floor(n / 26) - 1; } return s; }
function xmlEscape(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }

function buildXlsx(sheets) {
  // sheets: [{name, cols:[{header, num?}], rows:[ [cellValue,...] ]}]
  const sheetXml = [];
  const sheetRels = [];
  sheets.forEach((sh, si) => {
    let body = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>`;
    // header
    let row = `<row r="1">`;
    sh.cols.forEach((c, ci) => { row += `<c r="${colLetter(ci)}1" s="1" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(c.header)}</t></is></c>`; });
    row += `</row>`;
    body += row;
    sh.rows.forEach((r, ri) => {
      const rn = ri + 2;
      let rr = `<row r="${rn}">`;
      r.forEach((cell, ci) => {
        const isNum = sh.cols[ci] && sh.cols[ci].num && cell !== "" && cell != null;
        const style = isNum ? 2 : 0;
        const ref = `${colLetter(ci)}${rn}`;
        if (isNum) rr += `<c r="${ref}" s="${style}"><v>${cell}</v></c>`;
        else if (cell === "" || cell == null) rr += `<c r="${ref}" s="0"/>`;
        else rr += `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(cell)}</t></is></c>`;
      });
      rr += `</row>`;
      body += rr;
    });
    body += `</sheetData></worksheet>`;
    sheetXml.push(body);
    sheetRels.push({ id: si + 1, name: sh.name });
  });

  const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${sheetRels.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("\n")}
</Types>`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
${sheetRels.map((s, i) => `<sheet name="${xmlEscape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("\n")}
</sheets>
</workbook>`;

  const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheetRels.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("\n")}
<Relationship Id="rId${sheetRels.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
<fonts count="3">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
<font><color rgb="FFFF0000"/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFC6E0B4"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", ct);
  zip.folder("_rels").file(".rels", rootRels);
  const xl = zip.folder("xl");
  xl.file("workbook.xml", wb);
  xl.folder("_rels").file("workbook.xml.rels", wbRels);
  xl.file("styles.xml", styles);
  sheetXml.forEach((sx, i) => { xl.folder("worksheets").file(`sheet${i + 1}.xml`, sx); });
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ---------------- 主流程 ----------------
async function main() {
  const args = process.argv.slice(2);
  let input = null, outXlsx = null, outCsv = null, outMd = null;
  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--input") input = args[++i];
    else if (args[i] === "--out") outXlsx = args[++i];
    else if (args[i] === "--csv") outCsv = args[++i];
    else if (args[i] === "--md") outMd = args[++i];
  }
  if (!input) { console.error("用法: node organize.js --input 提取结果.json --out 发票汇总.xlsx [--csv x.csv] [--md x.md]"); process.exit(1); }

  const records = JSON.parse(fs.readFileSync(input, "utf-8"));
  if (!Array.isArray(records)) {
    console.error("输入 JSON 须为发票对象数组");
    process.exit(1);
  }
  const processed = processRecords(records);
  const stats = aggregateByPerson(processed);

  // 处理日志
  const logRows = processed.map((r, i) => ({
    file: r.source_file, page: r.page || "1",
    page_type: r.needs_vision ? "invoice_image" : "invoice_text",
    strategy: r.needs_vision ? "vision_ocr" : "text_extract",
    result: r._incomplete ? "needs_vision" : (r.error ? "failed" : "success"),
    note: r._incomplete ? "需多模态读图提取" : (r._anomaly || (r._amount_check === "一致" ? "正常" : r._amount_check)),
  }));

  const summaryMd = buildMarkdown(processed, stats);
  console.log(summaryMd);

  // CSV（始终生成）
  const csvPath = outCsv || (outXlsx ? outXlsx.replace(/\.xlsx$/i, ".csv") : input.replace(/\.json$/i, ".csv"));
  fs.writeFileSync(csvPath, "\uFEFF" + buildCsv(processed));
  console.log(`CSV 已写入 ${csvPath}`);

  if (outMd) fs.writeFileSync(outMd, summaryMd);

  // Excel
  if (outXlsx) {
    if (!JSZip) {
      console.error("⚠️ 未找到 jszip，无法生成 Excel；已输出 CSV 备用文件（可用 Excel 直接打开）。");
    } else {
      const detailCols = [
        { header: "序号" }, { header: "报销人" }, { header: "发票代码" }, { header: "发票号码" }, { header: "开票日期" },
        { header: "发票类型" }, { header: "购买方名称" }, { header: "购买方税号" }, { header: "销售方名称" }, { header: "销售方税号" },
        { header: "费用类型" }, { header: "金额(不含税)", num: true }, { header: "税率" }, { header: "税额", num: true }, { header: "价税合计", num: true },
        { header: "所属部门" }, { header: "所属项目" }, { header: "报销月份" }, { header: "是否重复" }, { header: "金额校验" },
        { header: "文件名金额核对" }, { header: "异常标记" }, { header: "备注" }, { header: "校验码" }, { header: "文件名" },
      ];
      const detailRows = processed.map((r, i) => [
        i + 1, r.reimbursement_person, r.invoice_code, r.invoice_number, r.invoice_date, r.invoice_type,
        r.buyer_name, r.buyer_tax_id, r.seller_name, r.seller_tax_id, r.expense_category,
        parseNum(r.amount_ex_tax) ?? "", r.tax_rate, parseNum(r.tax_amount) ?? "", parseNum(r.total_amount) ?? "",
        r.department || "", r.project || "", r.reimbursement_month, r._is_duplicate, r._amount_check,
        r._name_check, r._anomaly, r.remarks, r.check_code, r.source_file,
      ]);
      const personCols = [{ header: "报销人" }, { header: "发票张数", num: true }, { header: "不含税金额合计", num: true }, { header: "税额合计", num: true }, { header: "价税合计", num: true }];
      const personRows = stats.rows.map((p) => [p.person, p.count, p.amount, p.tax, p.total]);
      personRows.push(["合计", stats.sum.count, stats.sum.amount, stats.sum.tax, stats.sum.total]);
      const logCols = [{ header: "文件名" }, { header: "页码" }, { header: "页面类型" }, { header: "提取策略" }, { header: "处理结果" }, { header: "备注" }];
      const logDataRows = logRows.map((l) => [l.file, l.page, l.page_type, l.strategy, l.result, l.note]);

      const buf = await buildXlsx([
        { name: "发票明细", cols: detailCols, rows: detailRows },
        { name: "按人员统计", cols: personCols, rows: personRows },
        { name: "处理日志", cols: logCols, rows: logDataRows },
      ]);
      fs.writeFileSync(outXlsx, buf);
      console.log(`Excel 已写入 ${outXlsx}`);
    }
  }
}

if (require.main === module) main().catch((e) => { console.error(e); process.exit(1); });
module.exports = {
  processRecords,
  process: processRecords, // 兼容旧名；勿在本文件内再定义 function process（会遮蔽 Node 全局 process）
  aggregateByPerson,
  classifyExpense,
  isInvoiceSourceFile,
};
