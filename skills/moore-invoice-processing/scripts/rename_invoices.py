# -*- coding: utf-8 -*-
"""
发票批量重命名工具
==================
命名规范：日期_票据类型_补充信息_金额.扩展名

- 日期   : 票面识别，YYYYMMDD；多日票据用 YYYYMMDD-YYYYMMDD
- 类型   : 火车票 / 退票 / 网约车发票 / 网约车行程
- 补充   : 火车票=发站-到站；网约车=省+市（多城市用 + 连接）
- 金额   : 两位小数，无货币符号

用法:
    python rename_invoices.py --root <目录> [--dry-run] [--execute]
"""
import argparse
import json
import os
import re
import sys
from collections import defaultdict

try:
    import pymupdf as fitz
except ImportError:
    import fitz

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

SCAN_EXT = {".pdf", ".ofd", ".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}

_OCR = None


def _get_ocr():
    """延迟加载 OCR 引擎。首次约 5s 下载/加载模型，之后复用。"""
    global _OCR
    if _OCR is None:
        from rapidocr_onnxruntime import RapidOCR
        _OCR = RapidOCR()
    return _OCR


def _ocr_words(img_path):
    """OCR -> [(x0, x1, y, text)]，坐标做到与 PDF words 同构，后续逻辑可直接复用。"""
    res, _ = _get_ocr()(img_path)
    words = []
    for item in (res or []):
        box, txt = item[0], item[1]
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        words.append((min(xs), max(xs), round(min(ys), 1), txt.strip()))
    words.sort(key=lambda w: (w[2], w[0]))
    return _enhance_station_row(img_path, words)


def _enhance_station_row(img_path, words, scale=2):
    """全页 OCR 对票面大号加粗的站名常常丢掉末尾的「站」字（实测 dpi 300/400/600 均如此）。

    解决办法：定位到车次所在的横条ROI，裁出来放大后再识别一次 —— 实测这一步能把
    「苏州南 / 福州」还原成完整的「苏州南站 G1315 福州站」。识别不到就原样返回，不臆造。
    """
    hits = [i for i, w in enumerate(words) if re.fullmatch(r"[A-Z]\d{1,4}", w[3])]
    if not hits:
        return words
    try:
        from PIL import Image
    except Exception:
        return words
    import tempfile
    try:
        im = Image.open(img_path)
        W, H = im.size
        idx = hits[0]
        x0, x1, y_top = words[idx][0], words[idx][1], words[idx][2]
        band = max(20, int(H * 0.02))
        cx0 = max(0, int(x0 - W * 0.45))
        cx1 = min(W, int(x1 + W * 0.45))
        cy0 = max(0, int(y_top - band))
        cy1 = min(H, int(y_top + band * 2))
        crop = im.crop((cx0, cy0, cx1, cy1))
        if crop.width < 10 or crop.height < 10:
            return words
        crop = crop.resize((crop.width * scale, crop.height * scale), Image.LANCZOS)
        tmp = os.path.join(tempfile.gettempdir(), "_invoice_enhance.png")
        crop.save(tmp)
        try:
            res, _ = _get_ocr()(tmp)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)
        if not res:
            return words
        new_words = []
        for i, w in enumerate(words):
            if i != idx:
                new_words.append(w)
                continue
            for box, txt, _score in res:
                xs = [p[0] for p in box]
                ys = [p[1] for p in box]
                new_words.append((min(xs) / scale + cx0, max(xs) / scale + cx0,
                                  round(min(ys) / scale + cy0, 1), txt.strip()))
        new_words.sort(key=lambda w: (w[2], w[0]))
        return new_words
    except Exception:
        return words

ILLEGAL = re.compile(r'[\\/:*?"<>|\r\n\t]')
DATE_CN = re.compile(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
DATE_ISO = re.compile(r"(20\d{2})-(\d{1,2})-(\d{1,2})")
MONEY = re.compile(r"[¥￥]\s*(-?\d[\d,]*\.\d{2})")
MONEY_BARE = re.compile(r"(?<![\d.])(\d{1,6}\.\d{2})(?![\d.])")
AMOUNT_CN_TOTAL = re.compile(r"合计\s*([\d.]+)\s*元")
# 车次可能夹在两站中间（OCR 整行识别）或跟在后面（PDF 文本层），两种都支持
TRAIN_STATION = re.compile(
    r"([\u4e00-\u9fa5]{1,6}?站)\s*([A-Z]\d{1,4})?\s*([\u4e00-\u9fa5]{1,6}?站)\s*([A-Z]\d{1,4})?"
)

# 城市 -> 省份（按实际业务补齐；未命中的交给 cpca 兜底）
CITY_PROVINCE = {
    "苏州市": "江苏省", "南京市": "江苏省", "无锡市": "江苏省", "常州市": "江苏省",
    "徐州市": "江苏省", "南通市": "江苏省", "扬州市": "江苏省", "镇江市": "江苏省",
    "泰州市": "江苏省", "盐城市": "江苏省", "淮安市": "江苏省", "连云港市": "江苏省",
    "宿迁市": "江苏省",
    "福州市": "福建省", "厦门市": "福建省", "泉州市": "福建省", "漳州市": "福建省",
    "莆田市": "福建省", "三明市": "福建省", "南平市": "福建省", "龙岩市": "福建省",
    "宁德市": "福建省",
    "杭州市": "浙江省", "宁波市": "浙江省", "温州市": "浙江省", "嘉兴市": "浙江省",
    "湖州市": "浙江省", "绍兴市": "浙江省", "金华市": "浙江省", "台州市": "浙江省",
    "上海市": "上海市", "北京市": "北京市", "天津市": "天津市", "重庆市": "重庆市",
    "深圳市": "广东省", "广州市": "广东省", "东莞市": "广东省", "佛山市": "广东省",
}
_DIRECT_CITY = {"北京": "北京市", "上海": "上海市", "天津": "天津市", "重庆": "重庆市"}

# 命名模板：同类字段在不同口径下的显示名。
#   standard    通用报销口径，默认
#   railway     年月日_火车发票_出发站_到达站_金额
#   legacy12306 出发站-到达站-日期-金额（兼容 railway-invoice 老格式，仅火车票生效）
TEMPLATES = {
    "standard": {"train": "火车票", "refund": "退票",
                 "ride_invoice": "网约车发票", "trip_sheet": "网约车行程"},
    "railway":  {"train": "火车发票", "refund": "退票",
                 "ride_invoice": "网约车发票", "trip_sheet": "网约车行程"},
    "legacy12306": {"train": "火车票", "refund": "退票",
                    "ride_invoice": "网约车发票", "trip_sheet": "网约车行程"},
}


def read_pages(path):
    """返回 (pages, plain)。pages = [(页码, [(x0, x1, y, text), ...]), ...]

    PDF 走文本层；图片直接 OCR；PDF 若无文本层（扫描件）自动降级为 OCR。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        doc = fitz.open(path)
        pages = []
        for pno, page in enumerate(doc):
            words = page.get_text("words")
            words.sort(key=lambda w: (round(w[1], 1), round(w[0], 1)))
            pages.append((pno, [(w[0], w[2], round(w[1], 1), w[4]) for w in words]))
        plain = "\n".join(p.get_text("text") for p in doc)
        if len(re.sub(r"\s+", "", plain)) < 20:          # 扫描件：没有可用文本层
            pages = _pdf_pages_via_ocr(doc)
            plain = "\n".join(t for _, ws in pages for *_, t in ws)
        doc.close()
        return pages, plain
    if ext in IMAGE_EXT:
        words = _ocr_words(path)
        return [(0, words)], "\n".join(t for *_, t in words)
    return [], ""


def _pdf_pages_via_ocr(doc, dpi=300):
    """把每一页渲染成位图再 OCR，输出与文本层同构的 pages。"""
    import tempfile
    pages = []
    tmpdir = tempfile.mkdtemp(prefix="invoice_ocr_")
    try:
        for pno, page in enumerate(doc):
            tmp = os.path.join(tmpdir, "p%d.png" % pno)
            page.get_pixmap(dpi=dpi).save(tmp)
            try:
                pages.append((pno, _ocr_words(tmp)))
            finally:
                if os.path.exists(tmp):
                    os.remove(tmp)
    finally:
        try:
            os.rmdir(tmpdir)
        except Exception:
            pass
    return pages


def lines_of(pages, tol=3.5):
    """按 y 聚合成行。

    PDF 的 y 是 pt（票据页通常几百），OCR 出来的是像素坐标（数千），
    同一行的词间距被等比放大，所以行阈值要随坐标尺度自适应，否则一行会被劈成多行。
    """
    ys = [w[2] for _, ws in pages for w in ws]
    if ys and max(ys) > 1000:
        tol = max(tol, max(ys) / 80)
    out = []
    for pno, words in pages:
        cur_y, buf = None, []
        for x0, x1, y, t in words:
            if cur_y is None or abs(y - cur_y) > tol:
                if buf:
                    out.append((pno, cur_y, buf))
                buf, cur_y = [], y
            buf.append((x0, t))
        if buf:
            out.append((pno, cur_y, buf))
    return out


def city_column(pages):
    """按「城市」表头的x坐标取整列，解决城市名被换行拆成「苏州」+「市」的问题。"""
    cities = []
    for pno, words in pages:
        hdr = [w for w in words if w[3] == "城市"]
        if not hdr:
            continue
        hx0, hx1, hy = hdr[0][0], hdr[0][1], hdr[0][2]
        cells = [(y, t) for x0, x1, y, t in words
                 if y > hy + 5 and (hx0 - 8) <= x0 <= (hx1 + 8)
                 and not re.match(r"^[\d.]+$", t)]
        cells.sort()
        cur_y, buf = None, []
        for y, t in cells:
            if cur_y is None or y - cur_y <= 9:      # 同一单元格内的换行
                buf.append(t)
            else:
                if buf:
                    cities.append("".join(buf))
                buf = [t]
            cur_y = y
        if buf:
            cities.append("".join(buf))
    return [c for c in cities if c]


def join(line):
    return "".join(t for _, t in line)


def norm_date(y, m, d):
    return "%04d%02d%02d" % (int(y), int(m), int(d))


def clean_text(s):
    """清理空格与文件名非法字符。"""
    s = ILLEGAL.sub("", s or "")
    return re.sub(r"\s+", "", s).strip()


def to_province_city(city):
    city = city.strip()
    if not city.endswith(("市", "区", "县")) and city in _DIRECT_CITY:
        city = _DIRECT_CITY[city]
    prov = CITY_PROVINCE.get(city)
    if prov:
        return prov + city
    try:
        import cpca
        df = cpca.transform([re.sub(r"(省|市|自治区|自治州|地区|盟)$", "", city)])
        row = df.iloc[0]
        p, c = str(row["省"]), str(row["市"])
        if p and p != "nan" and c and c != "nan":
            return p + c
    except Exception:
        pass
    return city


# ---------------------------------------------------------------- 火车票
def parse_train(path):
    """铁路电子客票：日期取乘车日期，站名用票面原始写法。"""
    pages, plain = read_pages(path)
    lines = lines_of(pages)
    is_refund = "退票费" in plain

    # 乘车日期：形如「2026年09月11日09:02开」，务必避开「开票日期」行的日期
    ride_date = None
    # OCR 会插空格并把车厢号夹在日期与时间中间（「2026年09月11日02车10D号09: 02开」），
    # 故允许日期与「HH:MM开」之间夹少量其它字符。
    ride_re = re.compile(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日[^开]{0,20}?\d{1,2}:\s*\d{2}开")
    m = ride_re.search(re.sub(r"\s+", "", plain))
    if m:
        ride_date = norm_date(*m.groups())
    if ride_date is None:
        for _, _, ln in lines:
            s = join(ln)
            if "开票日期" in s:
                continue
            m2 = DATE_CN.search(s)
            if m2 and re.search(r"\d{2}:\d{2}开", s):
                ride_date = norm_date(*m2.groups())
                break

    from_st = to_st = None
    for _, _, ln in lines:
        m = TRAIN_STATION.search(join(ln))
        if m:
            from_st, to_st = m.group(1), m.group(3)
            break

    amount = None
    for m in MONEY.finditer(plain):
        amount = float(m.group(1).replace(",", ""))
    return {
        "date": ride_date,
        "from_station": from_st,
        "to_station": to_st,
        "amount": amount,
        "is_refund": is_refund,
    }


# ---------------------------------------------------------------- 网约车行程单
def parse_trip_sheet(path):
    pages, plain = read_pages(path)
    lines = lines_of(pages)
    compact = re.sub(r"\s+", "", plain)

    start = end = None
    m = re.search(r"行程起止日期[：:]?\s*(20\d{2}-\d{1,2}-\d{1,2})\s*至\s*(20\d{2}-\d{1,2}-\d{1,2})", compact)
    if m:
        a = DATE_ISO.search(m.group(1))
        b = DATE_ISO.search(m.group(2))
        start, end = norm_date(*a.groups()), norm_date(*b.groups())

    amount = None
    m2 = AMOUNT_CN_TOTAL.search(compact)
    if m2:
        amount = round(float(m2.group(1)), 2)

    cities = city_column(pages)
    platform = "曹操" if ("曹操" in plain or "优行" in plain) else ("滴滴" if "滴滴" in plain else "")
    return {"date": start, "date_end": end, "amount": amount,
            "cities": cities, "platform": platform}


def _looks_like_city_column(line, idx):
    """行程单「城市」列紧跟上车时间之后，排除起点/终点地名里的随机市字。"""
    toks = [t for _, t in line]
    for j in range(idx):
        if re.match(r"^\d{2}-\d{2}$", toks[j]):
            return True
    return False


# ---------------------------------------------------------------- 网约车发票
def parse_ride_invoice(path):
    pages, plain = read_pages(path)
    compact = re.sub(r"\s+", "", plain)

    amount = None
    m = re.search(r"（小写）\s*[¥￥]?\s*(\d+\.\d{2})", compact)
    if m:
        amount = round(float(m.group(1)), 2)
    if amount is None:
        hits = MONEY.findall(plain)
        if hits:
            amount = round(float(hits[-1].replace(",", "")), 2)

    travel = sorted({norm_date(*x.groups()) for x in DATE_ISO.finditer(compact)})

    # 开票日期：三重兜底链的最后一级。
    # 注意：PDF 内容流里「开票日期：」与其后的日期常被拆散，必须从坐标重组的行里取。
    invoice_date = None
    for _, _, ln in lines_of(pages):
        s = join(ln)
        if "开票日期" in s:
            m3 = DATE_CN.search(s)
            if m3:
                invoice_date = norm_date(*m3.groups())
                break
    if invoice_date is None:
        m3 = re.search(r"开票日期[：:]\s*(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", compact)
        if m3:
            invoice_date = norm_date(*m3.groups())

    seller = ""
    for m2 in re.finditer(r"([\u4e00-\u9fa5（）()]{4,40}?(?:科技有限公司|有限公司))", compact):
        seller = m2.group(1).strip()
    platform = ("曹操" if ("曹操" in compact or "吉利优行" in compact or "优行" in compact)
                else ("滴滴" if "滴滴" in compact else ""))

    # 城市兜底：发票票面无「城市」列时，退回销方注册地所在市
    seller_city = ""
    m4 = re.search(r"([\u4e00-\u9fa5]{2,8}?市)", seller or "")
    if m4:
        seller_city = m4.group(1)
    return {"amount": amount, "travel_dates": travel, "seller": seller,
            "platform": platform, "invoice_date": invoice_date, "seller_city": seller_city}


# ---------------------------------------------------------------- 主流程
def classify(path):
    _, plain = read_pages(path)
    t = re.sub(r"\s+", "", plain)
    if "行程起止日期" in t or ("行程单" in t and ("笔行程" in t or "行程明细" in t)):
        return "网约车行程"
    # OCR 常把「电子发票（铁路电子客票）」切成碎片，改用特征词投票而非整串匹配
    rail = sum(k in t for k in ("铁路电子客票", "铁路电子", "电子客票号", "12306",
                                "车次", "二等座", "一等座", "票价"))
    if rail >= 2:
        return "火车票"
    if "开票日期" in t and ("交通运输服务" in t or "客运服务费" in t or "价税合计" in t):
        return "网约车发票"
    if rail >= 1:
        return "火车票"
    return ""


def build_plan(root, template="standard", station_sep="-"):
    TPL = TEMPLATES.get(template, TEMPLATES["standard"])
    targets = []
    for dp, dn, fn in os.walk(root):
        rel = os.path.relpath(dp, root)
        # 注意 rel 为 "." 时不能直接 split 判断，否则 "." 本身以点开头会把根目录文件全漏掉
        parts = [] if rel == "." else rel.split(os.sep)
        if any(p.startswith(("_", ".")) for p in parts):
            continue
        for f in fn:
            if os.path.splitext(f)[1].lower() in SCAN_EXT:
                targets.append(os.path.join(dp, f))
    targets.sort()

    raw = defaultdict(dict)
    for p in targets:
        try:
            raw[p]["type"] = classify(p)
        except Exception as e:
            raw[p] = {"type": "", "error": "解析失败: %s" % e}

    trip_index = {}
    for p, info in raw.items():
        if info.get("type") != "网约车行程" or info.get("error"):
            continue
        d = parse_trip_sheet(p)
        info.update(d)
        key = (round(d["amount"], 2) if d["amount"] else None, d["platform"])
        trip_index[key] = d

    results = []
    for p in targets:
        base = os.path.basename(p)
        ext = os.path.splitext(p)[1].lower()
        info = raw[p]
        rec = {"path": p, "old": base, "ext": ext, "type": info.get("type", ""), "skip_reason": ""}
        if info.get("error"):
            rec["skip_reason"] = info["error"]
            results.append(rec)
            continue
        if not rec["type"]:
            rec["skip_reason"] = "无法识别票面类型"
            results.append(rec)
            continue

        try:
            if rec["type"] == "火车票":
                d = parse_train(p)
                date = d["date"]
                cat = TPL["refund" if d["is_refund"] else "train"]
                extra = station_sep.join([clean_text(d["from_station"]), clean_text(d["to_station"])]) \
                    if d["from_station"] and d["to_station"] else ""
                rec["stations"] = [clean_text(d["from_station"]), clean_text(d["to_station"])]
                amount = d["amount"]
                if not date:
                    rec["skip_reason"] = "未识别到乘车日期"
                elif not extra or extra == "-":
                    rec["skip_reason"] = "未识别到发站/到站"
                elif amount is None:
                    rec["skip_reason"] = "未识别到金额"
            else:
                if rec["type"] == "网约车行程":
                    d = parse_trip_sheet(p)
                    cat = TPL["trip_sheet"]
                    amount = d["amount"]
                    trips = d["cities"]
                    start, end = d["date"], d["date_end"]
                else:
                    d = parse_ride_invoice(p)
                    cat = TPL["ride_invoice"]
                    amount = d["amount"]
                    start = end = None          # 必须显式初始化，否则会串用上一轮的值
                    trips = []
                    ref = trip_index.get((round(amount, 2) if amount else None, d["platform"]))
                    if ref:
                        start = ref["date"] or d["invoice_date"]
                        end = ref["date_end"] or ref["date"]
                        trips = ref["cities"]
                    else:
                        trips = []
                    if not start:                       # 兜底：出行明细日期 -> 开票日期
                        start = d["travel_dates"][0] if d["travel_dates"] else d["invoice_date"]
                        end = d["travel_dates"][-1] if d["travel_dates"] else d["invoice_date"]
                    if not trips:                       # 兜底：销方注册地所在市
                        trips = [d["seller_city"]] if d["seller_city"] else []
                date = start if start == end or not end else "%s-%s" % (start, end)
                raw_cities = list(dict.fromkeys(trips))
                if raw_cities:
                    provs, resolved = [], []
                    for c in raw_cities:
                        pc = to_province_city(c)
                        prov = pc[: -len(c)] if pc.endswith(c) and len(pc) > len(c) else ""
                        provs.append(prov)
                        resolved.append(clean_text(c))
                    if len(set(provs)) == 1 and provs[0]:        # 同省只标一次省份
                        extra = clean_text(provs[0]) + "+".join(resolved)
                    else:
                        extra = "+".join(clean_text(to_province_city(c)) for c in raw_cities)
                else:
                    extra = ""
                if not date:
                    rec["skip_reason"] = "未识别到日期"
                elif not extra:
                    rec["skip_reason"] = "未识别到消费城市"
                elif amount is None:
                    rec["skip_reason"] = "未识别到金额"
        except Exception as e:
            rec["skip_reason"] = "解析异常: %s" % e
            results.append(rec)
            continue

        if rec["skip_reason"]:
            results.append(rec)
            continue

        amount_s = "%.2f" % amount
        if template == "legacy12306" and rec["type"] == "火车票" and rec.get("stations"):
            fields = ["-".join(rec["stations"]), date, amount_s]
        else:
            fields = [date, cat]
            if extra:
                fields.append(extra)
            fields.append(amount_s)
        new_name = clean_text("_".join(fields)) + rec["ext"]
        rec["new"] = new_name
        results.append(rec)
    return results


def resolve_conflicts(results):
    used = defaultdict(int)
    for r in results:
        if "new" not in r:
            continue
        n = r["new"]
        used[n] += 1
        if used[n] > 1:
            stem, ext = os.path.splitext(n)
            r["new"] = "%s_%d%s" % (stem, used[n], ext)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--execute", action="store_true", help="实际执行重命名（默认仅预览）")
    ap.add_argument("--json", default="", help="导出重命名清单 JSON")
    ap.add_argument("--template", default="standard",
                    choices=["standard", "railway", "legacy12306"],
                    help="命名口径：standard=日期_类型_补充_金额；"
                         "railway=日期_火车发票_出发站_到达站_金额；"
                         "legacy12306=出发站-到达站-日期-金额（仅火车票）")
    ap.add_argument("--station-sep", default="-", help="发站与到站之间的分隔符，默认 -")
    args = ap.parse_args()

    results = build_plan(args.root, args.template, args.station_sep)
    resolve_conflicts(results)

    ok = [r for r in results if "new" in r]
    skipped = [r for r in results if "new" not in r]

    print("=" * 96)
    print("重命名预览：当前目录 %s" % args.root)
    print("=" * 96)
    for r in results:
        if "new" in r:
            flag = "  *改动" if r["new"] != r["old"] else "  相同 "
            print("%s | %s" % (flag, r["old"]))
            print("        -> %s" % r["new"])
        else:
            print("  [跳过] | %s" % r["old"])
            print("         原因：%s" % r["skip_reason"])

    print("\n" + "=" * 96)
    print("汇总：成功 %d / 跳过 %d / 失败 %d / 合计 %d"
          % (len(ok), len(skipped), 0, len(results)))
    print("=" * 96)

    if args.json:
        json.dump(results, open(args.json, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
        print("清单已导出：%s" % args.json)

    if not args.execute:
        print("\n（预览模式，未做任何改动。加 --execute 执行）")
        return

    done = fail = 0
    for r in ok:
        src = r["path"]
        dst = os.path.join(os.path.dirname(src), r["new"])
        try:
            if os.path.abspath(src) == os.path.abspath(dst):
                continue
            if os.path.exists(dst):
                raise FileExistsError(dst)
            os.rename(src, dst)
            done += 1
        except Exception as e:
            fail += 1
            print("  [失败] %s -> %s : %s" % (r["old"], r["new"], e))
    print("\n执行完成：成功 %d，失败 %d" % (done, fail))


if __name__ == "__main__":
    main()
