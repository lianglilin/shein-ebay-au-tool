# -*- coding: utf-8 -*-
"""
规则与底层工具（纯标准库，零依赖）

这个文件只放三块东西，被 ebay_tool.py（主程序）和 shein_fetch.py（抓取）复用：
  1) 可调参数与默认路径（PRICE_MIN/MAX、MARKUP_PRESETS、DEFAULT_INPUT/OUTPUT）
  2) 标题规则引擎（关键词表 build_ebay_title / check_title，澳式拼写、垃圾词、风险词…）
  3) 表格读写（list_sheets / read_sheet）与价格小工具（parse_price / price_ok /
     calc_ebay_price / clean_num）以及 xlsx 用的 _xml_escape / _col_name

界面、抓取、选品流程、汇总表写出分别在 ebay_tool.py / shein_fetch.py 里。
"""

import os
import re
import zipfile
import xml.etree.ElementTree as ET

# ---------- 可调参数 ----------
PRICE_MIN = 9.0
PRICE_MAX = 30.0
PRICE_RULE_TIERED = "tiered_9_20"
PRICE_RULE_TIERED_LABEL = "9-20 +10 / 20.01以上×1.5"
MARKUP_PRESETS = ["1.2", "1.5", "2.0", PRICE_RULE_TIERED_LABEL]

# 默认输入表放在用户的下载目录，可由界面重新选择
DEFAULT_INPUT = os.path.expanduser(
    os.path.join("~", "Downloads", "SHEIN选品链接.xlsx"))
# 默认输出目录不包含个人项目名，可由界面重新选择
DEFAULT_OUTPUT = os.path.expanduser(
    os.path.join("~", "Desktop", "SHEIN选品结果"))

# ---------- 标题规则词表 ----------
# 禁品 / 大牌：标题命中就【不抓】（不打开商品页 / 跳过），不是改标题用的
# 想加词就往对应分类里加，正则语法；改完重启程序生效
BLOCK_PATTERNS = [
    # 儿童用品（禁品）
    (r"\bkids?\b|\bkiddie\b|\bchildren\b|\bchild\b|\bbaby\b|\bbabies\b|"
     r"\binfant\b|\btoddler\b|\bnewborn\b", "儿童用品 kid/child/baby"),
    # 液体 / 精油 / 洗护（跨境物流 & 平台合规风险）
    (r"essential oils?|fragrance oils?|aromatherapy|massage oil|"
     r"hair oils?", "精油/香薰"),
    (r"\bshampoos?\b|(?<!air )\bconditioners?\b|body wash|hand sanitizer|"
     r"sanitizer|disinfectant", "洗护/液体"),
    # 大牌（eBay VeRO / 侵权高风险）
    (r"louis vuitton|\blv\b|chanel|gucci|dior|prada|hermes|hermès|rolex|"
     r"burberry|coach|michael kors|fendi|balenciaga|versace|armani|"
     r"yves saint laurent|\bysl\b|estee lauder|estée lauder|lancome|lancôme|"
     r"kerastase|kérastase|clinique|shiseido|\bsk-ii\b|la mer|bobbi brown|"
     r"nike|adidas|new balance|puma|under armour", "大牌（侵权风险）"),
]


def find_blocked(title, extra_words=None):
    """标题命中禁品/大牌词就返回命中的说明，没命中返回空字符串"""
    t = (title or "")
    low = t.lower()
    for pat, label in BLOCK_PATTERNS:
        m = re.search(pat, low, re.I)
        if m:
            return "%s（命中「%s」）" % (label, m.group(0))
    for w in (extra_words or []):
        w = str(w).strip()
        if w and w.lower() in low:
            return "自定义禁品词（命中「%s」）" % w
    return ""


JUNK_WORDS = [
    "premium", "high quality", "high-quality", "fashionable", "fashion",
    "essential", "perfect", "multifunctional", "multi-functional",
    "multi functional", "durable", "durable and", "modern", "luxury",
    "ideal for", "suitable for", "applicable for", "aesthetic",
    "elegant", "unique", "creative", "portable and", "new arrival",
    "hot sale", "best seller", "limited time", "shop online",
    "free shipping", "high-end", "upgrade", "upgraded", "exquisite",
    "versatile", "convenient", "practical", "beautiful", "stunning",
    "charming", "gorgeous", "superior", "professional",
    "designed specifically for", "designed for", "specifically for",
    "can hold", "perfect for", "great for", "for hanging", "such as",
    "including", "various", "multiple", "etc",
    "it includes", "this is", "made of", "features",
]

# 必须删掉的违规/风险词（医疗、敏感、功效承诺）
DANGER_WORDS = [
    "anti-wrinkle", "anti wrinkle", "anti-aging", "anti aging",
    "hair growth", "anti-hair-loss", "anti hair loss", "hair loss",
    "therapy", "therapeutic", "medical", "treatment", "treats", "cures",
    "relieve pain", "relieves pain", "pain relief", "relieve",
    "detox", "lymphatic", "cellulite", "slimming", "fat burning",
    "weight loss", "body shaping", "acne", "pimples", "eczema",
    "tendonitis", "carpal tunnel", "arthritis", "menstrual", "period pain",
    "improves circulation", "blood circulation", "antibacterial kills",
    "sterilize", "disinfection", "air purification",
    "groin", "pubic", "inguinal", "genital", "intimate area",
]

# 风格词统一成买家更常搜的写法
STYLE_FIX = [
    (r"\bbohemian\b", "Boho"),
]

# 拼写：美式 -> 澳式
SPELLING_AU = [
    (r"\borganizer\b", "organiser"), (r"\borganizers\b", "organisers"),
    (r"\borganizing\b", "organising"), (r"\borganized\b", "organised"),
    (r"\bcolor\b", "colour"), (r"\bcolors\b", "colours"),
    (r"\bcolored\b", "coloured"),
    (r"\bjewelry\b", "jewellery"), (r"\bjewellerys\b", "jewellery"),
    (r"\bpersonalized\b", "personalised"), (r"\bpersonalize\b", "personalise"),
    (r"\bcustomized\b", "customised"), (r"\bcustomize\b", "customise"),
    (r"\bgray\b", "grey"),
    (r"\btire\b", "tyre"), (r"\btires\b", "tyres"),
    (r"\baluminum\b", "aluminium"),
    (r"\bpajamas\b", "pyjamas"),
    (r"\bfavor\b", "favour"), (r"\bfavors\b", "favours"),
    (r"\bmicrofiber\b", "microfibre"),
    (r"\bmeter\b", "metre"), (r"\bmeters\b", "metres"),
    (r"\bcenter\b", "centre"),
    (r"\breusable\b", "reusable"),
]

# 澳洲本地用词替换
LOCAL_AU = [
    (r"\btrash can\b", "rubbish bin"), (r"\btrash bins\b", "rubbish bins"),
    (r"\bgarbage bag\b", "bin liner"), (r"\bgarbage bags\b", "bin liners"),
    (r"\btrunk\b", "boot"),
    (r"\bwindshield\b", "windscreen"),
    (r"\bleash\b", "lead"), (r"\bleashes\b", "leads"),
    (r"\bwrench\b", "spanner"), (r"\bwrenches\b", "spanners"),
    (r"\bclothespin\b", "clothes peg"), (r"\bclothespins\b", "clothes pegs"),
    (r"\bchristmas ornaments\b", "christmas baubles"),
    (r"\bpillow case\b", "cushion cover"), (r"\bpillowcase\b", "cushion cover"),
]


# ============ xlsx 读取（纯标准库） ============
def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _col_index(col_letters):
    n = 0
    for ch in col_letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _read_entry(zf, name):
    try:
        return zf.read(name)
    except KeyError:
        return None


def list_sheets(xlsx_path):
    """返回 [{'name':..., 'path':...}]"""
    with zipfile.ZipFile(xlsx_path) as zf:
        wb = ET.fromstring(_read_entry(zf, "xl/workbook.xml"))
        rels_xml = _read_entry(zf, "xl/_rels/workbook.xml.rels")
        rid2target = {}
        if rels_xml:
            for rel in ET.fromstring(rels_xml):
                rid2target[rel.get("Id")] = rel.get("Target")
        out = []
        for node in wb.iter():
            if _local(node.tag) != "sheet":
                continue
            rid = None
            for k, v in node.attrib.items():
                if _local(k) == "id":
                    rid = v
            target = rid2target.get(rid, "")
            if target and not target.startswith("xl/"):
                target = "xl/" + target.lstrip("/")
            out.append({"name": node.get("name"), "path": target})
        return out


def _shared_strings(zf):
    data = _read_entry(zf, "xl/sharedStrings.xml")
    if not data:
        return []
    root = ET.fromstring(data)
    out = []
    for si in root:
        if _local(si.tag) != "si":
            continue
        buf = []
        for t in si.iter():
            if _local(t.tag) == "t" and t.text:
                buf.append(t.text)
        out.append("".join(buf))
    return out


def _sheet_rels(zf, sheet_path):
    base = os.path.basename(sheet_path)
    data = _read_entry(zf, "xl/worksheets/_rels/%s.rels" % base)
    out = {}
    if not data:
        return out
    for rel in ET.fromstring(data):
        out[rel.get("Id")] = (rel.get("Target"), rel.get("Type", ""))
    return out


def read_sheet(xlsx_path, sheet_name):
    """
    读取指定工作表，返回 (rows, hyperlinks)
      rows: [ {rownum, cells:{colidx:text}} ]
      hyperlinks: { 'G2': 'https://...' }
    """
    with zipfile.ZipFile(xlsx_path) as zf:
        sheets = list_sheets(xlsx_path)
        target = None
        for s in sheets:
            if s["name"] == sheet_name:
                target = s["path"]
                break
        if not target:
            raise ValueError("\u627e\u4e0d\u5230\u5de5\u4f5c\u8868: %s" % sheet_name)

        shared = _shared_strings(zf)
        root = ET.fromstring(_read_entry(zf, target))

        # --- 单元格 ---
        rows = []
        for row in root.iter():
            if _local(row.tag) != "row":
                continue
            rownum = int(row.get("r") or 0)
            cells = {}
            for c in row:
                if _local(c.tag) != "c":
                    continue
                ref = c.get("r") or ""
                m = re.match(r"([A-Z]+)", ref)
                if not m:
                    continue
                ci = _col_index(m.group(1))
                ctype = c.get("t")
                val = None
                for child in c:
                    ln = _local(child.tag)
                    if ln == "v":
                        val = child.text
                    elif ln == "is":
                        buf = []
                        for t in child.iter():
                            if _local(t.tag) == "t" and t.text:
                                buf.append(t.text)
                        val = "".join(buf)
                if ctype == "s" and val is not None:
                    try:
                        val = shared[int(val)]
                    except Exception:
                        pass
                cells[ci] = val
            rows.append({"rownum": rownum, "cells": cells})

        # --- 超链接 ---
        links = {}
        rels = _sheet_rels(zf, target)
        for hl in root.iter():
            if _local(hl.tag) != "hyperlink":
                continue
            rid = None
            for k, v in hl.attrib.items():
                if _local(k) == "id":
                    rid = v
            if rid and rid in rels:
                tgt = rels[rid][0]
                if tgt.startswith("http"):
                    links[hl.get("ref")] = tgt
                else:
                    links[hl.get("ref")] = tgt
            elif hl.get("location"):
                links[hl.get("ref")] = hl.get("location")
        return rows, links


# ============ 标题规则引擎 ============
def _strip_site_suffix(t):
    t = re.sub(r"for Sale Australia.*$", "", t, flags=re.I)
    t = re.sub(r"\|\s*New Collection Online.*$", "", t, flags=re.I)
    t = re.sub(r"\|\s*SHEIN Australia.*$", "", t, flags=re.I)
    t = re.sub(r"\bSHEIN\b", "", t, flags=re.I)
    return t


def _strip_qty_prefix(t):
    # 只去掉「一件代发」标记：1pc / 1 pcs / 1 piece / 1 set；
    # 3pcs、15pcs 这种是真实件数，要留在标题里
    t = re.sub(r"^\s*[\[\uff08(]?\s*1\s*(pcs|pieces|piece|packs|pack|sets|set|pc)"
               r"\s*[\]\uff09)]?\s*",
               "", t, flags=re.I)
    # 没有数字的裸前缀，例如 "pc Vintage ..." / "pcs Shoe Rack"
    t = re.sub(r"^\s*(pcs|pieces|piece|packs|pack|sets|set|pc)\s+", "", t, flags=re.I)
    t = re.sub(r"^\s*\[?\s*\d+\s*[*xX]\s*\]?\s*", "", t)
    return t


def _strip_qty_anywhere(t):
    """删除句中残留的一件代发标记：1pc / 1 pcs / 1 piece / 1 set
    注意只删「1」，保留 4 Piece 这种真实件数。"""
    t = re.sub(r"(?i)(^|\s)1\s*(pcs|pieces|piece|packs|pack|sets|set|pc)(?=\s|$)",
               " ", t)
    return t


def _drop_brand(t):
    """
    首词如果是全大写英文(>=3位, 不含数字)、且不在白名单里, 视为品牌删掉。
    含数字的（如 SK5、V-613）不动 —— 那通常是材料型号而不是品牌。
    """
    m = re.match(r"^([A-Z][A-Z0-9]{1,19})\s+", t)
    if not m:
        return t
    word = m.group(1)
    keep = {"USB", "LED", "LCD", "ABS", "PP", "PVC", "3D", "4D", "5D", "DIY",
            "XL", "XXL", "AU", "UK", "US", "SET", "NEW", "PC", "PCS", "CM", "MM",
            "EDC", "PU", "PE", "EVA", "TPE", "SOS", "UV", "HD", "TPU"}
    if word in keep or any(ch.isdigit() for ch in word):
        return t
    return t[m.end():]


def _remove_words(t, words):
    for w in words:
        t = re.sub(r"(?i)(?<![a-z])" + re.escape(w) + r"(?![a-z])", " ", t)
    return t


def _dedup_words(t):
    parts = t.split()
    seen = set()
    out = []
    for p in parts:
        key = p.lower().strip(",.")
        if key and key in seen:
            continue
        seen.add(key)
        out.append(p)
    return " ".join(out)


def _tidy(t):
    t = t.replace("|", " ").replace("/", " / ")
    t = re.sub(r"\s*[\-\u2013\u2014]+\s*", " ", t)   # 破折号 -> 空格
    t = re.sub(r"[.;]+\s*", " ", t)                  # 句号/分号 -> 空格
    t = re.sub(r"[,:]+\s*", " ", t)
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"\s*/\s*", "/", t)
    return t.strip(" -,:;.")


# 结尾不能留的虚词 / 半截词
DANGLING = {
    "for", "with", "and", "or", "of", "the", "a", "an", "in", "on", "to", "by",
    "from", "at", "as", "suitable", "designed", "specifically", "can", "hold",
    "etc", "including", "include", "such", "various", "multiple", "new", "use",
    "using", "used", "is", "are", "be", "your", "our",
    "protects", "comes", "features", "made", "there", "that", "which",
}

SMALL_WORDS = {"a", "an", "the", "and", "or", "for", "with", "of", "in", "on",
               "to", "by", "from", "at", "as"}

KEEP_UPPER = {"USB", "LED", "LCD", "ABS", "PP", "PVC", "DIY", "XL", "XXL",
              "AU", "UK", "US", "CM", "MM", "KG", "ML", "PC", "PCS", "3D",
              "4D", "5D", "2T", "ID", "TV", "HD", "AC", "DC"}


def _strip_dangling(t):
    words = t.split()
    while words:
        w = words[-1].strip(" -,:;()")
        if w == "" or w.lower() in DANGLING:
            words.pop()
            continue
        break
    return " ".join(words)


def _title_case(t):
    out = []
    for i, w in enumerate(t.split()):
        bare = w.strip("(),")
        if bare.upper() in KEEP_UPPER:
            out.append(w)
            continue
        low = w.lower()
        if i > 0 and low in SMALL_WORDS:
            out.append(low)
            continue
        if w[:1].isalpha():
            out.append(w[:1].upper() + w[1:])
        else:
            out.append(w)
    return " ".join(out)


def _truncate(t, limit):
    """按词边界截断到 limit 以内，并去掉结尾虚词"""
    if len(t) <= limit:
        return t
    cut = t[:limit]
    if " " in cut:
        cut = cut[:cut.rfind(" ")]
    return cut


def build_ebay_title(src_title, limit=80):
    """把希音标题按规则改写成 eBay 标题（<=80）"""
    t = src_title or ""
    t = _strip_site_suffix(t)
    t = _strip_qty_prefix(t)
    t = _strip_qty_anywhere(t)
    t = _drop_brand(t)
    t = t.replace("&", " and ").replace("\u2013", " ").replace("\u2014", " ")
    t = _remove_words(t, DANGER_WORDS)
    t = _remove_words(t, JUNK_WORDS)
    t = _dedup_words(t)
    for pat, rep in SPELLING_AU:
        t = re.sub(pat, rep, t, flags=re.I)
    for pat, rep in LOCAL_AU:
        t = re.sub(pat, rep, t, flags=re.I)
    for pat, rep in STYLE_FIX:
        t = re.sub(pat, rep, t, flags=re.I)
    t = _tidy(t)
    t = _strip_dangling(t)
    if len(t) > limit:
        t = _strip_dangling(_truncate(t, limit))
    t = _title_case(t)
    t = re.sub(r"\s+", " ", t).strip(" -,:;")
    return t


def check_title(title):
    """返回问题列表"""
    problems = []
    if not title:
        problems.append("\u6807\u9898\u4e3a\u7a7a")
        return problems
    if len(title) > 80:
        problems.append("\u8d85\u8fc7 80 \u5b57\u7b26")
    low = title.lower()
    for w in DANGER_WORDS:
        if w in low:
            problems.append("\u542b\u98ce\u9669\u8bcd: " + w)
    return problems


# ============ 价格 ============
def parse_price(v):
    if v is None:
        return None
    s = str(v).replace(",", "").replace("AU$", "").replace("$", "").strip()
    try:
        return float(s)
    except Exception:
        return None


def clean_num(v):
    """把 '1.0' 这类浮点字符串转成 '1'，其它原样返回"""
    if v is None:
        return ""
    s = str(v).strip()
    if s == "":
        return ""
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
        return str(f)
    except Exception:
        return s


def price_ok(p):
    return p is not None and PRICE_MIN <= p <= PRICE_MAX


def normalize_price_rule(value):
    """把界面选项转成计算用的规则；无效输入返回 None。"""
    text = str(value or "").strip()
    if text == PRICE_RULE_TIERED or text == PRICE_RULE_TIERED_LABEL:
        return PRICE_RULE_TIERED
    text = text.replace("×", "").replace("x", "").replace("X", "").strip()
    try:
        return float(text)
    except Exception:
        return None


def calc_ebay_price(p, factor):
    """计算 eBay 售价。

    兼容原来的固定倍数（1.2/1.5/2.0），以及新增的阶梯规则：
    9 <= 希音价 <= 20 加 10；20.01 及以上乘 1.5。
    """
    if p is None:
        return None
    if factor == PRICE_RULE_TIERED or factor == PRICE_RULE_TIERED_LABEL:
        p = float(p)
        return round(p + 10.0 if p <= 20.0 else p * 1.5, 2)
    return round(float(p) * float(factor), 2)


# ============ 输出 xlsx（最小实现） ============
def _xml_escape(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _col_name(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s
