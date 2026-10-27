# -*- coding: utf-8 -*-
"""SHEIN -> eBay AU 选品工具主程序。

负责桌面界面、选品流程、商品图片下载和 XLSX 汇总结果写出。
依赖 websocket-client 和本机 Chrome/Edge。
"""

import datetime
import os
import queue
import random
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor

import core_rules as CR
from shein_fetch import (NeedHuman, NeedLanguage, SheinScraper, Stopped,
                         find_browser, find_chrome, find_edge,
                         is_connection_error, with_url_params)

APP_VERSION = "0.7.2"
APP_TITLE = "SHEIN -> eBay 选品工具  v%s" % APP_VERSION

IS_MAC = sys.platform == "darwin"


def _packaged_exe():
    """打包运行时返回真正的 exe 路径，源码运行返回 None。
    坑：Nuitka onefile 的 sys.executable 指向临时解包目录（退出就删），
    真正的 exe 要从 __compiled__.process_exe / sys.argv[0] 拿。"""
    c = globals().get("__compiled__")                     # Nuitka
    if c is not None:
        for attr in ("process_exe", "original_argv0"):
            p = getattr(c, attr, None)
            if p and os.path.exists(p):
                return os.path.abspath(p)
    if getattr(sys, "frozen", False):                     # PyInstaller / cx_Freeze
        return os.path.abspath(sys.executable)
    a0 = sys.argv[0] if sys.argv else ""
    if a0.lower().endswith(".exe") and os.path.exists(a0):
        return os.path.abspath(a0)
    return None


def _app_dir():
    """程序自己的目录（放 Chrome 档案 / 全局记录用）
    - 源码运行：脚本所在目录
    - 打包后：exe 所在目录
    - macOS .app：包内是只读的，放到 ~/Library/Application Support/SHEIN选品工具"""
    exe = _packaged_exe()
    if exe:
        base = os.path.dirname(exe)
        if IS_MAC and ".app/Contents/MacOS" in base:
            return os.path.expanduser("~/Library/Application Support/SHEIN选品工具")
        return base
    return os.path.dirname(os.path.abspath(__file__))


HERE = _app_dir()
try:
    os.makedirs(HERE, exist_ok=True)
except Exception:
    pass
PROFILE_DIR = os.path.join(HERE, "_chrome_profile")
if IS_MAC:
    EDGE_SOURCE_PROFILE_DIR = os.path.expanduser(
        "~/Library/Application Support/Microsoft Edge")
else:
    EDGE_SOURCE_PROFILE_DIR = os.path.join(
        os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
        "Microsoft", "Edge", "User Data")
EDGE_PROFILE_DIR = os.path.join(HERE, "_edge_profile")
HIST_FILE = os.path.join(HERE, "_已选过记录.txt")   # 全局记忆：这个工具抓过/选过的商品
DEFAULT_PORT = 9223


def _close_default_edge(log=print):
    """关闭正常 Edge，确保能复制到一致的 Cookie 数据。"""
    if not sys.platform.startswith("win"):
        return
    try:
        ps = (
            "Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | "
            "Where-Object { $_.CommandLine -like '*--profile-directory=Default*' "
            "-or $_.CommandLine -like '*--type=crashpad-handler*' } | "
            "Select-Object -ExpandProperty ProcessId"
        )
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                           capture_output=True, text=True, timeout=30)
        pids = [x for x in (r.stdout or "").split() if x.strip().isdigit()]
        for pid in pids:
            subprocess.run(["taskkill", "/PID", pid, "/T", "/F"],
                           capture_output=True, timeout=20)
        if pids:
            log("已关闭正常 Edge，准备同步会话档案 …")
            time.sleep(2)
    except Exception as exc:
        log("关闭正常 Edge 时忽略错误：%s" % exc)


def prepare_edge_profile(log=print):
    """从正常 Edge 的 Default 档案复制一份可供调试抓取的独立副本。

    Edge 不允许直接用默认档案开启远程调试；这里只复制 Cookie、本地存储、
    优惠和登录相关数据，不复制缓存、历史记录和扩展缓存。
    """
    ready = (
        os.path.exists(os.path.join(EDGE_PROFILE_DIR, "Local State"))
        and os.path.exists(os.path.join(
            EDGE_PROFILE_DIR, "Default", "Preferences"))
        and os.path.exists(os.path.join(
            EDGE_PROFILE_DIR, "Default", "Network"))
    )
    if ready:
        return True
    src = os.path.join(EDGE_SOURCE_PROFILE_DIR, "Default")
    if not os.path.isdir(src):
        log("！找不到 Edge Default 档案：%s" % src)
        return False

    _close_default_edge(log)
    os.makedirs(os.path.join(EDGE_PROFILE_DIR, "Default"), exist_ok=True)
    local_state = os.path.join(EDGE_SOURCE_PROFILE_DIR, "Local State")
    try:
        if os.path.exists(local_state):
            shutil.copy2(local_state, EDGE_PROFILE_DIR)
        for name in (
                "Preferences", "Secure Preferences", "Login Data",
                "Login Data For Account", "Extension Cookies",
                "Web Data", "Shortcuts"):
            src_file = os.path.join(src, name)
            if os.path.isfile(src_file):
                shutil.copy2(src_file,
                             os.path.join(EDGE_PROFILE_DIR, "Default", name))
        for name in (
                "Network", "Local Storage", "Sessions", "WebStorage",
                "Sync Data", "EdgeCoupons"):
            src_dir = os.path.join(src, name)
            dst_dir = os.path.join(EDGE_PROFILE_DIR, "Default", name)
            if os.path.isdir(src_dir):
                shutil.copytree(src_dir, dst_dir, dirs_exist_ok=True)
        log("Edge 当前档案已复制到工具专用目录：%s" % EDGE_PROFILE_DIR)
        return True
    except Exception as exc:
        log("！同步 Edge 档案失败：%s" % exc)
        return False


def browser_settings(name="Chrome"):
    """返回浏览器可执行文件、用户目录和额外启动参数。

    Edge 默认使用当前 Default 档案，因为价格优惠可能和该档案的
    Cookie/登录状态绑定；Chrome 仍使用工具自己的独立档案。
    """
    label = "Edge" if str(name or "").strip().lower() == "edge" else "Chrome"
    exe = find_edge() if label == "Edge" else find_chrome()
    if label == "Edge":
        return {
            "name": label,
            "exe": exe,
            "profile": EDGE_PROFILE_DIR,
            "args": [
                "--profile-directory=Default",
                "--lang=en-AU",
                "--accept-lang=en-AU,en",
            ],
        }
    return {
        "name": label,
        "exe": exe,
        "profile": PROFILE_DIR,
        "args": [],
    }

IMG_MAX_PER_ITEM = 0          # 每个商品最多下载几张图；0 = 不限，页面能拿到的全下
IMG_EMBED_H = 120             # 汇总表里图片显示高度(px)
IMG_EMBED_MARGIN_PX = 4       # 图片距离商品行上、左边缘的留白
DATA_ROW_HEIGHT_PT = 96       # 所有商品行统一行高，保证有图/无图行不跳变

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36")

IMG_COL = 12
TXT = "商品信息.txt"

# ---- 汇总表（按老板那份的列顺序）----
CAT_HEADERS = ["序号", "选品日期", "希音sku", "产品图片", "产品金额", "运费",
               "希音链接", "希音卖家ID", "希音标题", "eBay价格", "上架店铺",
               "变体", "备注", "是否上架（日期）", "评分", "销量", "库存"]
CAT_IMG_COL = 3
CAT_LINK_COL = 6
DEFAULT_CAT_URL = "https://au.shein.com/Beauty-Health-c-1864.html"

# 抓取节奏：手动慢慢看不会被拦，所以我们刻意放慢
PACE_PRESETS = {
    "稳妥（慢，最不容易被拦）": ((10.0, 20.0), 6, (30, 60)),
    "正常": ((5.0, 10.0), 10, (20, 40)),
    "快（容易触发验证）": ((1.0, 3.0), 20, (10, 20)),
}

# 被希音拦住（要验证 / 要登录）时怎么办
WALL_ACTIONS = {
    "自动歇一会儿再试（不登录，推荐）": "cool",
    "等我手动处理（在浏览器里点一下）": "wait",
    "跳过这个商品，继续下一个": "skip",
}


# ---------------- 图片 ----------------
def download_image(url, tries=4, timeout=25):
    last = None
    for n in range(tries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA,
                "Referer": "https://au.shein.com/",
                "Accept": "image/avif,image/webp,image/png,image/*,*/*;q=0.8",
            })
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = r.read()
            if data and len(data) > 512:
                return data
            last = "空数据"
        except Exception as exc:
            last = exc
            time.sleep(1.0 * (n + 1))
    raise RuntimeError("下载失败 %s (%s)" % (url[-40:], last))


def ext_of(data):
    if data[:2] == b"\xff\xd8":
        return ".jpg"
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return ".png"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return ".webp"
    return ".jpg"


def image_for_embed(data):
    """把要嵌进 xlsx 的图片转成 Excel 稳定显示的格式。

    商品文件夹保留原始 WebP，但 Excel/WPS 对 WebP 的兼容性不一致，直接
    嵌入会出现“图片对象存在但显示空白”。这里只转换用于表格预览的字节，
    不改文件夹里的原图。
    """
    if not data:
        return data
    if ext_of(data) in (".jpg", ".jpeg", ".png"):
        return data
    try:
        from io import BytesIO
        from PIL import Image, ImageOps
        with Image.open(BytesIO(data)) as im:
            im = ImageOps.exif_transpose(im)
            if im.mode in ("RGBA", "LA") or (
                    im.mode == "P" and "transparency" in im.info):
                rgba = im.convert("RGBA")
                bg = Image.new("RGB", rgba.size, (255, 255, 255))
                bg.paste(rgba, mask=rgba.getchannel("A"))
                im = bg
            else:
                im = im.convert("RGB")
            im.thumbnail((900, 1200), Image.Resampling.LANCZOS)
            out = BytesIO()
            im.save(out, format="JPEG", quality=88, optimize=True)
            return out.getvalue()
    except Exception:
        # 没有 Pillow 或图片损坏时保留原字节；调用方可从日志看到原始问题。
        return data


def download_images(urls, folder, limit=IMG_MAX_PER_ITEM, log=print):
    """并发下载，返回 [(文件名, 字节)]，第一个即主图"""
    if limit and limit > 0:
        urls = urls[:limit]
    if not urls:
        return []
    out = []
    with ThreadPoolExecutor(max_workers=4) as pool:
        futs = {pool.submit(download_image, u): i for i, u in enumerate(urls)}
        results = {}
        for fut in futs:
            idx = futs[fut]
            try:
                results[idx] = fut.result()
            except Exception as exc:
                log("      图片 %d 下载失败: %s" % (idx + 1, exc))
    for idx in sorted(results):
        data = results[idx]
        name = "%d%s" % (idx + 1, ext_of(data))
        try:
            with open(os.path.join(folder, name), "wb") as fh:
                fh.write(data)
            out.append((name, data))
        except Exception as exc:
            log("      图片写盘失败: %s" % exc)
    return out


# ---------------- 汇总表（含内嵌图片，纯标准库） ----------------
EMU_PX = 9525
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="%s">'
    '<fonts count="2">'
    '<font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>'
    '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font>'
    '</fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border/></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="4">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
    '<alignment vertical="center"/></xf>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">'
    '<alignment vertical="center"/></xf>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>' % NS_MAIN)


def _emu(px):
    return int(px * EMU_PX)


def _data_row_height_pt(img_h):
    """按嵌入图片高度计算统一商品行高；所有数据行必须使用同一高度。"""
    return max(DATA_ROW_HEIGHT_PT, int(float(img_h) * 0.75 + 6))


def _date_cell_value(value):
    """把 Excel 日期序列转成稳定的 YYYY-MM-DD 文本。

    openpyxl/我们的纯 XML 写入器在重写工作簿时会丢失日期格式；
    保住显示值比保留裸序列号更重要，因此日期列统一写成文本。
    """
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, datetime.date):
        return value.isoformat()
    try:
        text = str(value).strip()
        num = float(text)
        if 20000 <= num <= 80000:
            return (datetime.datetime(1899, 12, 30) +
                    datetime.timedelta(days=num)).strftime("%Y-%m-%d")
    except Exception:
        pass
    return value


def _drawing_xml(items, cx, cy, col_off=0, row_off=0):
    parts = []
    for i, (col, row, rid) in enumerate(items):
        parts.append(
            '<xdr:oneCellAnchor>'
            '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:from>'
            '<xdr:ext cx="%d" cy="%d"/>'
            '<xdr:pic><xdr:nvPicPr>'
            '<xdr:cNvPr id="%d" name="Picture %d"/>'
            '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>'
            '</xdr:nvPicPr><xdr:blipFill>'
            '<a:blip xmlns:r="%s" r:embed="%s"/>'
            '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
            '</xdr:pic><xdr:clientData/></xdr:oneCellAnchor>'
            % (col, col_off, row, row_off, cx, cy,
               i + 2, i + 2, NS_REL, rid, cx, cy))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<xdr:wsDr xmlns:xdr="%s" xmlns:a="%s">%s</xdr:wsDr>'
            % (NS_XDR, NS_A, "".join(parts)))


# ============ 多工作表工作簿（一个文件放多批，按编号区间分页） ============
def _sheet_range_name(rows, prefix="备份"):
    """按数据里的编号算工作表名，例如 备份109-125"""
    nums = []
    for r in rows:
        try:
            nums.append(int(float(r[0])))
        except Exception:
            continue
    if not nums:
        return prefix
    return "%s%d-%d" % (prefix, min(nums), max(nums))


def _trim_trailing_blank_rows(rows):
    """去掉工作簿尾部由 XML 产生的空行，避免重复写几千个空行。"""
    rows = list(rows or [])
    while rows and all(v in (None, "") for v in rows[-1]):
        rows.pop()
    return rows


def _order_output_rows(rows, images=None, links=None):
    """展示顺序：所有通过行在前，所有跳过行在后。

    不改变行内容/编号，只重排展示行，并重新映射图片和超链接的行号。
    跳过行统一带 ``跳过：`` 前缀；兼容旧表里没有前缀的备注。
    """
    records = []
    for idx, row in enumerate(rows or [], start=2):
        note = str(row[12] or "") if len(row) > 12 else ""
        is_skip = note.startswith("跳过")
        records.append((1 if is_skip else 0, idx, row,
                        (images or {}).get(idx),
                        (links or {}).get(idx)))
    records.sort(key=lambda item: (item[0], item[1]))
    ordered_rows = [item[2] for item in records]
    ordered_imgs = {}
    ordered_links = {}
    for new_idx, (_skip, _old_idx, _row, image, link) in enumerate(
            records, start=2):
        if image is not None:
            ordered_imgs[new_idx] = image
        if link:
            ordered_links[new_idx] = link
    return ordered_rows, ordered_imgs, ordered_links


def _sheet_range_bounds(name):
    m = re.match(r"^备份(\d+)-(\d+)$", str(name or "").strip())
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def _row_identity(row):
    try:
        return "num:%d" % int(float(row[0]))
    except Exception:
        return "url:%s" % str(row[6] if len(row) > 6 else "")


def _merge_summary_rows(existing, current_rows, current_imgs, current_links):
    """按序号合并，不删除本次没有出现的旧行。"""
    records = {}
    order = []

    def add(row, image=None, link=None):
        key = _row_identity(row)
        if key not in records:
            order.append(key)
        old = records.get(key, {})
        records[key] = {
            "row": row,
            "image": image if image is not None else old.get("image"),
            "link": link or old.get("link"),
        }

    if existing:
        old_imgs = existing.get("images") or {}
        old_links = existing.get("links") or {}
        for idx, row in enumerate(existing.get("rows") or [], 2):
            add(row, old_imgs.get(idx), old_links.get(idx))
    for idx, row in enumerate(current_rows or [], 2):
        add(row, (current_imgs or {}).get(idx),
            (current_links or {}).get(idx))

    merged_rows, merged_imgs, merged_links = [], {}, {}
    for new_idx, key in enumerate(order, 2):
        rec = records[key]
        merged_rows.append(rec["row"])
        if rec["image"] is not None:
            merged_imgs[new_idx] = rec["image"]
        if rec["link"]:
            merged_links[new_idx] = rec["link"]
    return _order_output_rows(merged_rows, merged_imgs, merged_links)


def _safe_sheet_name(name, used=None):
    name = re.sub(r"[\[\]:*?/\\]", "-", str(name or "Sheet"))[:31] or "Sheet"
    used = used if used is not None else set()
    base, i = name, 2
    while name.lower() in used:
        suf = "(%d)" % i
        name = base[:31 - len(suf)] + suf
        i += 1
    used.add(name.lower())
    return name


def write_workbook(path, sheets):
    """把多张表写进一个工作簿。
    sheets: [{name, headers, rows, images{行:字节}, links{行:url}, img_col, link_col, img_h}]
    """
    media = []            # [(文件名, 字节)]
    drawings = []         # [(sheet序号, anchors, 图片显示高度)]
    sheet_xmls = []       # [(路径, xml)]
    sheet_rels = []       # [(路径, xml 或 None)]
    exts = set()
    used_names = set()

    for si, sh in enumerate(sheets, 1):
        headers = sh.get("headers") or []
        rows = sh.get("rows") or []
        images = sh.get("images") or {}
        links = sh.get("links") or {}
        date_cols = {ci for ci, h in enumerate(headers)
                     if "日期" in str(h)}
        hidden_cols = {ci for ci, h in enumerate(headers)
                       if str(h).strip() in ("评分", "销量", "库存")}
        img_col = sh.get("img_col")
        link_col = sh.get("link_col")
        img_h = sh.get("img_h", IMG_EMBED_H)
        if img_col is None:
            img_col = 0
        all_rows = [headers] + list(rows)

        anchors = []
        link_items = []
        for ri, blob in sorted(images.items()):
            if not blob:
                continue
            blob = image_for_embed(blob)
            ext = ext_of(blob)
            exts.add(ext.lstrip("."))
            fname = "image%d%s" % (len(media) + 1, ext)
            media.append((fname, blob))
            anchors.append({"row": ri - 1, "col": img_col, "rid": None, "fname": fname})
        for ri, url in sorted(links.items()):
            if url:
                link_items.append((ri, url))

        body = []
        for ri, row in enumerate(all_rows, 1):
            cells = []
            for ci, val in enumerate(row):
                if ci in date_cols:
                    val = _date_cell_value(val)
                if val is None or val == "":
                    continue
                ref = "%s%d" % (CR._col_name(ci), ri)
                style = ' s="3"' if (link_col is not None and ci == link_col
                                     and ri in links) else ' s="2"'
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    cells.append('<c r="%s"%s><v>%s</v></c>' % (ref, style, val))
                else:
                    cells.append('<c r="%s"%s t="inlineStr"><is>'
                                 '<t xml:space="preserve">%s</t></is></c>'
                                 % (ref, style, CR._xml_escape(val)))
            attr = (' ht="%g" customHeight="1"' % _data_row_height_pt(img_h)
                    if ri > 1 else "")
            body.append('<row r="%d"%s>%s</row>' % (ri, attr, "".join(cells)))

        col_parts = []
        if anchors and img_col is not None:
            col_parts.append('<col min="%d" max="%d" width="19" customWidth="1"/>'
                             % (img_col + 1, img_col + 1))
        for ci in sorted(hidden_cols):
            col_parts.append('<col min="%d" max="%d" width="12" hidden="1" customWidth="1"/>'
                             % (ci + 1, ci + 1))
        cols = "<cols>%s</cols>" % "".join(col_parts) if col_parts else ""

        rids = []          # 这张表的 rels
        hl = ""
        if link_items and link_col is not None:
            start = 2 if anchors else 1
            items = []
            for i, (ri, url) in enumerate(link_items):
                rid = "rId%d" % (start + i)
                items.append('<hyperlink ref="%s%d" r:id="%s"/>'
                             % (CR._col_name(link_col), ri, rid))
                rids.append((rid, "hyperlink", url, True))
            hl = "<hyperlinks>%s</hyperlinks>" % "".join(items)

        drawing_tag = ""
        if anchors:
            drawing_tag = '<drawing r:id="rId1"/>'
            rids.insert(0, ("rId1", "drawing", "../drawings/drawing%d.xml" % si, False))

        sname = _safe_sheet_name(sh.get("name") or ("Sheet%d" % si), used_names)
        sheet_xmls.append((sname, "xl/worksheets/sheet%d.xml" % si,
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           '<worksheet xmlns="%s" xmlns:r="%s">%s<sheetData>%s</sheetData>%s%s'
                           '</worksheet>'
                           % (NS_MAIN, NS_REL, cols, "".join(body), hl, drawing_tag)))
        if rids:
            rel = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
            for rid, kind, target, ext_link in rids:
                rel.append('<Relationship Id="%s" Type="%s/%s" Target="%s"%s/>'
                           % (rid, NS_REL, kind, CR._xml_escape(target),
                              ' TargetMode="External"' if ext_link else ""))
            rel.append("</Relationships>")
            sheet_rels.append(("xl/worksheets/_rels/sheet%d.xml.rels" % si, "".join(rel)))
        if anchors:
            drawings.append((si, anchors, img_h))

    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>']
    for ext in sorted(exts):
        ct.append('<Default Extension="%s" ContentType="image/%s"/>'
                  % (ext, "jpeg" if ext in ("jpg", "jpeg") else ext))
    ct.append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    ct.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    for si in range(1, len(sheet_xmls) + 1):
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % si)
    for si, _anchors, _img_h in drawings:
        ct.append('<Override PartName="/xl/drawings/drawing%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' % si)
    ct.append("</Types>")

    wb_sheets, wb_rels = [], []
    for si, (sname, _p, _x) in enumerate(sheet_xmls, 1):
        wb_sheets.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
                         % (CR._xml_escape(sname), si, si))
        wb_rels.append('<Relationship Id="rId%d" Type="%s/worksheet" '
                       'Target="worksheets/sheet%d.xml"/>' % (si, NS_REL, si))
    wb_rels.append('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>'
                   % (len(sheet_xmls) + 1, NS_REL))
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets></workbook>'
                % (NS_MAIN, NS_REL, "".join(wb_sheets)))

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "".join(ct))
        zf.writestr("_rels/.rels",
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
                    '</Relationships>' % NS_REL)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels",
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    + "".join(wb_rels) + "</Relationships>")
        zf.writestr("xl/styles.xml", STYLES_XML)
        for _sname, spath, sxml in sheet_xmls:
            zf.writestr(spath, sxml)
        for rpath, rxml in sheet_rels:
            zf.writestr(rpath, rxml)
        for si, anchors, sheet_img_h in drawings:
            cx = _emu(int(sheet_img_h * 0.75))
            cy = _emu(int(sheet_img_h))
            zf.writestr("xl/drawings/drawing%d.xml" % si,
                        _drawing_xml(
                            [(a["col"], a["row"], "rId%d" % (i + 1))
                             for i, a in enumerate(anchors)],
                            cx, cy,
                            col_off=_emu(IMG_EMBED_MARGIN_PX),
                            row_off=_emu(IMG_EMBED_MARGIN_PX)))
            rel = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
            for i, a in enumerate(anchors, 1):
                rel.append('<Relationship Id="rId%d" Type="%s/image" '
                           'Target="../media/%s"/>' % (i, NS_REL, a["fname"]))
            rel.append("</Relationships>")
            zf.writestr("xl/drawings/_rels/drawing%d.xml.rels" % si, "".join(rel))
        for fname, blob in media:
            zf.writestr("xl/media/" + fname, blob)


# ---------------- 编号（接着已有内容往下排，绝不覆盖） ----------------
LOCK_NAME = ".选品工具.lock"


def _sleep(sec, stop_event=None, log=None, label=""):
    """可以被「停止」打断的等待；返回 True 表示中途被停了"""
    end = time.time() + float(sec)
    last_mark = None
    while time.time() < end:
        if stop_event is not None and stop_event.is_set():
            return True
        time.sleep(1)
        left = max(0, int(round(end - time.time())))
        if log and left and left % 30 == 0 and left != last_mark:
            log("      %s，还剩 %d 秒（点「停止」可中断）" % (label or "等待中", left))
            last_mark = left
    return False


def open_folder(path):
    """用系统默认方式打开文件夹（Windows: 资源管理器 / macOS: 访达 / Linux: xdg-open）"""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)                       # noqa: S606
        elif IS_MAC:
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
        return True
    except Exception:
        return False


def _pid_alive(pid):
    """这个进程还活着吗（只查询，不杀）。Windows 走 ctypes，macOS/Linux 走 os.kill(pid,0)"""
    if not pid:
        return False
    if not sys.platform.startswith("win"):
        try:
            os.kill(int(pid), 0)
            return True
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        except Exception:
            return True
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        h = k32.OpenProcess(0x1000, False, int(pid))     # QUERY_LIMITED_INFORMATION
        if not h:
            return False
        code = ctypes.c_ulong()
        ok = k32.GetExitCodeProcess(h, ctypes.byref(code))
        k32.CloseHandle(h)
        return bool(ok) and code.value == 259            # STILL_ACTIVE
    except Exception:
        return True                                      # 判断不了就当活着


def acquire_lock(out_dir, log=print):
    """同一个输出目录同时只能有一个实例在跑，否则会互相覆盖编号"""
    path = os.path.join(out_dir, LOCK_NAME)
    try:
        if os.path.exists(path):
            age = time.time() - os.path.getmtime(path)
            try:
                with open(path, encoding="utf-8") as fh:
                    info = fh.read().strip()
            except Exception:
                info = ""
            m = re.search(r"pid=(\d+)", info)
            pid = int(m.group(1)) if m else 0
            if age < 300 and pid and _pid_alive(pid) and pid != os.getpid():
                log("！这个输出目录已经有一个任务在跑了（%s，%d 秒前还在动）"
                    % (info or "?", age))
                log("  同一个目录同时只能跑一个：等它跑完，或者换个输出目录")
                log("  如果确认那个任务已经关了，把这个文件删掉再跑：%s" % path)
                return False
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("pid=%d 起于 %s"
                     % (os.getpid(),
                        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        return True
    except Exception:
        return True


def touch_lock(out_dir):
    try:
        os.utime(os.path.join(out_dir, LOCK_NAME), None)
    except Exception:
        pass


def clean_tmp(out_dir):
    """清掉上次异常退出留下的 .xlsx.tmp"""
    try:
        for name in os.listdir(out_dir):
            if name.endswith(".xlsx.tmp"):
                try:
                    os.remove(os.path.join(out_dir, name))
                except Exception:
                    pass
    except Exception:
        pass


GENERIC_COLOURS = {"multicolor", "multi color", "multi colour", "multi-colour",
                   "mixed", "assorted", "as shown", "random", "colorful",
                   "colourful", "none", ""}


def _colour_is_generic(name):
    """希音常把色块名统一写成 Multicolor，这种名字写进标题/变体里没意义"""
    n = re.sub(r"[\s\-_]+", " ", str(name or "").strip().lower())
    return n in GENERIC_COLOURS


def _show_colour(name):
    return "" if _colour_is_generic(name) else (name or "")


def existing_max_no(out_dir):
    """输出目录里已有文件夹的最大编号（文件夹名是纯数字才算）"""
    mx = 0
    try:
        for name in os.listdir(out_dir):
            if name.isdigit():
                mx = max(mx, int(name))
    except Exception:
        pass
    return mx


def next_start_no(out_dir, sheet_last_row=0):
    """这一批的起始编号 = max(输出目录已有最大号, 表格最后行号) + 1"""
    return max(existing_max_no(out_dir), int(sheet_last_row or 0)) + 1


def newest_result_file(out_dir, prefix="选品_结果_"):
    """输出目录里最新的那份汇总表（用来接着往后加）"""
    best, best_t = None, -1.0
    try:
        for name in os.listdir(out_dir):
            if name.startswith(prefix) and name.endswith(".xlsx"):
                t = os.path.getmtime(os.path.join(out_dir, name))
                if t > best_t:
                    best, best_t = name, t
    except Exception:
        pass
    return os.path.join(out_dir, best) if best else None


def max_no_in_rows(rows):
    mx = 0
    for r in rows:
        try:
            mx = max(mx, int(float(r[0])))
        except Exception:
            pass
    return mx


def _title_key(t):
    """标题指纹：只留字母数字，取前 60 位，用来判断「是不是同一个东西」"""
    s = re.sub(r"[^a-z0-9]+", "", str(t or "").lower())
    return s[:60] if len(s) >= 12 else ""


def record_history(goods_id, sku, title):
    """把抓过的商品记进全局清单，下次不管用哪个输出目录都不会再重复"""
    try:
        with open(HIST_FILE, "a", encoding="utf-8") as fh:
            fh.write("%s\t%s\t%s\n" % (goods_id or "", sku or "",
                                       _title_key(title)))
    except Exception:
        pass


def collect_seen(out_dir=None, xlsx=None, log=print):
    """收集「已经上过/选过」的商品，避免新选品重复：
       返回 (goods_id 集合, 希音sku 集合, 标题指纹 集合)
       来源：① 你的主表（所有工作表）② 输出目录里以前的汇总表 ③ 商品文件夹里的文档
    """
    goods, skus, titles = set(), set(), set()

    def add_sku(v):
        s = str(v or "").strip().lower()
        if len(s) >= 6 and re.fullmatch(r"[a-z0-9]+", s):
            skus.add(s)

    def add_url(u):
        m = re.search(r"-p-(\d+)\.html", str(u or ""))
        if m:
            goods.add(m.group(1))

    def add_title(t):
        k = _title_key(t)
        if k:
            titles.add(k)

    if xlsx and os.path.exists(xlsx):
        try:
            for s in CR.list_sheets(xlsx):
                try:
                    rows, links = CR.read_sheet(xlsx, s["name"])
                except Exception:
                    continue
                for r in rows:
                    if r["rownum"] == 1:
                        continue
                    c = r["cells"]
                    add_sku(c.get(2))
                    add_title(c.get(6))
                    add_title(c.get(7))
                    add_title(c.get(8))
                for url in (links or {}).values():
                    add_url(url)
        except Exception as exc:
            log("（读主表 %s 失败: %s）" % (os.path.basename(xlsx), exc))

    # 工具自己的全局记忆（换输出目录也不会重复）
    if os.path.exists(HIST_FILE):
        try:
            with open(HIST_FILE, encoding="utf-8", errors="ignore") as fh:
                for line in fh:
                    parts = line.rstrip("\n").split("\t")
                    if not parts:
                        continue
                    if len(parts) > 0:
                        gid = parts[0].strip()
                        if gid:
                            goods.add(gid)
                    if len(parts) > 1:
                        add_sku(parts[1])
                    if len(parts) > 2:
                        k = parts[2].strip()
                        if len(k) >= 12:
                            titles.add(k[:60])
        except Exception:
            pass

    if out_dir and os.path.isdir(out_dir):
        for name in os.listdir(out_dir):
            path = os.path.join(out_dir, name)
            if name.endswith(".xlsx") and not name.startswith("~$"):
                try:
                    for s in CR.list_sheets(path):
                        rows, links = CR.read_sheet(path, s["name"])
                        for r in rows:
                            if r["rownum"] == 1:
                                continue
                            c = r["cells"]
                            add_sku(c.get(2))
                            add_title(c.get(6))
                            add_title(c.get(7))
                            add_title(c.get(8))
                        for url in (links or {}).values():
                            add_url(url)
                except Exception:
                    pass
            elif os.path.isdir(path):
                doc = os.path.join(path, TXT)
                if os.path.exists(doc):
                    try:
                        with open(doc, encoding="utf-8", errors="ignore") as fh:
                            txt = fh.read()
                        for m in re.finditer(r"-p-(\d+)\.html", txt):
                            goods.add(m.group(1))
                        for m in re.finditer(r"sku=([A-Za-z0-9]+)", txt):
                            add_sku(m.group(1))
                        for m in re.finditer(r"^标题:\s*(.+)$", txt, re.M):
                            add_title(m.group(1))
                    except Exception:
                        pass
    return goods, skus, titles


def _sheet_images(zf, names, sheet_path):
    """取某张工作表里内嵌的图片 {表里的行号: 图片字节}"""
    rp = "xl/worksheets/_rels/%s.rels" % os.path.basename(sheet_path)
    if rp not in names:
        return {}
    dtarget = ""
    for rel in ET.fromstring(zf.read(rp)):
        if (rel.get("Type") or "").endswith("/drawing"):
            dtarget = (rel.get("Target") or "").replace("../", "")
            break
    dpath = "xl/" + dtarget if dtarget else ""
    if not dpath or dpath not in names:
        return {}
    drels = {}
    drp = "xl/drawings/_rels/%s.rels" % os.path.basename(dpath)
    if drp in names:
        for rel in ET.fromstring(zf.read(drp)):
            drels[rel.get("Id")] = (rel.get("Target") or "").replace("../", "")
    ns = {"xdr": NS_XDR, "a": NS_A}
    out = {}
    for anchor in ET.fromstring(zf.read(dpath)):
        row_el = anchor.find(".//xdr:from/xdr:row", ns)
        blip = anchor.find(".//a:blip", ns)
        if row_el is None or blip is None:
            continue
        target = drels.get(blip.get("{%s}embed" % NS_REL), "")
        media = "xl/" + target
        if target and media in names:
            try:
                out[int(row_el.text) + 1] = zf.read(media)
            except Exception:
                pass
    return out


def _num_if_number(v):
    """读回来的短数字串还原成数字；长数字串（比如 sku）保持文本，避免精度丢失"""
    if isinstance(v, str):
        s = v.strip()
        if re.fullmatch(r"-?\d{1,15}", s):
            try:
                return int(s)
            except Exception:
                return v
        if re.fullmatch(r"-?\d{1,12}\.\d{1,6}", s):
            try:
                return float(s)
            except Exception:
                return v
    return v


def read_workbook(path):
    """读回工作簿： [{name, headers, rows, images{行:字节}, links{行:url}}]"""
    out = []
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
            raise ValueError("不是有效的 xlsx 工作簿")
        rels = {}
        for rel in ET.fromstring(zf.read("xl/_rels/workbook.xml.rels")):
            rels[rel.get("Id")] = rel.get("Target") or ""
        pairs = []
        for node in ET.fromstring(zf.read("xl/workbook.xml")).iter():
            if node.tag.rsplit("}", 1)[-1] != "sheet":
                continue
            rid = node.get("{%s}id" % NS_REL) or node.get("id") or ""
            target = (rels.get(rid) or "").lstrip("/")
            if target:
                pairs.append((node.get("name") or "", target))

    for sname, target in pairs:
        spath = target if target.startswith("xl/") else "xl/" + target
        rows, links = CR.read_sheet(path, sname)
        ncol = 0
        for r in rows:
            for ci in r["cells"]:
                ncol = max(ncol, ci + 1)
        data = []
        for r in sorted(rows, key=lambda x: x["rownum"]):
            data.append([_num_if_number(r["cells"].get(i)) for i in range(ncol)])
        link_map = {}
        for ref, url in (links or {}).items():
            m = re.match(r"([A-Z]+)(\d+)$", ref or "")
            if m:
                link_map[int(m.group(2))] = url
        with zipfile.ZipFile(path) as zf:
            imgs = _sheet_images(zf, set(zf.namelist()), spath)
        out.append({"name": sname,
                    "headers": data[0] if data else [],
                    "rows": data[1:] if len(data) > 1 else [],
                    "images": imgs, "links": link_map,
                    "img_col": CAT_IMG_COL, "link_col": CAT_LINK_COL})
    return out


def base_workbook(out_dir):
    """累积用的工作簿：优先 选品汇总.xlsx，其次目录里最新的结果表，都没有就新建"""
    std = os.path.join(out_dir, "选品汇总.xlsx")
    if os.path.exists(std):
        return std
    old = newest_result_file(out_dir)
    if old:
        return old
    return std


# ---------------- 变体挑选 / 标题多选处理 ----------------
SLASH_RE = re.compile(r"(?<![0-9.])(\d{1,4}(?:\s*/\s*\d{1,4})+)(?![0-9.])")
UNIT_RE = re.compile(
    r"\s*[- ]?\s*(tiers?|layers?|pieces?|pcs|packs?|sets?|sizes?|cm|mm|"
    r"inch(?:es)?|pairs?|colou?rs?|grids?|floors?|sheets?|types?)\b", re.I)
GENERIC_UNIT = (r"(?:tiers?|layers?|pieces?|pcs|packs?|sets?|sizes?|cm|mm|"
                r"colou?rs?|pairs?|grids?|floors?|sheets?|types?)")


def _au_fix(text):
    for pat, rep in CR.SPELLING_AU:
        text = re.sub(pat, rep, text, flags=re.I)
    return text


def _disp_name(name):
    """变体名展示用：灰度拼写澳式化，One Size 规范大小写"""
    n = _au_fix(name or "").strip()
    n = re.sub(r"(?i)[\s\-_,]+one[\s\-_]*size\s*$", "", n).strip(" -,_")
    if re.fullmatch(r"(?i)one[\s\-_]*size", n):
        return "One Size"
    return n or "One Size"


def _slash_alts(title):
    alts = []
    for m in SLASH_RE.finditer(title or ""):
        for a in re.findall(r"\d+", m.group(1)):
            if a not in alts:
                alts.append(a)
    return alts


def _variant_matches_alt(name, alt):
    pat = re.compile(r"(?<![0-9])" + alt + r"\s*[- ]?\s*" + GENERIC_UNIT, re.I)
    return bool(pat.search(name or ""))


def _spec_skeleton(name):
    """变体名的"规格族"：把开头的尺码/件数 token 去掉，剩下的是族名。
    'S' -> ''（纯尺码族）
    'L Size Pot + Scraper + Oil Brush + Bamboo Rake' -> 'Size Pot + Scraper + …'
    这样"纯尺码"和"尺码+套装"不会被混在一起挑。"""
    n = (name or "").strip()
    m = re.match(r"^(?:\d+(?:\.\d+)?\s*(?:pcs|pc|pieces?|pack|sets?)?|"
                 r"XXS|XS|S|M|L|XL|XXL|[2-9]XL)\b[\s\-_/]*", n, re.I)
    return n[m.end():].strip() if m else n


def pick_variants(variants, title="", price_min=None, price_max=None,
                  want=None, min_stock=0):
    """挑几个有库存、且价格都在区间内的变体（默认最多 3 个，由界面「变体个数」控制）。
    如果标题里是多选（13/14/16/17/20/22pcs），优先挑同一个规格的变体，
    这样标题定成哪个规格、变体就是哪个规格，不会打架。"""
    lo = CR.PRICE_MIN if price_min is None else price_min
    hi_lim = CR.PRICE_MAX if price_max is None else price_max
    want = 3 if want is None else max(1, int(want))
    min_stock = int(min_stock or 0)
    notes = []
    in_stock = [v for v in variants if (v.get("stock") or 0) > 0]
    if len(in_stock) < len(variants):
        notes.append("排除 %d 个无库存变体" % (len(variants) - len(in_stock)))
    if min_stock:
        enough = [v for v in in_stock if (v.get("stock") or 0) >= min_stock]
        if len(enough) < len(in_stock):
            notes.append("排除 %d 个库存低于 %d 的变体"
                         % (len(in_stock) - len(enough), min_stock))
        in_stock = enough
    valid = [v for v in in_stock
             if v.get("price") is not None and lo <= v["price"] <= hi_lim]
    out_of_range = len(in_stock) - len(valid)
    if out_of_range > 0:
        notes.append("排除 %d 个价格不在 %g-%g 的变体"
                     % (out_of_range, lo, hi_lim))
    if not valid:
        return [], notes

    alts = _slash_alts(title)
    if alts and len(valid) > 1:
        groups = {}
        for v in valid:
            for a in alts:
                if _variant_matches_alt(v["name"], a):
                    groups.setdefault(a, []).append(v)
                    break
        if groups:
            alt = max(groups, key=lambda k: len(groups[k]))
            if len(groups) > 1:
                notes.append("标题是多选(%s)，统一按 %s 这档挑变体"
                             % ("/".join(alts), alt))
            valid = groups[alt]

    # 按"规格族"分组：不要混着挑（例如 S/M/L 和 "L Size Pot + Scraper…" 是两族）
    if len(valid) > 1:
        fams = {}
        for v in valid:
            fams.setdefault(_spec_skeleton(v["name"]), []).append(v)
        if len(fams) > 1:
            best_key = max(fams, key=lambda k: (len(fams[k]),
                                                -min(len(x["name"].split())
                                                     for x in fams[k])))
            if len(fams[best_key]) >= 2:
                notes.append("变体按规格族挑：只取「%s」这一族（共 %d 个）"
                             % (best_key or "纯尺码 S/M/L", len(fams[best_key])))
                valid = fams[best_key]

    hi = min(want, len(valid))
    n = hi                      # 够就挑满，不够就有多少挑多少
    chosen = random.sample(valid, n)
    chosen.sort(key=lambda v: v["name"])
    if n < 2:
        notes.append("只有 1 个可用变体")
    return chosen, notes


def resolve_title_options(title, vnames):
    """标题里的 2/3 Tier 这类多选：能从变体确定就定成一个值"""
    notes = []

    def sub(m):
        alts = re.findall(r"\d+", m.group(1))
        found = []
        for alt in alts:
            pat = re.compile(r"(?<![0-9])" + alt + r"\s*[- ]?\s*" + GENERIC_UNIT,
                             re.I)
            for vn in vnames:
                if pat.search(vn):
                    if alt not in found:
                        found.append(alt)
                    break
        if found:
            if len(found) > 1 and int(found[0]) > int(found[1]):
                found = [found[1], found[0]]
            return found[0]
        # 退一步：变体名里只要出现这个数字（且不是 cm/mm 这种尺寸）也算
        loose = []
        for alt in alts:
            pat = re.compile(r"(?<![0-9])" + alt + r"(?![0-9])"
                             r"(?!\s*(?:cm|mm|inch|in|kg|ml|°))", re.I)
            for vn in vnames:
                if pat.search(vn):
                    loose.append(alt)
                    break
        if len(loose) == 1:
            return loose[0]
        notes.append("多选未定 %s：选中的变体看不出来，需人工确认" % m.group(1))
        return m.group(0)

    return SLASH_RE.sub(sub, title or ""), notes


def build_title(src_title, vnames, limit=80):
    t, notes = resolve_title_options(src_title or "", vnames)
    return _trim_tail_fragment(CR.build_ebay_title(t, limit)), notes


def _colour_state(detail):
    """返回 (当前颜色, 颜色总数, 其它颜色列表)
    希音的「颜色」很多时候是一个颜色一个独立商品（各有自己的 goods_id）"""
    detail = detail or {}
    colors = detail.get("colors") or []
    cur = detail.get("color") or ""
    others = [c for c in colors
              if c.get("goods_id") and c["goods_id"] != detail.get("goods_id")]
    return cur, len(colors), others


def _colour_url(c):
    slug = "-".join((c.get("url_name") or "").split())
    if not slug or not c.get("goods_id"):
        return ""
    return "https://au.shein.com/%s-p-%s.html" % (slug, c["goods_id"])


def _colour_lines(detail):
    cur, total, others = _colour_state(detail)
    if total <= 1:
        return ["颜色: %s" % (cur or "-")] if cur else []
    return ["颜色/款式: %s（这条链接里挂了 %d 个同系列商品，希音按色块展示，"
            "每个色块是独立商品）" % (cur or "-", total)]


def build_title_for(detail, vnames, limit=80):
    """标题：先按规则生成；多颜色的商品，如果还有位置就把颜色缀到结尾"""
    src = (detail or {}).get("title") or ""
    t, notes = build_title(src, vnames, limit)
    cur, total, _ = _colour_state(detail)
    cur = _show_colour(cur)
    if total > 1 and cur and cur.lower() not in t.lower():
        if len(t) + 1 + len(cur) <= limit:
            t = "%s %s" % (t, cur)
    return t, notes


TAIL_ARTICLES = {"the", "a", "an", "this", "these", "those", "that"}
TAIL_WEAK = {"minimalist", "solid", "simple", "premium", "durable", "stylish",
             "elegant", "classic", "unique", "portable", "lightweight",
             "versatile", "aesthetic", "practical", "travel", "camping",
             "outdoor", "indoor", "daily", "fashionable", "creative"}


def _trim_tail_fragment(t):
    """截断后如果尾巴是半截从句（… Organising the These Boxes Feature），从冠词处切掉"""
    words = t.split()
    if len(words) < 5:
        return t
    for i in range(len(words) - 1, max(0, len(words) - 6), -1):
        if words[i].lower().strip(",.-;:") in TAIL_ARTICLES:
            cut = " ".join(words[:i]).strip(" -,:;.")
            if len(cut.split()) >= 3:
                return CR._strip_dangling(cut)
            break
    # 结尾一个两个虚形容词也去掉（… Quick Dry Towel Minimalist Solid）
    while len(words) > 4 and words[-1].lower().strip(",.-;:") in TAIL_WEAK:
        words.pop()
    if len(words) != len(t.split()):
        return " ".join(words).strip(" -,:;.")
    return t


def title_alternatives(src_title, limit=80):
    """多选定不下来时，把两个版本标题都算出来给人工挑"""
    out = []
    for m in SLASH_RE.finditer(src_title or ""):
        for alt in re.findall(r"\d+", m.group(1)):
            t = src_title[:m.start()] + alt + src_title[m.end():]
            fixed = CR.build_ebay_title(t, limit)
            um = UNIT_RE.match(src_title[m.end():])
            label = ("%s %s" % (alt, um.group(1)) if um else alt)
            label = label.replace("/", " ")
            out.append((label, fixed))
    return out


def _write_doc(folder, ebay_title, names, chosen, factor, url, note, imgs, alts,
               detail=None):
    lines = ["标题: %s" % ebay_title]
    lines += _colour_lines(detail)
    lines += ["变体: %s" % ", ".join(names),
              "价格: %s" % " / ".join(
                  "%s" % CR.calc_ebay_price(v["price"], factor)
                  for v in chosen if v.get("price") is not None)]
    lines.append("明细:")
    for v in chosen:
        p = v.get("price")
        ep = CR.calc_ebay_price(p, factor) if p is not None else ""
        lines.append("  %s | sku=%s | 希音 AU$%s -> eBay %s | 库存 %s"
                     % (_disp_name(v["name"]), v.get("sku_code") or "-",
                        p if p is not None else "-", ep, v.get("stock") or 0))
    if alts:
        lines.append("备选标题（标题里是多选，需人工定一个）:")
        for label, t in alts:
            lines.append("  [%s] %s" % (label, t))
    lines.append("链接: %s" % url)
    lines.append("图片: %d 张" % len(imgs))
    if note:
        lines.append("备注: %s" % note)
    with open(os.path.join(folder, TXT), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


def run(input_xlsx, sheet_name, factor, out_dir, log,
        do_scrape=True, do_images=True, limit=0,
        port=DEFAULT_PORT, captcha_wait=300, stop_event=None,
        price_min=None, price_max=None, variant_limit=3,
        min_stock=0, min_rating=0, skip_no_rating=False,
        filter_blocked=False, block_words=None, skip_local=True,
        browser_name="Chrome", browser_exe=None,
        browser_profile=None, browser_args=None):
    lo = CR.PRICE_MIN if price_min is None else float(price_min)
    hi_lim = CR.PRICE_MAX if price_max is None else float(price_max)
    rows, links = CR.read_sheet(input_xlsx, sheet_name)
    log("读到 %d 行（含表头）" % len(rows))
    os.makedirs(out_dir, exist_ok=True)
    clean_tmp(out_dir)
    if not acquire_lock(out_dir, log):
        return None, 0, 0
    # 表格链接模式也使用与分类模式相同的 17 列模板，并写入
    # 输出目录的「选品汇总.xlsx」。如果输入表本身就是这个文件，
    # 就直接更新对应工作表。
    summary_path = os.path.join(out_dir, "选品汇总.xlsx")
    prev_sheets = []
    if os.path.exists(summary_path):
        try:
            prev_sheets = read_workbook(summary_path)
            for sh in prev_sheets:
                sh["rows"] = _trim_trailing_blank_rows(sh.get("rows") or [])
        except Exception as exc:
            log("！已有选品汇总读不动（%s），本次先不接旧表" % exc)
            prev_sheets = []
    sheet_last = max((r["rownum"] for r in rows), default=1)
    summary_max = max(
        [max_no_in_rows(sh.get("rows") or []) for sh in prev_sheets] or [0])
    # 工作表名本身也代表一个区间，例如「备份35-45」即使当前只保留了
    # 部分数据行，下一批也必须从 46 开始。
    for sh in prev_sheets:
        bounds = _sheet_range_bounds(sh.get("name"))
        if bounds:
            summary_max = max(summary_max, bounds[1])
    start_no = max(next_start_no(out_dir, sheet_last), summary_max + 1)
    log("起始编号: %d（表格最后一行是 %d，输出目录/汇总表已有的最大编号接着往下排）"
        % (start_no, sheet_last))
    input_nums = []
    for rr in rows:
        if rr["rownum"] == 1:
            continue
        try:
            input_nums.append(int(float(rr["cells"].get(0))))
        except Exception:
            pass
    data_row_count = sum(1 for rr in rows if rr["rownum"] != 1)
    if str(sheet_name).startswith("备份"):
        target_sheet = sheet_name
    elif input_nums:
        target_sheet = "备份%d-%d" % (
            min(input_nums),
            max(input_nums) if len(input_nums) == data_row_count
            else min(input_nums) + data_row_count - 1)
    else:
        # 只贴链接、不填序号的手动表：自动接在当前最大编号后面。
        target_sheet = "备份%d-%d" % (
            start_no, start_no + max(0, data_row_count - 1))
    if input_nums and not str(sheet_name).startswith("备份"):
        # 输入的是单独的手动表时，优先并入已有的、覆盖该序号区间的
        # 备份工作表，避免另建「手动表格链接」表。
        lo_num, hi_num = min(input_nums), max(input_nums)
        for old in prev_sheets:
            bounds = _sheet_range_bounds(old.get("name"))
            if bounds and bounds[0] <= lo_num and hi_num <= bounds[1]:
                target_sheet = old["name"]
                break
    log("目标工作表: %s" % target_sheet)
    backup_state = {"done": not os.path.exists(summary_path)}

    def save_table():
        try:
            touch_lock(out_dir)
            if os.path.exists(summary_path) and not backup_state["done"]:
                bakdir = os.path.join(out_dir, "_修复备份")
                os.makedirs(bakdir, exist_ok=True)
                stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                shutil.copy2(summary_path,
                             os.path.join(bakdir,
                                          "选品汇总_写入前_%s.xlsx" % stamp))
                backup_state["done"] = True
            existing = next((s for s in prev_sheets
                             if s.get("name") == target_sheet), None)
            if existing:
                ordered_rows, ordered_imgs, ordered_links = _merge_summary_rows(
                    existing, out_rows, embeds, out_links)
            else:
                ordered_rows, ordered_imgs, ordered_links = _order_output_rows(
                    out_rows, embeds, out_links)
            cur = {"name": target_sheet, "headers": CAT_HEADERS,
                   "rows": ordered_rows, "images": ordered_imgs,
                   "links": ordered_links, "img_col": CAT_IMG_COL,
                   "link_col": CAT_LINK_COL}
            sheets = [s for s in prev_sheets
                      if s.get("name") != target_sheet] + [cur]
            tmp = summary_path + ".tmp"
            write_workbook(tmp, sheets)
            os.replace(tmp, summary_path)
        except Exception as exc:
            log("      (写入选品汇总失败: %s)" % exc)

    scraper = None
    if do_scrape:
        selected_exe = browser_exe or find_browser(browser_name)
        selected_profile = browser_profile or PROFILE_DIR
        if not selected_exe:
            log("！没找到 %s，本次不抓希音详情，改用表格里的数据"
                % browser_name)
            do_scrape = False
        else:
            try:
                scraper = SheinScraper(
                    selected_profile, port=port, log=log,
                    chrome=selected_exe, browser_name=browser_name,
                    extra_args=browser_args or [])
                if not scraper.start():
                    log("！%s 启动失败，本次不抓希音详情" % browser_name)
                    do_scrape = False
                else:
                    try:
                        scraper.connect()
                    except Exception as exc:
                        if is_connection_error(exc) and scraper.reconnect():
                            log("%s 初次连接断开，已自动重连" % browser_name)
                        else:
                            raise
                    log("%s 已就绪（如果出现人机验证，请在窗口里处理一下）"
                        % browser_name)
            except Exception as exc:
                log("！抓取初始化失败: %s" % exc)
                do_scrape = False

    out_rows = []
    embeds = {}
    out_links = {}
    made = skipped = 0
    n = 0
    try:
        for r in rows:
            if stop_event is not None and stop_event.is_set():
                log("收到停止指令，先把已抓到的写进表 …")
                break
            rn = r["rownum"]
            if rn == 1:
                continue
            c = r["cells"]
            seq = c.get(0)
            sku = c.get(2)
            price = CR.parse_price(c.get(4))
            variant_txt = c.get(11) or ""
            sh_title = c.get(8) or ""
            memo = c.get(12)
            url = links.get("G%d" % rn, "") or links.get("B%d" % rn, "")
            if not url:
                # 手动链接表允许把 URL 直接写在 B 列（普通文本也可以）。
                for raw in (c.get(1), c.get(6)):
                    text = str(raw or "").strip()
                    if text.startswith("http"):
                        url = text
                        break
            url_invalid = bool(url and not re.match(r"^https?://", str(url)))
            if url_invalid:
                url = ""
            if not sh_title and not sku and not url:
                continue
            if limit and n >= limit:
                log("到达设定条数 %d，停止" % limit)
                break
            n += 1
            # 先分配本行编号，后面创建商品文件夹时要用到它。
            # 之前这行放在写文件夹之后，导致通过商品触发 UnboundLocalError。
            # 表格里如果有手填序号（A 列），优先沿用；否则从输出目录接着排。
            try:
                num = int(float(seq)) if seq not in (None, "") else None
            except Exception:
                num = None
            if num is None:
                num = start_no + len(out_rows)
            log("[%d] 第%d行  %s"
                % (n, rn, ((sh_title or url) or "")[:56]))

            info = None
            chosen = []
            notes = []
            if url_invalid:
                notes.append("链接为空或不是 http/https URL，跳过")
            ebay_title = ""
            imgs = []

            conn_failed = False
            language_failed = False
            if url and do_scrape:
                for attempt in range(2):
                    try:
                        info = scraper.fetch_product(
                            url, captcha_wait=captcha_wait, stop=stop_event)
                        log("      抓到: goods_id=%s | 变体 %d | 图 %d"
                            % (info["goods_id"], len(info["variants"]),
                               len(info["images"])))
                        if info.get("error"):
                            notes.append(info["error"])
                        break
                    except NeedLanguage as exc:
                        log("！页面语言不是英文：%s" % exc)
                        language_failed = True
                        break
                    except Stopped:
                        log("收到停止指令，先把已抓到的写进表 …")
                        conn_failed = True
                        break
                    except Exception as exc:
                        if is_connection_error(exc) and attempt == 0:
                            log("      浏览器调试连接断开，正在重连并重试当前商品 …")
                            if scraper.reconnect():
                                continue
                            log("！%s 重连失败，停止本批，避免继续写失败记录"
                                % browser_name)
                            notes.append("浏览器连接断开，已停止本次")
                            conn_failed = True
                            break
                        if is_connection_error(exc):
                            log("！重试后浏览器连接仍断开，停止本批")
                            notes.append("浏览器连接断开，已停止本次")
                            conn_failed = True
                            break
                        log("      抓取失败: %s" % exc)
                        notes.append("抓取失败: %s" % exc)
                        info = None
                        break
            if conn_failed or language_failed:
                break

            title_src = (((info or {}).get("title") or sh_title) or "").strip()
            vs = (info or {}).get("variants") or []
            if vs:
                chosen, pn = pick_variants(vs, title=title_src,
                                           price_min=lo, price_max=hi_lim,
                                           want=variant_limit,
                                           min_stock=min_stock)
                notes += pn
            elif (price is not None and not min_stock and not do_scrape
                  and not url_invalid):
                # 只有用户主动关闭抓取时，才使用输入表的旧价格/变体。
                # 详情页抓取失败时不能把表格价格和另一个变体名拼在一起。
                chosen = [{"name": _disp_name(variant_txt) or "One Size",
                           "sku_code": sku or "", "price": price, "stock": 0}]

            rating = (info or {}).get("rating")
            reviews = (info or {}).get("reviews") or 0
            if filter_blocked:
                hit = CR.find_blocked(title_src, block_words)
                if hit:
                    notes.append("禁品/大牌：%s，跳过" % hit)
                    chosen = []
            if chosen and skip_local and info is not None and info.get("is_local"):
                notes.append("希音标记 Local（本地仓），跳过")
                chosen = []
            if chosen and min_rating:
                if rating is None and skip_no_rating:
                    notes.append("暂无评分，跳过")
                    chosen = []
                elif rating is not None and rating < min_rating:
                    notes.append("评分 %s < %g，跳过" % (rating, min_rating))
                    chosen = []

            status = "跳过"
            note = ""
            alts = []
            if not chosen:
                note = "; ".join(notes) or "没有价格/变体"
            else:
                cur_colour, n_colours, _others = _colour_state(info)
                cur_colour = _show_colour(cur_colour)
                names = []
                for v in chosen:
                    nm = _disp_name(v["name"])
                    names.append("%s / %s" % (cur_colour, nm) if cur_colour else nm)
                if n_colours > 1:
                    notes.append("这条链接里挂了 %d 个同系列商品（希音按色块展示）"
                                 % n_colours)
                ebay_title, tnotes = build_title_for(info, names)
                notes += tnotes
                notes += CR.check_title(ebay_title)
                prices = [v["price"] for v in chosen if v.get("price") is not None]
                if not prices or not all(lo <= p <= hi_lim for p in prices):
                    note = "变体价格不在 %g-%g 区间" % (lo, hi_lim)
                else:
                    status = "通过"
                    note = "; ".join([x for x in notes if x])
                    if any("多选未定" in x for x in notes):
                        alts = title_alternatives(title_src)
                    folder = os.path.join(out_dir, str(num))
                    try:
                        os.makedirs(folder, exist_ok=True)
                        if do_images and info and info.get("images"):
                            imgs = download_images(info["images"], folder,
                                                   IMG_MAX_PER_ITEM, log)
                            log("      图片 %d 张" % len(imgs))
                        _write_doc(folder, ebay_title, names, chosen, factor,
                                   url, note, imgs, alts, info)
                    except Exception as exc:
                        log("      写商品文件夹失败: %s" % exc)

            # 统一使用与分类模式相同的 17 列模板。
            date_txt = c.get(1)
            try:
                raw_date = float(date_txt)
                if raw_date > 20000:
                    date_txt = (datetime.datetime(1899, 12, 30) +
                                datetime.timedelta(days=raw_date)
                                ).strftime("%Y-%m-%d")
                else:
                    date_txt = str(date_txt)
            except Exception:
                date_txt = str(date_txt or "")
            if not date_txt:
                date_txt = datetime.date.today().isoformat()

            base_price = None
            ps = [v["price"] for v in chosen if v.get("price") is not None]
            if ps:
                base_price = min(ps)
            elif price is not None:
                base_price = price

            ebay_txt = ""
            if chosen:
                if ps:
                    eps = []
                    for p in ps:
                        e = CR.calc_ebay_price(p, factor)
                        if e not in eps:
                            eps.append(e)
                    ebay_txt = " / ".join("%g" % e for e in eps)
            if chosen:
                var_txt = "; ".join(_disp_name(v["name"]) for v in chosen)
            elif vs:
                # 变体被库存/价格规则过滤时，不能误写成“无”。
                var_txt = ("%s（无可上架变体）"
                           % "; ".join(_disp_name(v["name"]) for v in vs))
            else:
                var_txt = "无"
            store_txt = (info or {}).get("store") or ""
            rating_txt = (("%s（%s条）" % (rating, reviews))
                          if (rating and reviews) else
                          (str(rating) if rating else
                           ("暂无评分" if info else "")))
            stock_txt = (str(min((v.get("stock") or 0) for v in chosen))
                         if chosen else "")
            sold_txt = str((info or {}).get("sold") or "")
            input_sku = str(sku or "").strip()
            page_sku = str((info or {}).get("skc_name") or "").strip()
            sku_txt = input_sku or page_sku
            sku_note = ""
            if input_sku and page_sku and input_sku != page_sku:
                sku_note = "SKU不一致：表格=%s，页面=%s，保留表格SKU" % (
                    input_sku, page_sku)
            note_txt = note or memo or ""
            if sku_note:
                note_txt = "; ".join([x for x in (note_txt, sku_note) if x])
            if status != "通过" and not note_txt.startswith("跳过"):
                note_txt = "跳过：" + (note_txt or "未通过筛选")
            if status == "通过":
                made += 1
            else:
                # 跳过的照样写进汇总表（备注里说明原因），只是不建文件夹、不落盘
                skipped += 1
                log("      跳过：不建文件夹、不落盘")
            if status != "通过" and info and info.get("images") and not imgs:
                try:      # 表格里还是要能看到封面，只取到内存
                    imgs = [("mem", download_image(info["images"][0]))]
                except Exception as exc:
                    log("      主图下载失败: %s" % exc)
            out_rows.append([num, date_txt, sku_txt, "", base_price or "",
                             0, url, store_txt, title_src, ebay_txt, "无",
                             var_txt, note_txt, "", rating_txt, sold_txt,
                             stock_txt])
            if url:
                out_links[len(out_rows) + 1] = url
            if imgs:
                embeds[len(out_rows) + 1] = imgs[0][1]
            save_table()
    finally:
        if scraper:
            scraper.close()

    save_table()
    log("-" * 46)
    log("完成：通过 %d 条，跳过 %d 条" % (made, skipped))
    log("汇总表: %s" % summary_path)
    return summary_path, made, skipped


# ================= 模式二：从分类页自动抓商品 =================
def run_category(cat_url, factor, out_dir, log, want=10, do_images=True,
                 port=DEFAULT_PORT, captcha_wait=300,
                 scroll_rounds=10, stop_event=None, seen_xlsx=None,
                 colour_limit=3,
                 pause=(3.0, 8.0), break_every=12, break_sec=(20, 40),
                 price_min=None, price_max=None,
                 wall_action="cool", cool_min=6.0, cool_after=3, cool_long=15.0,
                 variant_limit=3, min_rating=4.0, skip_no_rating=False,
                 min_stock=10, filter_blocked=True, block_words=None,
                 skip_local=True, auto_price_filter=True, max_pages=5,
                 write_dropped=False, browser_name="Chrome",
                 browser_exe=None, browser_profile=None, browser_args=None):
    lo = CR.PRICE_MIN if price_min is None else float(price_min)
    hi_lim = CR.PRICE_MAX if price_max is None else float(price_max)
    wall_action = wall_action if wall_action in ("cool", "wait", "skip") else "cool"
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    os.makedirs(out_dir, exist_ok=True)
    clean_tmp(out_dir)
    if not acquire_lock(out_dir, log):
        return None, 0, 0
    selected_exe = browser_exe or find_browser(browser_name)
    selected_profile = browser_profile or PROFILE_DIR
    if not selected_exe:
        log("！没找到 %s，抓不了" % browser_name)
        return None, 0, 0
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    # ---- 累积工作簿：一个文件里按编号区间分工作表（备份109-125 / 备份126-150 …）----
    book = base_workbook(out_dir)
    prev_sheets = []
    if os.path.exists(book):
        try:
            prev_sheets = read_workbook(book)
            for sh in prev_sheets:
                if not str(sh.get("name") or "").startswith("备份"):
                    sh["name"] = _sheet_range_name(sh.get("rows") or [])
            log("接上已有工作簿：%s（%d 张表：%s）"
                % (os.path.basename(book), len(prev_sheets),
                   "、".join(str(s.get("name")) for s in prev_sheets)))
        except Exception as exc:
            log("！已有工作簿读不动(%s)，本次另存一个新文件" % exc)
            prev_sheets = []
            book = os.path.join(out_dir, "选品汇总_%s.xlsx" % stamp)
    mx = existing_max_no(out_dir)
    for sh in prev_sheets:
        mx = max(mx, max_no_in_rows(sh.get("rows") or []))
    start_no = mx + 1
    log("起始编号: %d" % start_no)
    log("价格区间: %g - %g（币种跟抓的站点走，澳洲站 AU$、美国站 $）" % (lo, hi_lim))
    scraper = SheinScraper(
        selected_profile, port=port, log=log, chrome=selected_exe,
        browser_name=browser_name, extra_args=browser_args or [])
    if not scraper.start():
        log("！%s 启动失败" % browser_name)
        return None, 0, 0
    batch_rows, batch_imgs, batch_links = [], {}, {}
    batch_processed_nums = []  # 本批处理过的编号；跳过项不写表，但编号仍不重复
    next_no = start_no
    dropped_rows = []          # 卡面价超区间、没打开商品页的，收尾时统一补一行
    made = skipped = 0
    state = {"book": book, "failed": False, "last": 0.0}

    def save_book(force=False):
        """把这张新表并进工作簿；节流避免每件都重写整个文件"""
        if not batch_rows:
            return
        now = time.time()
        touch_lock(out_dir)
        if not force and now - state["last"] < 12:
            return
        state["last"] = now
        ordered_rows, ordered_imgs, ordered_links = _order_output_rows(
            batch_rows, batch_imgs, batch_links)
        cur = {"name": _sheet_range_name(
                   [[n] for n in batch_processed_nums]),
               "headers": CAT_HEADERS, "rows": ordered_rows,
               "images": ordered_imgs, "links": ordered_links,
               "img_col": CAT_IMG_COL, "link_col": CAT_LINK_COL}
        for attempt in (0, 1):
            try:
                tmp = state["book"] + ".tmp"
                write_workbook(tmp, prev_sheets + [cur])
                os.replace(tmp, state["book"])
                return
            except Exception as exc:
                if attempt == 0 and not state["failed"]:
                    state["failed"] = True
                    newb = os.path.join(out_dir, "选品汇总_%s.xlsx" % stamp)
                    log("！写不进 %s（%s），本次改写到 %s"
                        % (os.path.basename(state["book"]), exc,
                           os.path.basename(newb)))
                    state["book"] = newb
                    continue
                log("      (工作簿暂存失败: %s)" % exc)
                return

    try:
        try:
            scraper.connect()
        except Exception as exc:
            if is_connection_error(exc):
                log("%s 初次连接断开，正在重新打开 …" % browser_name)
                if not scraper.reconnect():
                    log("！%s 重连失败，本次不继续" % browser_name)
                    return None, 0, 0
            else:
                raise
        log("%s 已就绪（弹人机验证的话，请你在窗口里过一下）"
            % browser_name)
        log("打开分类页找商品 …")
        # ---- 先读历史，翻页时就知道「够不够选」，不白翻 ----
        seen_goods, seen_skus, seen_titles = collect_seen(
            out_dir, seen_xlsx or CR.DEFAULT_INPUT, log)
        log("历史比对：%d 个 goods_id / %d 个 sku / %d 个标题指纹"
            % (len(seen_goods), len(seen_skus), len(seen_titles)))

        def _fresh(lst):
            return [x for x in lst
                    if str(x.get("goods_id") or "") not in seen_goods]

        # 搜索页/分类页支持价格筛选（min_price/max_price）——加上去，
        # 候选池里就基本全是区间内的商品，产出率会高很多
        eff_url = cat_url
        lower_url = (cat_url or "").lower()
        if auto_price_filter and (
                "pdsearch" in lower_url
                or "-c-" in lower_url
                or "-sc-" in lower_url
                or "/recommend/" in lower_url):
            eff_url = with_url_params(cat_url,
                                      {"min_price": "%g" % lo,
                                       "max_price": "%g" % hi_lim},
                                      drop=("min_price", "max_price"))
            log("已给链接加价格筛选：min_price=%g&max_price=%g" % (lo, hi_lim))
        product_url = re.search(r"-p-(\d+)\.html", cat_url or "")
        if product_url:
            # 用户给的是商品详情链接时，不能被当成分类页去扫描推荐卡片。
            gid = product_url.group(1)
            log("检测到商品详情链接，只抓取该商品：%s" % gid)
            items = [{"goods_id": gid, "url": cat_url, "title": "",
                      "price": None, "img": ""}]
        else:
            items = None
            # 页面卡片数不等于可上架数。必须多收集候选，给库存/评分/Local 等
            # 过滤条件留出余量，否则“要 10 个”可能只凑到 5 个就耗尽。
            candidate_target = min(max(want * 4, 40), 400)
            log("本批需要凑够 %d 个可上架商品，候选池目标至少 %d 个"
                % (want, candidate_target))
            for c_attempt in range(2):
                try:
                    items = scraper.fetch_category(
                        eff_url, scroll_rounds=scroll_rounds, max_pages=max_pages,
                        stop_after=min(max(candidate_target * 3, 120), 500),
                        log=log, stop=stop_event,
                        enough=lambda lst: len(_fresh(lst)) >= candidate_target,
                        # 列表页被拦时别干等 5 分钟（那样看起来就像卡死），最多等 2 分钟
                        captcha_wait=min(captcha_wait, 120))
                    break
                except NeedLanguage as exc:
                    log("！列表页不是英文：%s" % exc)
                    return None, 0, 0
                except NeedHuman as exc:
                    log("！列表页被希音拦住了：%s" % exc)
                    log("  建议：点「换新会话」换一个访客会话，或者歇几分钟再跑")
                    return None, 0, 0
                except Stopped:
                    log("收到停止指令，列表页加载已中止")
                    return None, 0, 0
                except Exception as exc:
                    if is_connection_error(exc) and c_attempt == 0:
                        log("列表页加载时浏览器连接断开，正在重连后再试一次 …")
                        if scraper.reconnect():
                            continue
                        log("！%s 重连失败，本次不继续" % browser_name)
                        return None, 0, 0
                    log("！列表页加载失败：%s" % exc)
                    return None, 0, 0
        log("候选池共 %d 个商品" % len(items))
        if not items:
            log("！一个商品都没抓到（页面结构变了 / 被拦了）")
            return None, 0, 0
        fresh = _fresh(items)
        if len(fresh) < len(items):
            log("跳过 %d 个已经在表里出现过的商品" % (len(items) - len(fresh)))
        if fresh:
            items = fresh
        else:
            log("！这一页看到的商品都在表里了，先照跑，备注里会标「疑似重复」")

        # 严格按页面顺序排队：页面上从左到右、从上到下，一个一个看。
        # 命中禁品/大牌、或者卡面价已知但超区间的 → 不打开商品页（收尾在表里补一行说明）；
        # 其余（区间内的 + 卡面价没抓到的）按**原顺序**进队列 —— 不做"谁先谁后"的重排。
        queue, blocked_n, out_n = [], 0, 0
        for it in items:
            hit = CR.find_blocked(it.get("title"), block_words) if filter_blocked else ""
            if hit:
                dropped_rows.append((it, it.get("price"),
                                     "禁品/大牌：%s，未抓取" % hit))
                blocked_n += 1
                continue
            p = it.get("price")
            if p is not None and not (lo <= p <= hi_lim):
                if p > hi_lim:
                    why = "卡面价 %g，高于 %g，未抓取" % (p, hi_lim)
                else:
                    why = "卡面价 %g，低于 %g（不包邮档），未抓取" % (p, lo)
                dropped_rows.append((it, p, why))
                out_n += 1
                continue
            queue.append(it)
        log("按页面顺序排队：可抓 %d 个；不抓：禁品/大牌 %d 个、卡面价超区间 %d 个"
            % (len(queue), blocked_n, out_n))
        # 「抓取个数」是"要几个能上架的"，所以这里不能只留 want 个候选，
        # 要留着足够多的候选一路往下抓（最后没凑够就说明这个池子不够）。
        queue = queue[:max(candidate_target, want * 6, 60)]
        log("本次候选 %d 个（目标是凑够 %d 个可上架的）" % (len(queue), want))

        tasks = [{"info": it, "url": it.get("url") or "", "color": "",
                  "note": ""} for it in queue]
        done = 0
        max_scan = max(want * 6, 60)
        wall_streak = 0
        language_blocked = False
        while tasks:
            if stop_event is not None and stop_event.is_set():
                log("收到停止指令，先把已抓到的写进表 …")
                break
            if made >= want:
                log("已凑够 %d 个可上架的，停止" % want)
                break
            if done >= max_scan:
                log("已经处理 %d 个候选，还没凑够 %d 个可上架的，先停" % (done, want))
                break
            task = tasks.pop(0)
            done += 1
            it = task.get("info") or {}
            url = task.get("url") or ""
            log("[%d] %s"
                % (done, str(task.get("color") or it.get("title") or url)[:52]))

            detail = None
            notes = []
            if task.get("note"):
                notes.append(task["note"])
            tries = 0
            stop_now = False
            conn_retry = False
            connection_dead = False
            while True:
                tries += 1
                try:
                    detail = scraper.fetch_product(
                        url, stop=stop_event, wall_action=wall_action,
                        captcha_wait=(25 if wall_action == "cool" else captcha_wait))
                    wall_streak = 0
                    log("      抓到: %s | 变体 %d | 图 %d | 店铺 %s"
                        % (detail["goods_id"], len(detail["variants"]),
                           len(detail["images"]), detail.get("store") or "-"))
                    break
                except NeedLanguage as exc:
                    log("！页面语言不是英文：%s" % exc)
                    language_blocked = True
                    break
                except Stopped:
                    log("收到停止指令，先把已抓到的写进表 …")
                    stop_now = True
                    break
                except NeedHuman as exc:
                    log("      被希音拦住：%s" % exc)
                    if wall_action == "cool" and tries == 1:
                        log("      按访客身份（不登录）歇 %.0f 分钟再试一次 …" % cool_min)
                        if _sleep(cool_min * 60, stop_event, log, "等待人机验证冷却"):
                            stop_now = True
                            break
                        continue
                    notes.append("希音拦住（要验证/登录），跳过")
                    wall_streak += 1
                    detail = None
                    break
                except Exception as exc:
                    if is_connection_error(exc) and not conn_retry:
                        conn_retry = True
                        log("      浏览器调试连接断开，正在重连并重试当前商品 …")
                        if scraper.reconnect():
                            log("      已重连，继续当前商品")
                            continue
                        log("！%s 重连失败，停止本批，避免继续写失败记录"
                            % browser_name)
                        notes.append("浏览器连接断开，已停止本次")
                        connection_dead = True
                        break
                    if is_connection_error(exc):
                        log("！重试后浏览器连接仍断开，停止本批")
                        notes.append("浏览器连接断开，已停止本次")
                        connection_dead = True
                        break
                    log("      抓取失败: %s" % exc)
                    notes.append("抓取失败: %s" % exc)
                    detail = None
                    break

            if stop_now or connection_dead or language_blocked:
                if connection_dead:
                    log("本次抓取已停止（浏览器连接不可用），已抓到的结果会保留")
                elif language_blocked:
                    log("本次抓取已停止（页面不是英文），请手动切换 English 后重跑")
                break
            if wall_streak >= cool_after:
                log("      连续 %d 个都被希音拦住了，先歇 %.0f 分钟（不登录就慢慢来）…"
                    % (wall_streak, cool_long))
                wall_streak = 0
                if _sleep(cool_long * 60, stop_event, log, "连续被拦冷却"):
                    break

            title = ((detail or {}).get("title") or it.get("title") or "").strip()
            store = (detail or {}).get("store") or ""
            sku = (detail or {}).get("skc_name") or str(it.get("goods_id") or "")

            # ---- 疑似重复判断（不删，只提醒）----
            gid = str((detail or {}).get("goods_id") or it.get("goods_id") or "")
            suspect = []
            if gid and gid in seen_goods:
                suspect.append("疑似重复：这个商品已经在表里")
            else:
                for c in ((detail or {}).get("colors") or []):
                    cg = str(c.get("goods_id") or "")
                    if cg and cg != gid and cg in seen_goods:
                        suspect.append("疑似重复：同系列其它颜色已经选过")
                        break
            if _title_key(title) and _title_key(title) in seen_titles:
                suspect.append("疑似重复：跟已选商品的标题很像")
            if suspect:
                notes.append("; ".join(dict.fromkeys(suspect)))

            # ---- 多颜色（一个颜色 = 一个独立商品）：自动随机再抓几个颜色 ----
            if not task.get("color") and colour_limit > 1 and detail:
                colors = detail.get("colors") or []
                others_c = [c for c in colors
                            if c.get("goods_id")
                            and c["goods_id"] != detail.get("goods_id")
                            and c["goods_id"] not in seen_goods]
                if others_c:
                    k = min(colour_limit - 1, len(others_c))
                    picks = random.sample(others_c, k) if k else []
                    notes.append("这条链接里挂了 %d 个同系列商品（希音按色块展示），"
                                 "本次自动抓了 %d 个" % (len(colors), len(picks) + 1))
                    for c in reversed(picks):
                        cu = _colour_url(c)
                        if cu:
                            tasks.insert(0, {
                                "info": {"url": cu,
                                         "title": c.get("name") or "",
                                         "price": None,
                                         "goods_id": c.get("goods_id") or ""},
                                "url": cu, "color": c.get("name") or "",
                                "note": "同系列第 %d 个色块（%s），自动抓的"
                                        % (picks.index(c) + 2,
                                           c.get("name") or "-")})

            vs = (detail or {}).get("variants") or []
            chosen, pnotes = pick_variants(vs, title=title,
                                           price_min=lo, price_max=hi_lim,
                                           want=variant_limit,
                                           min_stock=min_stock)
            notes += pnotes
            in_stock = [v for v in vs if (v.get("stock") or 0) > 0]
            all_prices = [v["price"] for v in in_stock if v.get("price")]

            status = "通过"
            primary = ""
            # 一致性校验 ①：页面里的商品 ID 和链接里的不一致 = 读到了旧页面（SPA 竞态）
            url_gid = str((detail or {}).get("url_goods_id") or "")
            page_gid = str((detail or {}).get("goods_id") or "")
            if detail is not None and url_gid and page_gid and url_gid != page_gid:
                status, primary = "跳过", ("读到旧页面（页面商品ID %s ≠ 链接 %s），跳过"
                                          % (page_gid, url_gid))
            # 一致性校验 ②：选中的 sku 必须出现在页面 JSON-LD 的变体清单里
            page_skus = set((detail or {}).get("page_skus") or [])
            if status == "通过" and page_skus and chosen:
                bad = [v.get("sku_code") for v in chosen
                       if v.get("sku_code") and v["sku_code"] not in page_skus]
                if bad:
                    status, primary = "跳过", ("页面数据不一致（%d 个 sku 不在页面清单里），跳过"
                                              % len(bad))
            # 禁品 / 大牌：命中直接跳过（优先级最高，连评分都不看）
            hit = CR.find_blocked(title, block_words) if filter_blocked else ""
            if hit:
                status, primary = "跳过", "禁品/大牌：%s，跳过" % hit
            # 评分门槛：低于 N 星直接跳过（没评分的按设置处理）
            rating = (detail or {}).get("rating")
            reviews = (detail or {}).get("reviews") or 0
            if (status == "通过" and skip_local and detail is not None
                    and detail.get("is_local")):
                status, primary = "跳过", "希音标记 Local（本地仓），跳过"
            if status == "通过" and detail is not None and min_rating:
                if rating is None and skip_no_rating:
                    status, primary = "跳过", "暂无评分，跳过"
                elif rating is not None and rating < min_rating:
                    status, primary = "跳过", "评分 %s < %g，跳过" % (rating, min_rating)
            if not chosen:
                status = "跳过"
                if all_prices and min(all_prices) > hi_lim:
                    primary = primary or "价格太高跳过"
                elif all_prices and max(all_prices) < lo:
                    primary = primary or ("价格低于 %g（不包邮档），跳过" % lo)
                elif in_stock:
                    if min_stock and max((v.get("stock") or 0) for v in in_stock) < min_stock:
                        primary = primary or ("库存不足（最多 %d < %d），跳过"
                                              % (max((v.get("stock") or 0) for v in in_stock),
                                                 min_stock))
                    else:
                        primary = primary or "在库变体没有价格"
                elif vs:
                    primary = primary or "没有在库变体"
                else:
                    primary = primary or "没抓到变体/价格"
            if primary:
                notes = [primary] + [n for n in notes if not n.startswith("排除 ")]

            prices = [v["price"] for v in chosen if v.get("price") is not None]
            if prices:
                base_price = min(prices)
            elif all_prices:
                base_price = min(all_prices)
            else:
                base_price = it.get("price")
            ebay_price = (CR.calc_ebay_price(base_price, factor)
                          if base_price is not None else "")
            cur_colour, n_colours, _others = _colour_state(detail)
            cur_colour = _show_colour(cur_colour)
            vnames = []
            for v in chosen:
                nm = _disp_name(v["name"])
                vnames.append("%s / %s" % (cur_colour, nm) if cur_colour else nm)
            if vnames:
                var_txt = "; ".join(vnames)
            elif vs:
                var_txt = ("%s（无可上架变体）"
                           % "; ".join(_disp_name(v["name"]) for v in vs))
            else:
                var_txt = "无"
            note_txt = "; ".join([x for x in notes if x])

            ebay_title, tnotes = build_title_for(detail, vnames)
            if tnotes:
                note_txt = "; ".join([x for x in (note_txt, "; ".join(tnotes)) if x])
            if status != "通过" and not note_txt.startswith("跳过"):
                note_txt = "跳过：" + (note_txt or primary or "未通过筛选")

            # 每个处理过的商品都占一个编号，避免后续批次重复。
            # 是否把跳过项写进汇总表，统一由「跳过的商品也写一行」控制。
            num = next_no
            next_no += 1
            batch_processed_nums.append(num)
            write_this_row = (status == "通过" or write_dropped)
            imgs, first_blob = [], None
            if status == "通过" and detail and detail.get("images") and do_images:
                folder = os.path.join(out_dir, str(num))
                try:
                    os.makedirs(folder, exist_ok=True)
                    imgs = download_images(detail["images"], folder,
                                           IMG_MAX_PER_ITEM, log)
                    if imgs:
                        first_blob = imgs[0][1]
                    _write_cat_doc(folder, ebay_title, vnames, chosen, factor,
                                   url, note_txt, imgs, store, tnotes, title,
                                   detail, status)
                    log("      图片 %d 张（%s）" % (len(imgs), status))
                except Exception as exc:
                    log("      写商品文件夹失败: %s" % exc)
            elif status != "通过":
                log("      跳过：不建文件夹" +
                    ("、写入汇总表留档" if write_this_row else "、未写入汇总表"))
            if first_blob is None and write_this_row:
                urls = (detail or {}).get("images") or []
                cand = urls[0] if urls else it.get("img")
                if cand:
                    try:      # 只进内存，不写盘
                        first_blob = download_image(cand)
                    except Exception as exc:
                        log("      主图下载失败: %s" % exc)
            if first_blob is None and status == "通过":
                notes.append("无图：主图下载失败")
                note_txt = "; ".join([x for x in notes if x])

            if detail:      # 抓到了真信息才记进全局清单（抓取失败的以后还能再试）
                record_history(gid, sku, title)
            if write_this_row:
                # 第 14 列「是否上架（日期）」留给人工上架时填写，脚本不自动填日期。
                batch_rows.append([num, today, sku, "",
                                   base_price if base_price is not None else "",
                                   0, url, store, title, ebay_price, "无",
                                   var_txt, note_txt, "",
                                   (("%s（%s条）" % (rating, reviews))
                                    if (rating and reviews)
                                    else (str(rating) if rating
                                          else ("暂无评分" if detail else ""))),
                                   it.get("sold") or "",
                                   ("%d" % min((v.get("stock") or 0) for v in chosen)
                                    if chosen else "")])
                row_no = len(batch_rows) + 1
                if first_blob:
                    batch_imgs[row_no] = first_blob
                if url:
                    batch_links[row_no] = url
            if status == "通过":
                made += 1
            else:
                skipped += 1
            if write_this_row:
                save_book()

            # ---- 限速：别把希音惹毛了 ----
            if tasks:
                if break_every and done % break_every == 0:
                    wait = random.uniform(*break_sec)
                    log("      已抓 %d 个，歇 %.0f 秒再继续（避免被限流）" % (done, wait))
                else:
                    wait = random.uniform(*pause) if pause else 0
                total_wait = int(wait)
                if total_wait >= 10:
                    log("      等待 %d 秒（点「停止」可立即中断）" % total_wait)
                for left in range(total_wait, 0, -1):
                    if stop_event is not None and stop_event.is_set():
                        break
                    time.sleep(1)
                    if total_wait >= 10 and (left - 1) % 10 == 0 and left > 1:
                        log("      等待中，还剩 %d 秒" % (left - 1))
    finally:
        scraper.close()

    if made < want and not tasks:
        log("本批候选已全部处理完，只凑够 %d/%d 个可上架商品"
            % (made, want))

    # 卡面价超区间/禁品的那些：没打开商品页；同样由复选框决定是否补行。
    if dropped_rows:
        skipped += len(dropped_rows)
    if dropped_rows and write_dropped:
        for it, p, why in dropped_rows:
            num = next_no
            next_no += 1
            batch_processed_nums.append(num)
            card_title = str(it.get("title") or "").strip()
            eb = CR.calc_ebay_price(p, factor) if p is not None else ""
            # 同上：未抓取行也不自动写「是否上架（日期）」。
            batch_rows.append([num, today, "", "", p if p is not None else "",
                               0, it.get("url") or "", "", card_title, eb, "无", "无",
                               "跳过：" + why, "", "", it.get("sold") or "", ""])
            if it.get("url"):
                batch_links[len(batch_rows) + 1] = it["url"]
        log("另外 %d 个卡面价超区间的已在表里补了行（备注：卡面价 x，未抓取）"
            % len(dropped_rows))
    elif dropped_rows:
        log("另有 %d 个卡面价超区间/禁品的没写进表（要留档就勾上「跳过的商品也写一行」）"
            % len(dropped_rows))

    save_book(force=True)
    log("-" * 46)
    log("完成：可上架 %d 个，跳过 %d 个（跳过项%s）"
        % (made, skipped, "已写表留档" if write_dropped else "未写进表"))
    log("工作簿: %s" % state["book"])
    return state["book"], made, skipped


def _write_cat_doc(folder, ebay_title, vnames, chosen, factor, url, note, imgs,
                   store, tnotes, shein_title, detail=None, status="通过"):
    lines = ["标题: %s" % ebay_title, "状态: %s" % status]
    if detail:
        _r, _rv = detail.get("rating"), detail.get("reviews") or 0
        lines.append("评分: %s / 5%s" % (_r if _r else "暂无",
                                         ("（%s 条评价）" % _rv) if _rv else ""))
    lines += _colour_lines(detail)
    lines += ["变体: %s" % ", ".join(vnames),
              "价格: %s" % " / ".join(
                  "%s" % CR.calc_ebay_price(v["price"], factor)
                  for v in chosen if v.get("price") is not None),
              "明细:"]
    for v in chosen:
        p = v.get("price")
        ep = CR.calc_ebay_price(p, factor) if p is not None else ""
        lines.append("  %s | sku=%s | 希音 AU$%s -> eBay %s | 库存 %s"
                     % (_disp_name(v["name"]), v.get("sku_code") or "-",
                        p if p is not None else "-", ep, v.get("stock") or 0))
    if tnotes:
        alts = title_alternatives(shein_title)
        lines.append("备选标题（标题里是多选，需人工定一个）:")
        for label, t in alts:
            lines.append("  [%s] %s" % (label, t))
    lines.append("卖家: %s" % (store or "-"))
    lines.append("链接: %s" % url)
    lines.append("图片: %d 张" % len(imgs))
    if note:
        lines.append("备注: %s" % note)
    with open(os.path.join(folder, TXT), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


# ---------------- 界面 ----------------
def main():
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox, scrolledtext
    except Exception as exc:
        print("无法加载 tkinter: %s" % exc)
        return

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("990x800")
    root.minsize(860, 600)
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)

    var_in = tk.StringVar(value=CR.DEFAULT_INPUT
                          if os.path.exists(CR.DEFAULT_INPUT) else "")
    var_out = tk.StringVar(value=CR.DEFAULT_OUTPUT
                           if os.path.exists(CR.DEFAULT_OUTPUT) else "")
    var_sheet = tk.StringVar()
    var_factor = tk.StringVar(value="2.0")
    var_limit = tk.StringVar(value="5")
    var_scrape = tk.BooleanVar(value=True)
    var_images = tk.BooleanVar(value=True)
    var_port = tk.StringVar(value=str(DEFAULT_PORT))
    var_browser = tk.StringVar(value="Edge")
    var_mode = tk.StringVar(value="cat")
    var_cat = tk.StringVar(value=DEFAULT_CAT_URL)
    var_colours = tk.StringVar(value="3")
    var_variants = tk.StringVar(value="3")
    var_minrating = tk.StringVar(value="4.0")
    var_norating = tk.BooleanVar(value=False)
    var_minstock = tk.StringVar(value="10")
    var_blocked = tk.BooleanVar(value=True)
    var_blockwords = tk.StringVar(value="")
    var_skiplocal = tk.BooleanVar(value=True)
    var_autoprice = tk.BooleanVar(value=True)
    var_maxpages = tk.StringVar(value="5")
    var_writedropped = tk.BooleanVar(value=False)
    var_pmin = tk.StringVar(value=str(int(CR.PRICE_MIN)))
    var_pmax = tk.StringVar(value=str(int(CR.PRICE_MAX)))
    var_pace = tk.StringVar(value=list(PACE_PRESETS)[0])
    var_wall = tk.StringVar(value=list(WALL_ACTIONS)[0])
    var_cool = tk.StringVar(value="6")

    txt = scrolledtext.ScrolledText(frm, height=12)
    log_q = queue.Queue()
    ui_q = queue.Queue()             # 后台读表结果，交给主线程更新控件
    state = {"running": False, "stop": None, "loading_sheets": False}

    def log(msg):
        log_q.put(str(msg))

    def pump():
        try:
            while True:
                msg = log_q.get_nowait()
                txt.insert("end", msg + "\n")
        except queue.Empty:
            pass
        try:
            while True:
                _kind, _path, _names, _err = ui_q.get_nowait()
                if _path == var_in.get():
                    if _err:
                        log("读取工作表失败: %s" % _err)
                    else:
                        cmb["values"] = _names
                        if _names:
                            keep = var_sheet.get()
                            var_sheet.set(keep if keep in _names else _names[0])
                        else:
                            var_sheet.set("")
                        log("工作表: " + (", ".join(_names) if _names else "(空)"))
        except queue.Empty:
            pass
        # 日志只保留最近 1500 行，不然跑久了 Tk 会越来越卡
        try:
            lines = int(txt.index("end-1c").split(".")[0])
            if lines > 1500:
                txt.delete("1.0", "%d.0" % (lines - 1400))
        except Exception:
            pass
        txt.see("end")
        if not state["running"] and str(btn_run["state"]) == "disabled":
            btn_run.config(state="normal")
        root.after(150, pump)

    def row(parent, label, var, browse=None, width=64):
        r = ttk.Frame(parent)
        r.pack(fill="x", pady=3)
        ttk.Label(r, text=label, width=14).pack(side="left")
        ttk.Entry(r, textvariable=var, width=width).pack(
            side="left", fill="x", expand=True)
        if browse:
            ttk.Button(r, text="浏览...", command=browse).pack(side="left", padx=4)
        return r

    def pick_in():
        p = filedialog.askopenfilename(
            title="选择输入表",
            filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        if p:
            var_in.set(p)
            load_sheets()

    def pick_out():
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            var_out.set(p)

    row(frm, "输入表", var_in, pick_in)

    r2 = ttk.Frame(frm)
    r2.pack(fill="x", pady=3)
    ttk.Label(r2, text="工作表", width=14).pack(side="left")
    cmb = ttk.Combobox(r2, textvariable=var_sheet, width=30, state="readonly")
    cmb.pack(side="left")

    def load_sheets():
        """后台读取表头：G 盘/网络盘首次访问慢时，不阻塞 Tk 主线程。"""
        path = (var_in.get() or "").strip()
        if not path:
            log("请先选择输入表")
            return
        if not os.path.exists(path):
            log("输入表不存在: %s" % path)
            return
        if state.get("loading_sheets"):
            log("正在读取工作表，请稍等 …")
            return
        state["loading_sheets"] = True
        log("正在读取工作表（网络盘可能需要几秒）…")

        def _worker():
            names, err = [], ""
            try:
                names = [s["name"] for s in CR.list_sheets(path)]
            except Exception as exc:
                err = str(exc)
            finally:
                state["loading_sheets"] = False
                ui_q.put(("sheets", path, names, err))

        threading.Thread(target=_worker, daemon=True).start()

    ttk.Button(r2, text="读取工作表", command=load_sheets).pack(side="left", padx=6)

    row(frm, "输出目录", var_out, pick_out)

    r3 = ttk.Frame(frm)
    r3.pack(fill="x", pady=3)
    ttk.Label(r3, text="抓取模式", width=14).pack(side="left")
    ttk.Radiobutton(r3, text="从希音分类/搜索页自动抓商品",
                    variable=var_mode, value="cat").pack(side="left")
    ttk.Radiobutton(r3, text="按表格里的链接抓",
                    variable=var_mode, value="table").pack(side="left", padx=12)

    row(frm, "分类/搜索页", var_cat)

    r4 = ttk.Frame(frm)
    r4.pack(fill="x", pady=3)
    ttk.Label(r4, text="定价规则", width=14).pack(side="left")
    ttk.Combobox(r4, textvariable=var_factor, values=CR.MARKUP_PRESETS,
                 width=30).pack(side="left")
    ttk.Label(r4, text="   抓取条数上限").pack(side="left")
    ttk.Entry(r4, textvariable=var_limit, width=8).pack(side="left")
    ttk.Label(r4, text="（0 = 全部）   调试端口").pack(side="left")
    ttk.Entry(r4, textvariable=var_port, width=8).pack(side="left")
    ttk.Label(r4, text="   浏览器").pack(side="left")
    ttk.Combobox(r4, textvariable=var_browser, values=("Edge", "Chrome"),
                 width=8, state="readonly").pack(side="left")

    r5 = ttk.Frame(frm)
    r5.pack(fill="x", pady=3)
    ttk.Label(r5, text="选项", width=14).pack(side="left")
    ttk.Checkbutton(r5, text="抓取希音详情（变体/库存/价格/图集）",
                    variable=var_scrape).pack(side="left")
    ttk.Checkbutton(r5, text="下载图片到商品文件夹",
                    variable=var_images).pack(side="left", padx=10)

    r6 = ttk.Frame(frm)
    r6.pack(fill="x", pady=3)
    ttk.Label(r6, text="同系列商品", width=14).pack(side="left")
    ttk.Label(r6, text="一条链接里的多个色块（每个色块是独立商品），抓").pack(side="left")
    ttk.Entry(r6, textvariable=var_colours, width=5).pack(side="left")
    ttk.Label(r6, text="个（含第一个）").pack(side="left")

    r6b = ttk.Frame(frm)
    r6b.pack(fill="x", pady=3)
    ttk.Label(r6b, text="变体个数", width=14).pack(side="left")
    ttk.Label(r6b, text="每个商品里挑的规格变体（尺寸/规格那种），固定挑").pack(side="left")
    ttk.Entry(r6b, textvariable=var_variants, width=5).pack(side="left")
    ttk.Label(r6b, text="个（不够就有多少挑多少）").pack(side="left")

    r6c = ttk.Frame(frm)
    r6c.pack(fill="x", pady=3)
    ttk.Label(r6c, text="最低评分", width=14).pack(side="left")
    ttk.Label(r6c, text="商品星级低于").pack(side="left")
    ttk.Entry(r6c, textvariable=var_minrating, width=5).pack(side="left")
    ttk.Label(r6c, text="分就不抓（0 = 不看评分）").pack(side="left")
    ttk.Checkbutton(r6c, text="没评分的也跳过",
                    variable=var_norating).pack(side="left", padx=10)

    r6d = ttk.Frame(frm)
    r6d.pack(fill="x", pady=3)
    ttk.Label(r6d, text="最低库存", width=14).pack(side="left")
    ttk.Label(r6d, text="每个变体库存至少").pack(side="left")
    ttk.Entry(r6d, textvariable=var_minstock, width=5).pack(side="left")
    ttk.Label(r6d, text="件才要（0 = 不看库存；希音显示 Only N left 就是库存紧张）").pack(side="left")

    r6e = ttk.Frame(frm)
    r6e.pack(fill="x", pady=3)
    ttk.Label(r6e, text="禁品/大牌", width=14).pack(side="left")
    ttk.Checkbutton(r6e, text="不抓（kid、精油、洗发水、LV 这类）",
                    variable=var_blocked).pack(side="left")
    ttk.Label(r6e, text="  追加禁品词(逗号分隔)").pack(side="left")
    ttk.Entry(r6e, textvariable=var_blockwords, width=28).pack(side="left")
    ttk.Checkbutton(r6e, text="  不抓 Local（本地仓）",
                    variable=var_skiplocal).pack(side="left", padx=8)

    r6f = ttk.Frame(frm)
    r6f.pack(fill="x", pady=3)
    ttk.Label(r6f, text="翻页/筛选", width=14).pack(side="left")
    ttk.Checkbutton(r6f, text="搜索页自动加价格筛选（池子更对价）",
                    variable=var_autoprice).pack(side="left")
    ttk.Label(r6f, text="  最多翻").pack(side="left")
    ttk.Entry(r6f, textvariable=var_maxpages, width=5).pack(side="left")
    ttk.Label(r6f, text="页（1 = 只抓当前页）").pack(side="left")
    ttk.Checkbutton(r6f, text="  跳过的商品也写一行（含未抓取）",
                    variable=var_writedropped).pack(side="left", padx=8)

    r7 = ttk.Frame(frm)
    r7.pack(fill="x", pady=3)
    ttk.Label(r7, text="价格区间", width=14).pack(side="left")
    ttk.Entry(r7, textvariable=var_pmin, width=8).pack(side="left")
    ttk.Label(r7, text=" - ").pack(side="left")
    ttk.Entry(r7, textvariable=var_pmax, width=8).pack(side="left")
    ttk.Label(r7, text="   快捷：").pack(side="left")
    cmb_site = ttk.Combobox(r7, width=20, state="readonly",
                            values=["澳洲站 9-30（AU$）", "美国站 10-35（$）"])
    cmb_site.pack(side="left")

    def apply_site(_evt=None):
        if cmb_site.current() == 1:
            var_pmin.set("10")
            var_pmax.set("35")
        elif cmb_site.current() == 0:
            var_pmin.set("9")
            var_pmax.set("30")

    cmb_site.bind("<<ComboboxSelected>>", apply_site)

    r8 = ttk.Frame(frm)
    r8.pack(fill="x", pady=3)
    ttk.Label(r8, text="抓取节奏", width=14).pack(side="left")
    ttk.Combobox(r8, textvariable=var_pace, state="readonly", width=24,
                 values=list(PACE_PRESETS)).pack(side="left")
    ttk.Label(r8, text="（放慢就不容易被弹验证；想快就选「快」）").pack(side="left")

    r9 = ttk.Frame(frm)
    r9.pack(fill="x", pady=3)
    ttk.Label(r9, text="被拦处理", width=14).pack(side="left")
    ttk.Combobox(r9, textvariable=var_wall, state="readonly", width=30,
                 values=list(WALL_ACTIONS)).pack(side="left")
    ttk.Label(r9, text="  歇").pack(side="left")
    ttk.Entry(r9, textvariable=var_cool, width=5).pack(side="left")
    ttk.Label(r9, text="分钟再试（访客身份，不登录）").pack(side="left")

    def on_run():
        try:
            if not var_out.get():
                messagebox.showerror(APP_TITLE, "请选择输出目录")
                return
            factor = CR.normalize_price_rule(var_factor.get())
            if factor is None:
                messagebox.showerror(APP_TITLE, "定价选项无效")
                return
            if isinstance(factor, float) and factor <= 1:
                messagebox.showerror(APP_TITLE, "系数必须大于 1")
                return
            limit = int(float(var_limit.get() or 0))
            port = int(var_port.get() or DEFAULT_PORT)
            pmin = float(var_pmin.get() or CR.PRICE_MIN)
            pmax = float(var_pmax.get() or CR.PRICE_MAX)
            if pmax <= pmin:
                messagebox.showerror(APP_TITLE, "价格上限要大于下限")
                return
            mode = var_mode.get()
            if mode == "table":
                if not var_in.get() or not os.path.exists(var_in.get()):
                    messagebox.showerror(APP_TITLE, "请选择正确的输入表")
                    return
                if not var_sheet.get():
                    messagebox.showerror(APP_TITLE, "请先读取并选择工作表")
                    return
            else:
                if not (var_cat.get() or "").startswith("http"):
                    messagebox.showerror(APP_TITLE, "请填希音分类页/搜索页链接")
                    return
            browser = browser_settings(var_browser.get())
            if not browser["exe"]:
                messagebox.showerror(
                    APP_TITLE, "没找到 %s，请先安装或改选另一个浏览器"
                    % browser["name"])
                return
            if browser["name"] == "Edge" and not prepare_edge_profile(log):
                messagebox.showerror(
                    APP_TITLE,
                    "无法同步 Edge 档案。请先关闭 Edge 后再试。")
                return
            log("=" * 46)
            log("开始运行  %s" % datetime.datetime.now().strftime("%H:%M:%S"))
            log("浏览器: %s" % browser["name"])
            if browser["name"] == "Edge":
                log("Edge 使用当前 Default 档案；若 Edge 已打开，会先关闭再以调试模式重启")
            if mode == "table":
                log("模式: 按表格链接抓")
                log("输入: %s" % var_in.get())
                log("工作表: %s   定价: %s   条数: %s"
                    % (var_sheet.get(), var_factor.get(), limit or "全部"))
            else:
                log("模式: 从分类页自动抓")
                log("分类页: %s" % var_cat.get())
                log("个数: %s   定价: %s" % (limit or 10, var_factor.get()))
            log("输出: %s" % var_out.get())
            stop_event = threading.Event()
            state["stop"] = stop_event
            state["running"] = True
            btn_run.config(state="disabled")

            def worker():
                try:
                    if mode == "table":
                        run(var_in.get(), var_sheet.get(), factor, var_out.get(),
                            log, do_scrape=bool(var_scrape.get()),
                            do_images=bool(var_images.get()),
                            limit=limit, port=port, stop_event=stop_event,
                            price_min=pmin, price_max=pmax,
                            variant_limit=max(1, int(var_variants.get() or 3)),
                            min_stock=int(float(var_minstock.get() or 0)),
                            min_rating=float(var_minrating.get() or 0),
                            skip_no_rating=bool(var_norating.get()),
                            filter_blocked=bool(var_blocked.get()),
                            block_words=[w.strip() for w in
                                         (var_blockwords.get() or "").split(",")
                                         if w.strip()],
                            skip_local=bool(var_skiplocal.get()),
                            browser_name=browser["name"],
                            browser_exe=browser["exe"],
                            browser_profile=browser["profile"],
                            browser_args=browser["args"])
                    else:
                        pace_pause, pace_every, pace_break = PACE_PRESETS.get(
                            var_pace.get(), list(PACE_PRESETS.values())[0])
                        run_category(var_cat.get(), factor, var_out.get(), log,
                                     want=limit or 10,
                                     do_images=bool(var_images.get()),
                                     port=port, stop_event=stop_event,
                                     seen_xlsx=(var_in.get() or CR.DEFAULT_INPUT),
                                     colour_limit=max(1, int(var_colours.get() or 3)),
                                     price_min=pmin, price_max=pmax,
                                     pause=pace_pause, break_every=pace_every,
                                     break_sec=pace_break,
                                     wall_action=WALL_ACTIONS.get(var_wall.get(), "cool"),
                                     cool_min=float(var_cool.get() or 6),
                                     variant_limit=max(1, int(var_variants.get() or 3)),
                                     min_rating=float(var_minrating.get() or 0),
                                     skip_no_rating=bool(var_norating.get()),
                                     min_stock=int(float(var_minstock.get() or 0)),
                                     filter_blocked=bool(var_blocked.get()),
                                     block_words=[w.strip() for w in
                                                  (var_blockwords.get() or "").split(",")
                                                  if w.strip()],
                                     skip_local=bool(var_skiplocal.get()),
                                     auto_price_filter=bool(var_autoprice.get()),
                                     max_pages=max(1, int(float(var_maxpages.get() or 1))),
                                     write_dropped=bool(var_writedropped.get()),
                                     browser_name=browser["name"],
                                     browser_exe=browser["exe"],
                                     browser_profile=browser["profile"],
                                     browser_args=browser["args"])
                except Exception:
                    log("出错了:\n" + traceback.format_exc())
                finally:
                    state["running"] = False
                    log("—— 本次结束 ——")

            threading.Thread(target=worker, daemon=True).start()
        except Exception:
            log("出错了:\n" + traceback.format_exc())

    def on_stop():
        if state["running"] and state["stop"] is not None:
            state["stop"].set()
            log(">>> 正在停止，等当前这个商品收尾 …")
        else:
            log("现在没有在跑")

    def on_force_stop():
        """强制停止：立刻停 + 把输出目录的锁删掉，不用等 5 分钟"""
        if state["running"] and state["stop"] is not None:
            state["stop"].set()
        lock = os.path.join(var_out.get() or ".", LOCK_NAME)
        try:
            if var_out.get() and os.path.exists(lock):
                os.remove(lock)
                log(">>> 已强制停止并解锁（删掉 %s）" % lock)
            else:
                log(">>> 已发送停止指令（没找到锁文件，不用解锁）")
        except Exception as exc:
            log("！解锁失败（%s）——请手动删掉：%s" % (exc, lock))
        log("    提醒：如果上一个任务其实还在跑，两边会同时写同一个目录")

    def on_close():
        if state["running"]:
            if not messagebox.askokcancel(
                    APP_TITLE, "还在跑，确定要关掉吗？\n"
                               "（已经抓到的商品都写进表了，可以直接关）"):
                return
        root.destroy()

    def on_new_session():
        """换新会话：把 Chrome 档案改名。下次运行就是全新访客会话，
        希音的「要登录」墙一般就没了（不用登录、也不影响已有结果）"""
        if state["running"]:
            log("！先停掉正在跑的任务，再换会话")
            return
        if var_browser.get() == "Edge":
            messagebox.showinfo(
                APP_TITLE,
                "Edge 模式使用当前 Edge 档案，不能让脚本重置它。\n"
                "如果只是遇到验证，请在弹出的 Edge 窗口里手动处理。")
            return
        if not messagebox.askokcancel(
                APP_TITLE, "会关掉脚本打开的 Chrome，并把 _chrome_profile 改名\n"
                           "（下次运行是全新访客会话）。确定吗？"):
            return
        if state.get("switching"):
            log("正在换会话，请稍等 …")
            return
        state["switching"] = True
        port = int(var_port.get() or DEFAULT_PORT)
        log("正在关闭旧 Chrome 会话 …")

        def _switch_worker():
            try:
                try:
                    import json as _json
                    import urllib.request as _rq
                    import websocket
                    v = _json.loads(_rq.urlopen(
                        "http://127.0.0.1:%d/json/version" % port, timeout=4).read())
                    ws = websocket.create_connection(v["webSocketDebuggerUrl"],
                                                     timeout=10, max_size=1 << 20)
                    ws.send(_json.dumps({"id": 1, "method": "Browser.close"}))
                    time.sleep(3)
                except Exception:
                    pass
                if os.path.isdir(PROFILE_DIR):
                    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    newp = "%s_旧_%s" % (PROFILE_DIR, stamp)
                    os.rename(PROFILE_DIR, newp)
                    log(">>> 已换新会话：旧档案改名成 %s" % os.path.basename(newp))
                    log("    下次运行就是全新访客会话，不用登录")
                else:
                    log(">>> 没有旧档案，下次运行会新建一个访客会话")
            except Exception as exc:
                log("！换会话失败：%s（多半是 Chrome 还开着，先手动关掉再试）" % exc)
            finally:
                state["switching"] = False

        threading.Thread(target=_switch_worker, daemon=True).start()

    btns = ttk.Frame(frm)
    btns.pack(side="bottom", fill="x", pady=(6, 0))
    btn_run = ttk.Button(btns, text="开始运行", command=on_run)
    btn_run.pack(side="left")
    ttk.Button(btns, text="停止", command=on_stop).pack(side="left", padx=8)
    ttk.Button(btns, text="强制停止", command=on_force_stop).pack(side="left")
    ttk.Button(btns, text="换新会话", command=on_new_session).pack(side="left", padx=8)

    def on_open_out():
        d = var_out.get()
        if d and os.path.isdir(d):
            open_folder(d)
        else:
            log("输出目录还不存在：%s" % (d or "(空)"))

    ttk.Button(btns, text="打开输出目录",
               command=on_open_out).pack(side="left")
    ttk.Button(btns, text="退出", command=on_close).pack(side="left", padx=8)
    root.protocol("WM_DELETE_WINDOW", on_close)

    # 日志框最后才 pack，窗口小的时候先挤日志，按钮永远看得见
    txt.pack(fill="both", expand=True, pady=8)

    log(APP_TITLE)
    log("Edge: %s" % (find_edge() or "未找到"))
    log("Chrome: %s" % (find_chrome() or "未找到"))
    try:
        import websocket  # noqa
        log("websocket-client: 已安装")
    except Exception:
        log("websocket-client: 未安装！请执行  pip install websocket-client")
    try:
        import PIL  # noqa
        log("Pillow: 已安装（WebP 会自动转 JPG 写入表格）")
    except Exception:
        log("Pillow: 未安装！WebP 图片可能显示空白，请执行 pip install -r requirements.txt")
    if var_in.get():
        load_sheets()
    root.after(150, pump)
    root.mainloop()


if __name__ == "__main__":
    if "--where" in sys.argv:
        # 自检：打包后数据目录在哪（窗口模式没有控制台，写到文件里）
        lines = ["程序目录: %s" % HERE,
                 "Edge 档案: %s" % EDGE_PROFILE_DIR,
                 "Chrome 档案: %s" % PROFILE_DIR,
                 "全局记录: %s" % HIST_FILE,
                 "运行方式: %s" % ("已编译 exe" if (
                     getattr(sys, "frozen", False) or "__compiled__" in globals())
                     else "源码"),
                 "sys.executable: %s" % sys.executable,
                 "sys.argv[0]: %s" % sys.argv[0],
                 "cwd: %s" % os.getcwd()]
        try:
            with open(os.path.join(HERE, "_where.txt"), "w", encoding="utf-8") as fh:
                fh.write("\n".join(lines) + "\n")
        except Exception:
            with open("_where.txt", "w", encoding="utf-8") as fh:
                fh.write("\n".join(lines) + "\n")
        sys.exit(0)
    main()
