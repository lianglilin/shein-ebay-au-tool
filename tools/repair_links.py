# -*- coding: utf-8 -*-
"""把汇总表「希音链接」列从标题改回真实 URL，同时保留原超链接目标。

用法：
    python tools/repair_links.py "D:\\path\\to\\选品汇总.xlsx"
    python tools/repair_links.py "D:\\path\\to\\选品汇总.xlsx" --dry-run
"""
import datetime
import os
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import ebay_tool as T  # noqa: E402


def _trim(rows):
    while rows and all(v in (None, "") for v in rows[-1]):
        rows.pop()
    return rows


def repair(path, dry_run=False):
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    sheets = T.read_workbook(path)
    changed = 0
    for sh in sheets:
        rows = _trim(sh.get("rows") or [])
        links = dict(sh.get("links") or {})
        for excel_row, row in enumerate(rows, 2):
            url = str(links.get(excel_row) or "").strip()
            if not url:
                continue
            if len(row) <= 6:
                row.extend([""] * (7 - len(row)))
            if row[6] != url:
                row[6] = url
                changed += 1
        sh["rows"] = rows
        sh["links"] = links
        print("%s：%d 行，%d 个超链接" % (sh["name"], len(rows), len(links)))

    print("准备修正链接列：%d 行" % changed)
    if dry_run:
        print("dry-run，不写文件")
        return 0

    root = os.path.dirname(path)
    bakdir = os.path.join(root, "_修复备份")
    os.makedirs(bakdir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = os.path.join(bakdir, "选品汇总_链接修复前_%s.xlsx" % stamp)
    shutil.copy2(path, bak)
    print("已备份：%s" % bak)

    tmp = path + ".link.tmp"
    try:
        T.write_workbook(tmp, sheets)
        check = T.read_workbook(tmp)
        for old, new in zip(sheets, check):
            old_rows = _trim(old.get("rows") or [])
            new_rows = _trim(new.get("rows") or [])
            if [list(r) for r in old_rows] != [list(r) for r in new_rows]:
                raise RuntimeError("工作表 %s 数据发生变化，取消替换" % old["name"])
            for row_no, url in (old.get("links") or {}).items():
                if str(new.get("links", {}).get(row_no) or "") != str(url):
                    raise RuntimeError("工作表 %s 第 %d 行超链接丢失"
                                       % (old["name"], row_no))
        os.replace(tmp, path)
        print("修复完成：%s" % path)
        return 0
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--dry-run"]
    if not args:
        print(__doc__)
        sys.exit(1)
    try:
        sys.exit(repair(args[0], "--dry-run" in sys.argv[1:]))
    except Exception as exc:
        print("修复失败：%s" % exc)
        sys.exit(1)
