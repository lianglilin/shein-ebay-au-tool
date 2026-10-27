# -*- coding: utf-8 -*-
r"""
开发小工具：把一份选品结果（另一个输出目录）并进目标输出目录。

会做三件事：
  1) 源工作簿里的每张表，行号统一往后偏移，并进目标工作簿（新表名 = 备份<起>-<止>）
  2) 源目录里的商品文件夹按同样的偏移改名搬过去（目标已有同号就跳过）
  3) 源工作簿备份到 <目标>\_来源备份\，原名不变、加个标签

用法：
    python tools/merge_results.py <源目录> <目标目录> <标签>
    例：python tools/merge_results.py "D:\\测试7_AU50" "D:\\选品" AU50
"""
import os
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import ebay_tool as T          # noqa: E402


def merge(src_dir, dst_dir, label="merge"):
    os.makedirs(dst_dir, exist_ok=True)
    src_book = None
    for name in sorted(os.listdir(src_dir)):
        if name.endswith(".xlsx") and not name.startswith("~$"):
            src_book = os.path.join(src_dir, name)
            break
    if not src_book:
        print("！%s 里没找到工作簿" % src_dir)
        return 1

    dst_book = os.path.join(dst_dir, "选品汇总.xlsx")
    dst_sheets = T.read_workbook(dst_book) if os.path.exists(dst_book) else []
    off = T.existing_max_no(dst_dir)
    for sh in dst_sheets:
        off = max(off, T.max_no_in_rows(sh.get("rows") or []))

    new_sheets = []
    for sh in T.read_workbook(src_book):
        if not sh.get("rows"):
            continue
        rows = []
        for r in sh["rows"]:
            r = list(r)
            try:
                r[0] = int(float(r[0])) + off
            except Exception:
                pass
            rows.append(r)
        new_sheets.append({"name": T._sheet_range_name(rows),
                           "headers": sh["headers"], "rows": rows,
                           "images": sh["images"], "links": sh["links"],
                           "img_col": T.CAT_IMG_COL, "link_col": T.CAT_LINK_COL})
    if not new_sheets:
        print("！%s 的工作簿是空的" % label)
        return 1

    T.write_workbook(dst_book, dst_sheets + new_sheets)
    moved = skipped = 0
    for name in sorted(os.listdir(src_dir)):
        p = os.path.join(src_dir, name)
        if os.path.isdir(p) and name.isdigit():
            tgt = os.path.join(dst_dir, str(int(name) + off))
            if os.path.exists(tgt):
                skipped += 1
                continue
            shutil.move(p, tgt)
            moved += 1
    bak = os.path.join(dst_dir, "_来源备份")
    os.makedirs(bak, exist_ok=True)
    shutil.move(src_book, os.path.join(
        bak, os.path.basename(src_book).replace(".xlsx", "_%s.xlsx" % label)))
    mx = max([T.max_no_in_rows(s["rows"]) for s in new_sheets] or [0])
    print("合并 %s：新表 %s | 编号 %d-%d | 搬了 %d 个文件夹，跳过 %d 个"
          % (label, [s["name"] for s in new_sheets], off + 1, mx, moved, skipped))
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    sys.exit(merge(sys.argv[1], sys.argv[2],
                   sys.argv[3] if len(sys.argv) > 3 else "merge"))
