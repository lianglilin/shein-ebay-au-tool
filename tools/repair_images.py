# -*- coding: utf-8 -*-
"""修复汇总表里缺失或显示空白的商品图片。

用途：
  1) 每个商品文件夹里取第一张图片，重新写入汇总表对应的「产品图片」列；
  2) WebP 通过 ebay_tool.image_for_embed 转成 JPG 再嵌入，避免 Excel 显示空白；
  3) 不改商品数据、编号、备注和日期列。

用法：
    python tools/repair_images.py "D:\\path\\to\\选品汇总.xlsx"
    python tools/repair_images.py "D:\\path\\to\\选品汇总.xlsx" --dry-run
"""
import datetime
import os
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import ebay_tool as T  # noqa: E402


IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".webp")


def _trim_rows(rows):
    """去掉工作簿尾部 XML 留下的空行，保留中间的空白记录。"""
    while rows and all(v in (None, "") for v in rows[-1]):
        rows.pop()
    return rows


def _first_image(folder):
    if not os.path.isdir(folder):
        return ""
    names = sorted(
        name for name in os.listdir(folder)
        if name.lower().endswith(IMAGE_EXTS)
        and os.path.isfile(os.path.join(folder, name))
    )
    return os.path.join(folder, names[0]) if names else ""


def repair(path, dry_run=False):
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    root = os.path.dirname(path)
    sheets = T.read_workbook(path)
    changed = 0

    for sh in sheets:
        rows = _trim_rows(sh.get("rows") or [])
        images = dict(sh.get("images") or {})
        for excel_row, row in enumerate(rows, 2):
            try:
                num = int(float(row[0]))
            except Exception:
                continue
            image_path = _first_image(os.path.join(root, str(num)))
            if not image_path:
                # 106/107 这种没有商品文件夹的记录保留原样，不伪造图片。
                continue
            with open(image_path, "rb") as fh:
                blob = fh.read()
            embedded = T.image_for_embed(blob)
            if T.ext_of(embedded) not in (".jpg", ".jpeg", ".png"):
                print("跳过：%s 的图片无法转成 Excel 兼容格式" % image_path)
                continue
            images[excel_row] = embedded
            changed += 1
        sh["rows"] = rows
        sh["images"] = images
        print("%s：%d 行，%d 张图" % (sh["name"], len(rows), len(images)))

    print("准备重写图片：%d 行" % changed)
    if dry_run:
        print("dry-run，不写文件")
        return 0

    backup_dir = os.path.join(root, "_修复备份")
    os.makedirs(backup_dir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = os.path.join(backup_dir, "选品汇总_修复前_%s.xlsx" % stamp)
    shutil.copy2(path, backup)
    print("已备份：%s" % backup)

    tmp = path + ".repair.tmp"
    try:
        T.write_workbook(tmp, sheets)
        # 重新读取校验：数据行不变，图片不再是 WebP。
        check = T.read_workbook(tmp)
        if [x["name"] for x in check] != [x["name"] for x in sheets]:
            raise RuntimeError("工作表名称发生变化，取消替换")
        for old, new in zip(sheets, check):
            old_rows = _trim_rows(old.get("rows") or [])
            new_rows = _trim_rows(new.get("rows") or [])
            if [list(r) for r in old_rows] != [list(r) for r in new_rows]:
                raise RuntimeError("工作表 %s 的数据发生变化，取消替换" % old["name"])
            for row_no, blob in (new.get("images") or {}).items():
                if T.ext_of(blob) == ".webp":
                    raise RuntimeError("工作表 %s 第 %d 行仍是 WebP" %
                                       (new["name"], row_no))
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
        sys.exit(repair(args[0], dry_run="--dry-run" in sys.argv[1:]))
    except Exception as exc:
        print("修复失败：%s" % exc)
        sys.exit(1)
