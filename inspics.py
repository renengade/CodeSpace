"""
图片批量插入 Excel 工具 — inspics.py
=====================================

功能说明
  将指定文件夹中的图片（.png/.jpg/.jpeg/.gif/.bmp）按文件名中的序号排序，
  批量插入 Excel：
    - 「图片」sheet：按序号依次排列图片，每张图片上方标注图片名称
    - 「汇总表」sheet：单独展示每张图片的位置、序号及图片名称

  文件名编号提取规则：取文件名开头的编号段（如 "01-01"），
  名称取编号后的文本部分（如 "01-01 测试图.png" → 编号 01-01、名称 测试图）；
  无编号的文件将被跳过并警告。

用法示例
  python inspics.py                          # 处理脚本所在目录
  python inspics.py --path D:\\photos          # 处理指定目录
  python inspics.py --path D:\\photos --visible --picwidth 15 --interval 4

依赖
  pip install xlwings pandas Pillow
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Sequence

import pandas as pd
import xlwings as xw
from PIL import Image

PIC_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp")


# ====================================================================
#  命令行参数解析
# ====================================================================

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="inspics.py",
        description="将文件夹中的图片按序号批量插入 Excel，汇总表单独放在一个 sheet",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例: python inspics.py --path D:/photos --picwidth 15 --interval 4",
    )
    parser.add_argument(
        "--path", "-p",
        default=None,
        help="图片所在文件夹（默认：脚本所在目录）",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="显示 Excel 窗口（默认后台运行）",
    )
    parser.add_argument(
        "--picwidth",
        type=float,
        default=20.0,
        help="图片宽度，单位 cm（默认 20）",
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=6,
        help="相邻图片之间的间隔行数（默认 6）",
    )
    return parser


# ====================================================================
#  工具函数
# ====================================================================

def _split_seq(filename: str) -> tuple[tuple[int, ...], str, str] | None:
    """从文件名中拆分编号和名称文本。

    编号 = 文件名开头的数字段（如 "01-01"），
    名称 = 编号之后、扩展名之前的文本部分（如 "测试图"）。
    无编号时返回 None。

    例:
      "01-01 测试图.png" -> ((1, 1), "01-01", "测试图")
      "02-系统截图.jpg"  -> ((2,), "02", "系统截图")
      "现场照片.png"     -> None
    """
    m = re.match(r"\d+(?:-\d+)*", filename)
    if not m:
        return None
    seq_str = m.group()
    sort_key = tuple(int(x) for x in seq_str.split("-"))
    rest = Path(filename).stem[m.end():]
    name = rest.lstrip("- _.").strip()
    return sort_key, seq_str, name


def _collect_images(folder: Path) -> list[tuple[tuple[int, ...], str, str, Path]]:
    """收集文件夹中带编号的图片文件，按编号排序。

    Returns:
        按编号升序排列的 [(排序键, 编号串, 名称文本, 文件路径), ...]，
        无编号的图片跳过并打印警告。
    """
    images: list[tuple[tuple[int, ...], str, str, Path]] = []
    for f in folder.iterdir():
        if not (f.is_file() and f.suffix.lower() in PIC_EXTENSIONS):
            continue
        parsed = _split_seq(f.name)
        if parsed is None:
            print(f"跳过（文件名无编号）: {f.name}")
            continue
        images.append((parsed[0], parsed[1], parsed[2], f))

    images.sort(key=lambda t: (t[0], t[3].name.lower()))
    return images


# ====================================================================
#  主流程
# ====================================================================

def insert_images_to_excel(
    folder: Path,
    visible: bool = False,
    picwidth: float = 20.0,
    interval: int = 6,
) -> None:
    """将 folder 中带序号的图片按序插入 Excel。

    Args:
        folder:   图片所在文件夹（输出文件也保存在这里）
        visible:  是否显示 Excel 窗口
        picwidth: 图片宽度（cm），高度按原图比例缩放
        interval: 相邻图片之间的间隔行数
    """
    images = _collect_images(folder)
    if not images:
        print("未找到带序号的图片文件，程序退出")
        return

    app = xw.App(visible=visible, add_book=False)
    wb = app.books.add()
    ws1 = wb.sheets.active                # 「图片」sheet
    ws1.range("a:a").row_height = 15      # 15 磅
    ws2 = wb.sheets.add("汇总表", after=ws1)  # 汇总表单独一个 sheet，放在图片之后

    # === 插入图片并记录每张图的实际位置 ===
    row = 6  # 图片起始行
    results: list[tuple[str, str, str]] = []  # (位置, 序号, 图片名称)

    for _, seq_str, pic_name, image_path in images:
        # 图片上方一行标注图片名称
        ws1.range(f"a{row - 1}").value = pic_name

        img = Image.open(image_path)
        h, w = img.height, img.width

        # 28.35 像素 = 1cm；宽度按 picwidth，高度等比缩放
        width_px = 28.35 * picwidth
        ws1.pictures.add(
            str(image_path),
            name=image_path.stem,
            update=True,
            left=0,
            top=ws1.range(f"a{row}").top,
            width=width_px,
            height=width_px * h / w,
        )

        # 图片插入后占用的行数：像素 → cm → 磅 → 行
        pic_height = ws1.pictures[-1].height
        rows_used = math.ceil(pic_height * (72 / 2.54) / 28.35 / 15)

        results.append((f"A{row}", seq_str, pic_name))
        row += rows_used + interval

    # === 在独立的「汇总表」sheet 中写入汇总表 ===
    ws2.range("a1").value = ["位置", "序号", "图片名称"]
    ws2.range("a1:c1").font.bold = True
    # 序号列按文本格式写入，避免 "01-01" 被 Excel 识别为日期
    ws2.range(f"b2:b{1 + len(results)}").number_format = "@"
    for i, (addr, seq, name) in enumerate(results):
        row = 2 + i
        ws2.range(f"a{row}").value = [addr, seq, name]
        # 序号单元格添加内部超链接：点击跳转到图片所在位置
        cell = ws2.range(f"b{row}")
        cell.api.Hyperlinks.Add(
            Anchor=cell.api,
            Address="",
            SubAddress=f"'{ws1.name}'!{addr}",
        )
    ws2.autofit()

    # === 保存输出（输出文件与图片同目录） ===
    date = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = folder / f"output{date}.xlsx"
    wb.save(str(out_path))

    summary = pd.DataFrame(results, columns=["位置", "序号", "图片名称"])
    print(summary)
    print(f"export successful! -> {out_path}")

    wb.close()
    app.quit()


# ====================================================================
#  程序入口
# ====================================================================

def main(argv: Sequence[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)

    folder = (
        Path(args.path).resolve()
        if args.path
        else Path(__file__).resolve().parent
    )
    if not folder.is_dir():
        print(f"错误：目录不存在 - {folder}")
        return 1

    insert_images_to_excel(
        folder=folder,
        visible=args.visible,
        picwidth=args.picwidth,
        interval=args.interval,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
