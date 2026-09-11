"""
IT 审计工作底稿生成工具 — mwp.py
===================================

功能说明
   基于模板和配置信息，批量生成 IT 审计工作底稿（.xlsx）。

   1. 读取 FileName.xlsx 中的支持文档清单，自动创建对应目录结构。
   2. 打开 ITaudit_v1/ 中的模板底稿，将占位信息替换为实际参数：
      - 文件名中的 XXXX → 客户简称，期间 → 审计年度
      - 所有 sheet 中的「审计期间」替换为实际审计期间
      - 所有 sheet 中的「编制/日期」替换为编制人/编制日期
      - 所有 sheet 中的「复核/日期」替换为复核人/复核日期
      - 所有替换内容的字体统一改为 Arial
   3. 在 ITGC 底稿（SD/CM/OM/SA）的内容 sheet 中，找到「支持文档」
      单元格，在其右侧列向下填入对应的支持文档清单。
   4. 在 ITELC 底稿的测试汇总表中，按 C 列编码（如 ELC01）匹配
      FileName.xlsx 中的文档，将文档名称填入 J 列、文档索引填入 K 列。
   5. 生成的底稿文件自动移入对应文件夹。
   6. 处理 IT审计发现汇总表 和 IT复杂性情况表。

依赖
   pip install xlwings   # 调用 Windows Excel 处理 .xlsx
   (其他均为 Python 内置模块：argparse, os, re, shutil, calendar, zipfile, xml)

用法示例
   python mwp.py huada 蒋jj limoumou 2026.7.27 2026 ELC SD OM CM

"""

from __future__ import annotations

import argparse
import calendar
import os
import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Sequence

import xlwings as xw


# ====================================================================
#  常量
# ====================================================================

TEMPLATE_DIR = "ITaudit_v1"
FILENAME_XLSX = "FileName.xlsx"
VALID_CHOICES = ("ELC", "SD", "CM", "OM", "SA")


# ====================================================================
#  命令行参数解析
# ====================================================================

def _parse_date(value: str) -> str:
    """验证并归一化日期为 YYYY.MM.DD 格式。"""
    if not re.fullmatch(r"\d{4}\.\d{1,2}\.\d{1,2}", value):
        raise argparse.ArgumentTypeError(
            f"日期格式错误: {value!r}，应为 YYYY.M.D，如 2026.7.27"
        )
    y, m, d = value.split(".")
    return f"{y}.{int(m):02d}.{int(d):02d}"


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="mwp.py",
        description="IT 审计工作底稿批量生成工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例: python mwp.py huada 蒋jj limoumou 2026.7.27 2026 ELC SD OM CM",
    )
    parser.add_argument("client", help="客户简称（替换文件名中的 XXXX）")
    parser.add_argument("name", help="编制人员姓名")
    parser.add_argument("reviewer", help="复核人员姓名")
    parser.add_argument("prepare", type=_parse_date, help="编制日期，格式 YYYY.M.D，如 2026.7.27")
    parser.add_argument("audity", help="审计年度，如 2026")
    parser.add_argument(
        "choices",
        nargs="+",
        choices=VALID_CHOICES,
        metavar="TYPE",
        help="底稿类型，可指定一个或多个: ELC SD CM OM SA",
    )
    return parser


# ====================================================================
#  工具函数
# ====================================================================

def _add_one_month(date_str: str) -> str:
    """在 YYYY.MM.DD 基础上加一个月，自动处理跨年和月末溢出。"""
    y, m, d = map(int, date_str.split("."))
    if m == 12:
        y, m = y + 1, 1
    else:
        m += 1
    d = min(d, calendar.monthrange(y, m)[1])
    return f"{y}.{m:02d}.{d:02d}"


def _read_folders(path: Path) -> list[str]:
    """用内置 zipfile + xml.etree 读取 FileName.xlsx 的 folder 列。

    不使用 pandas/openpyxl 等第三方库，纯 Python 内置模块完成。
    返回所有文档名称的列表（如 ['ELC01-01 战略规划与执行情况', ...]）。
    """
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

    with zipfile.ZipFile(path, "r") as z:
        # 读取共享字符串表（sharedStrings.xml）
        strings: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.parse(z.open("xl/sharedStrings.xml")).getroot()
            for si in root.findall(".//s:si", ns):
                t = si.find(".//s:t", ns)
                strings.append(t.text if t is not None and t.text else "")

        # 读取第一个 sheet（sheet1.xml）
        rows = ET.parse(z.open("xl/worksheets/sheet1.xml")).getroot().findall(".//s:row", ns)

        # 找到列头为 "folder" 的列号
        folder_col: str | None = None
        for row in rows:
            for cell in row.findall(".//s:c", ns):
                r = cell.get("r", "")
                v = cell.find(".//s:v", ns)
                if v is None or not v.text:
                    continue
                text = strings[int(v.text)] if cell.get("t") == "s" else v.text
                if text == "folder":
                    folder_col = "".join(ch for ch in r if ch.isalpha())
                    break
            if folder_col:
                break

        if folder_col is None:
            return []

        # 读取该列的所有数据行（跳过列头 "folder"）
        folders: list[str] = []
        for row in rows:
            for cell in row.findall(".//s:c", ns):
                r = cell.get("r", "")
                if "".join(ch for ch in r if ch.isalpha()) != folder_col:
                    continue
                v = cell.find(".//s:v", ns)
                if v is None or not v.text:
                    continue
                text = strings[int(v.text)] if cell.get("t") == "s" else v.text
                if text != "folder":
                    folders.append(text)
        return folders


# ====================================================================
#  主类 — IT 审计底稿生成器
# ====================================================================

class myWorkingPaper:
    """IT 审计底稿生成器

    工作流：
      1. 初始化（接收命令行参数）→ 2. 读取支持文档清单 → 3. 创建目录
      → 4. 打开模板 → 5. 替换占位单元格 → 6. 写入支持文档 → 7. 保存并移动
    """

    def __init__(
        self,
        client: str,
        name: str,
        reviewer: str,
        prepare: str,
        audity: str,
        choices: Sequence[str],
    ):
        """初始化参数

        Args:
            client:   客户简称（替换文件名中的 XXXX）
            name:     编制人员
            reviewer: 复核人员
            prepare:  编制日期，已归一化为 "YYYY.MM.DD"
            audity:   审计年度，如 "2026"
            choices:  底稿类型列表，如 ("ELC", "SD")
        """
        # 脚本所在目录作为根目录，不修改当前工作目录
        self.wpath = Path(__file__).resolve().parent

        # === Excel 实例设置 ===
        self.app = xw.App(visible=False, add_book=False)  # 后台运行，不显示 Excel 窗口
        self.app.display_alerts = False     # 关闭确认弹窗（覆盖保存等）
        self.app.screen_updating = False    # 关闭屏幕刷新（提升速度）
        self.app.enable_events = False      # 关闭事件响应（提升速度）

        # === 参数存储 ===
        self.client = client                           # 客户简称
        self.name = name                               # 编制人
        self.reviewer = reviewer                       # 复核人
        self.prepare = prepare                         # 编制日期，已归一化
        self.review_date = _add_one_month(self.prepare)  # 复核日期 = 编制日期 + 1 月
        self.audity = audity                           # 审计年度
        self.choices = list(choices)                   # 底稿类型列表
        self.support_cont: dict[str, list[str]] = {}   # 支持文档字典 {前缀: [文档列表]}

    def __enter__(self):
        return self

    def __exit__(self, *args):
        """退出时恢复 Excel 设置并释放进程"""
        self.app.enable_events = True
        self.app.quit()

    # ------------------------------------------------------------------
    #  步骤 1：读取 FileName.xlsx → 支持文档字典 support_cont
    # ------------------------------------------------------------------
    def read_support_cont(self, printf: bool = True) -> dict[str, list[str]]:
        """解析 FileName.xlsx，建立 {前缀: [文档列表]} 字典

        前缀由 re.findall(r'\\d+|\\D+', item)[0] 提取，
        例如 'ELC01-01 战略规划与执行情况' → 前缀 'ELC'。

        Returns:
            dict: 如 {'ELC': ['ELC01-01 战略规划与执行情况', ...], 'SD': [...], ...}
        """
        for item in _read_folders(self.wpath / FILENAME_XLSX):
            key = re.findall(r"\d+|\D+", item)[0]  # 提取前缀（ELC / SD / CM / OM / SA）
            self.support_cont.setdefault(key, []).append(item)

        if printf:
            for key, docs in sorted(self.support_cont.items()):
                print(f"{key}: {len(docs)} 个支持文档")
        return self.support_cont

    # ------------------------------------------------------------------
    #  步骤 2：创建输出目录结构
    # ------------------------------------------------------------------
    def paper_contents(self):
        """在 {客户简称}_{审计年度}/ 下创建底稿目录树

        结构示例：
          huada_2026/
          ├─ ITELC_企业层面控制_huada_2026/
          │  └─ ITELC_支持性文档_huada_2026/
          │     ├─ ELC01-01 战略规划与执行情况/
          │     └─ ...
          └─ ITGC_一般控制测试_huada_2026/
             └─ ITGC_支持性文档_huada_2026/
                ├─ SD01-01 可行性分析与立项审批/
                └─ ...
        """
        base = self.wpath / f"{self.client}_{self.audity}"
        base.mkdir(parents=True, exist_ok=True)

        for cont in self.choices:
            if cont == "ELC":
                p = base / f"ITELC_企业层面控制_{self.client}_{self.audity}" \
                          / f"ITELC_支持性文档_{self.client}_{self.audity}"
            else:
                p = base / f"ITGC_一般控制测试_{self.client}_{self.audity}" \
                          / f"ITGC_支持性文档_{self.client}_{self.audity}"
            p.mkdir(parents=True, exist_ok=True)
            for item in self.support_cont.get(cont, []):
                (p / item).mkdir(parents=True, exist_ok=True)
        print("Supporting documentation has been generated !!!")

    # ------------------------------------------------------------------
    #  工具：列号转换
    # ------------------------------------------------------------------
    @staticmethod
    def _col_index(col: str) -> int:
        """字母列号转数字：A→1, Z→26, AA→27"""
        n = 0
        for ch in col:
            n = n * 26 + ord(ch) - 64
        return n

    def _col_letter(self, idx: int) -> str:
        """数字转字母列号：1→A, 26→Z, 27→AA"""
        s = ""
        while idx > 0:
            idx, r = divmod(idx - 1, 26)
            s = chr(65 + r) + s
        return s

    # ------------------------------------------------------------------
    #  工具：用 Excel 原生 Find 定位单元格（比逐格 for 循环快几十倍）
    # ------------------------------------------------------------------
    def find_address(self, sht, keyword: str) -> str | None:
        """在 sheet 中查找第一个以 keyword 开头的单元格

        通过 Excel COM 的 Range.Find 方法实现，LookIn=-4163(xlValues)
        表示搜索显示值，LookAt=2(xlPart) 表示部分匹配。

        返回:
            str: 如 "$M$3"，若未找到则返回 None
        """
        found = sht.api.Cells.Find(What=keyword, LookIn=-4163, LookAt=2)
        if found is None:
            return None
        try:
            addr = found.Address  # 获取地址字符串（如 "$M$3"）
        except AttributeError:
            return None
        if not str(addr).startswith("$"):
            return None
        # 二次验证：确保该单元格的值确实以 keyword 开头
        cell = sht.range(addr)
        if cell.value is not None and str(cell.value).startswith(keyword):
            return cell.address
        return None

    # ------------------------------------------------------------------
    #  步骤 3：填写编制/日期、复核/日期、审计期间（并设字体 Arial）
    # ------------------------------------------------------------------
    def fill_cell(self, sht):
        """遍历当前 sheet，替换审计期间、编制/日期、复核/日期

        替换规则：
          - 以「审计期间」开头 → 「审计期间：2026.01.01-2026.12.31」
          - 以「编制」开头     → 「编制/日期：编制人/2026.07.27」
          - 以「复核」开头     → 「复核/日期：复核人/2026.08.27」

        替换后字体统一设置为 Arial。
        """
        for kw, val in [
            ("审计期间", f"审计期间：{self.audity}.01.01-{self.audity}.12.31"),
            ("编制", f"编制/日期：{self.name}/{self.prepare}"),
            ("复核", f"复核/日期：{self.reviewer}/{self.review_date}"),
        ]:
            addr = self.find_address(sht, kw)
            if addr:
                rng = sht.range(addr)
                rng.value = val
                rng.font.name = "Arial"

    # ------------------------------------------------------------------
    #  步骤 4：在内容 sheet 中填入支持文档清单
    # ------------------------------------------------------------------
    def fill_doc(self, cont: str, wb):
        """在 GC 底稿（SD/CM/OM/SA）的内容 sheet 中填入支持文档

        逻辑：
          1. 取 sheet 名称「-」前的部分作为键（如 "SD06-程序实施" → "SD06"）
          2. 在 support_cont[cont] 中查找以该键开头的文档
          3. 在 sheet 中找到「支持文档」单元格
          4. 在其右侧列（原列 + 1）纵向填入文档名称

        例如「支持文档」在 H14，则 I14 开始填入 SD06-01、SD06-02……
        每个填入的单元格字体设置为 Arial。
        """
        docs_source = self.support_cont.get(cont, [])
        for sht in wb.sheets:
            # 取 sheet 名称中的关键前缀
            key = sht.name.split("-")[0]
            docs = [x for x in docs_source if x.startswith(key)]
            if not docs:
                continue

            ad = self.find_address(sht, "支持文档")
            if not ad:
                continue

            m = re.search(r"\$?([A-Z]+)\$?(\d+)", ad)
            if not m:
                continue

            col, row = m.group(1), m.group(2)
            new_col = self._col_letter(self._col_index(col) + 1)

            # 纵向写入文档名称
            start_cell = sht.range(f"{new_col}{row}")
            for i, val in enumerate(docs):
                cell = start_cell.offset(i, 0)
                cell.value = val
                cell.font.name = "Arial"

    # ------------------------------------------------------------------
    #  步骤 5（ELC 专属）：按 C 列编码匹配文档，填入 J 列（名称）和 K 列（索引）
    # ------------------------------------------------------------------
    def _fill_elc_docs(self, ws):
        """在 ELC 测试汇总表中，根据 C 列编码填入对应的文档索引信息

        处理逻辑：
          1. 从第 13 行开始遍历到末尾
          2. 读取 C 列值（如 "ELC01"）
          3. 在 support_cont["ELC"] 中查找以该值开头的文档
          4. 对于每条文档，用正则 (ELC\\d+-\\d+)\\s+(.+) 拆分为索引和名称
          5. 所有名称换行拼接后写入 J 列，所有索引换行拼接后写入 K 列
          6. 字体统一设置为 Arial

        效果示例：
          C13 = "ELC01"
          → J13 = "战略规划与执行情况\n年度计划"
          → K13 = "ELC01-01\nELC01-02"
        """
        last = ws.used_range.last_cell.row
        elc_docs = self.support_cont.get("ELC", [])
        for r in range(13, last + 1):
            c_val = ws.range(f"C{r}").value
            if c_val is None or not str(c_val).startswith("ELC"):
                continue

            code = str(c_val).strip()
            docs = [d for d in elc_docs if d.startswith(code)]
            if not docs:
                continue

            names, indices = [], []
            for d in docs:
                m = re.match(r"(ELC\d+-\d+)\s+(.+)", d)
                if m:
                    indices.append(m.group(1))   # 如 "ELC01-01"
                    names.append(m.group(2))     # 如 "战略规划与执行情况"

            if names:
                cell_j = ws.range(f"J{r}")
                cell_j.value = "\n".join(names)
                cell_j.font.name = "Arial"
            if indices:
                cell_k = ws.range(f"K{r}")
                cell_k.value = "\n".join(indices)
                cell_k.font.name = "Arial"

    # ------------------------------------------------------------------
    #  工具：文件名 XXXX → 客户简称，期间 → 审计年度
    # ------------------------------------------------------------------
    def _new_name(self, filename: str) -> str:
        """替换模板文件名中的占位符

        例：
          "ITGC_SD_系统开发_XXXX_期间.xlsx"
          → "ITGC_SD_系统开发_huada_2026.xlsx"
        """
        return filename.replace("XXXX", self.client).replace("期间", self.audity)

    # ------------------------------------------------------------------
    #  主流程：打开模板 → 替换内容 → 保存 → 移动
    # ------------------------------------------------------------------
    def paper_prep(self, ptemplate: str):
        """执行所有底稿的模板处理

        Args:
            ptemplate: 模板文件所在目录名（如 "ITaudit_v1"）
        """
        template_path = self.wpath / ptemplate

        # 列出模板目录下所有 .xlsx 文件（排除 Excel 临时文件 ~$）
        files = [
            f for f in template_path.iterdir()
            if f.is_file() and f.suffix == ".xlsx" and not f.name.startswith("~$")
        ]
        file_names = [f.name for f in files]

        # 按类型分组：ITGC 文件以 "_" 分割后取第 2 段作为键（如 "SD"）
        gc = {f.name.split("_")[1]: f for f in files if f.name.startswith("ITGC")}
        elc_files = [f for f in files if f.name.startswith("ITELC")]
        elc = elc_files[0] if elc_files else None

        generated: list[str] = []

        # === 处理选中的底稿类型 ===
        for cont in self.choices:
            if cont == "ELC":
                if elc is None:
                    print("Warning: 未找到 ELC 对应的模板文件，跳过")
                    continue
                f = elc
            else:
                f = gc.get(cont)
                if f is None:
                    print(f"Warning: 未找到 {cont} 对应的模板文件，跳过")
                    continue

            wb = self.app.books.open(str(f))

            # 遍历所有 sheet，更新审计期间 / 编制/日期 / 复核/日期
            for sht in wb.sheets:
                self.fill_cell(sht)

            # 按类型执行不同的文档写入逻辑
            if cont == "ELC":
                self._fill_elc_docs(wb.sheets["测试汇总表"])  # ELC：按编码匹配文档索引
            else:
                self.fill_doc(cont, wb)                      # GC：按 sheet 前缀填入支持文档

            new = self._new_name(f.name)
            new_path = self.wpath / new
            wb.save(str(new_path))
            wb.close()
            generated.append(new)
            print(f"{new} generated !!!")

        # === 处理 IT审计发现汇总表 和 IT复杂性情况表 ===
        for f in files:
            if "IT复杂性" in f.name or "发现汇总" in f.name:
                wb = self.app.books.open(str(f))
                self.fill_cell(wb.sheets[0])
                new = self._new_name(f.name)
                new_path = self.wpath / new
                wb.save(str(new_path))
                wb.close()
                shutil.move(str(new_path), str(self.wpath / f"{self.client}_{self.audity}"))

        # === 移动生成的底稿到对应的子目录 ===
        base = self.wpath / f"{self.client}_{self.audity}"

        # 移动 ELC 底稿
        if elc is not None:
            elc_dir = base / f"ITELC_企业层面控制_{self.client}_{self.audity}"
            if elc_dir.exists():
                src = self.wpath / self._new_name(elc.name)
                if src.exists():
                    shutil.move(str(src), str(elc_dir / src.name))

        # 移动所有 ITGC 底稿
        gc_dir = base / f"ITGC_一般控制测试_{self.client}_{self.audity}"
        if gc_dir.exists():
            for f in [x for x in self.wpath.iterdir() if x.is_file() and x.name.startswith("ITGC")]:
                shutil.move(str(f), str(gc_dir / f.name))


# ====================================================================
#  程序入口
# ====================================================================

def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    with myWorkingPaper(
        client=args.client,
        name=args.name,
        reviewer=args.reviewer,
        prepare=args.prepare,
        audity=args.audity,
        choices=args.choices,
    ) as mwp:
        mwp.read_support_cont(False)    # 1. 读取 FileName.xlsx → 支持文档字典
        mwp.paper_contents()            # 2. 创建输出目录结构
        mwp.paper_prep(TEMPLATE_DIR)    # 3. 打开模板 → 替换内容 → 保存 → 移动

    print("The IT audit working papers have been generated !!!")
    return 0


if __name__ == "__main__":
    sys.exit(main())
