#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
读取指定目录下的 csv/xlsx 等文件，合并指定列后写出为 SQLite / xlsx / csv / tsv / json。

不传 -d 时默认扫描脚本所在目录；
不传 -c 时默认使用文件的所有列（多文件取并集）；
不传 -p 时不按文件名过滤（取目录下该类型的所有文件）；
不传 -o 时不写出，只做合并与统计。
传入 -o 时可多次指定，一次导出多种格式。

输出格式默认按 -o 的扩展名推断：
  .db / .sqlite / .sqlite3 -> sqlite
  .xlsx / .xls             -> xlsx
  .csv                     -> csv
  .tsv / .txt              -> tsv
  .json                    -> json
也可用 --format 显式覆盖（仅在只有一个 -o 时有效）。

示例:
  # 合并但不写出
  python xlsxmerge.py -d ./data -t csv

  # 只输出 SQLite
  python xlsxmerge.py -t csv -p "操作日志*" -o out.db -T logs

  # 只输出 xlsx
  python xlsxmerge.py -d ./data -t csv -p "日志*" -o merged.xlsx

  # 同时输出 SQLite + xlsx + csv
  python xlsxmerge.py -d ./data -t csv -o logs.db -T logs -o logs.xlsx -o logs.csv
"""

import argparse
import fnmatch
import json
import sqlite3
import sys
from pathlib import Path
from typing import Literal, cast

import pandas as pd

IfExists = Literal["replace", "append", "fail"]
OutputFormat = Literal["sqlite", "xlsx", "csv", "tsv", "json"]

SUPPORTED_EXTS = ("csv", "tsv", "txt", "xlsx", "xls")
SEPARATORS = {"csv": ",", "tsv": "\t", "txt": "\t"}
DEFAULT_TABLE = "data"

# 输出扩展名 -> 格式
OUTPUT_EXTS: dict[str, OutputFormat] = {
    ".db": "sqlite",
    ".sqlite": "sqlite",
    ".sqlite3": "sqlite",
    ".xlsx": "xlsx",
    ".xls": "xlsx",
    ".csv": "csv",
    ".tsv": "tsv",
    ".txt": "tsv",
    ".json": "json",
}


def script_dir() -> Path:
    """脚本所在目录（默认扫描目录）"""
    return Path(__file__).resolve().parent


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="合并指定目录下的文件列，可同时输出为 SQLite / xlsx / csv / tsv / json",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    p.add_argument("-d", "--dir", type=Path, default=script_dir(), help="要扫描的目录，默认脚本所在目录")
    p.add_argument(
        "-t", "--type", required=True, choices=SUPPORTED_EXTS, help="输入文件类型: " + ", ".join(SUPPORTED_EXTS)
    )
    p.add_argument(
        "-p",
        "--pattern",
        default=None,
        help='模糊匹配文件名(支持通配符 * ?)，例如 "日志*" 或 "*2024*"；' "不传则不按文件名过滤（取该类型所有文件）",
    )
    p.add_argument(
        "-c",
        "--columns",
        default=None,
        help="要合并的列名，逗号分隔，例如: 时间,用户,内容。" "不传则使用所有列(多文件取并集)",
    )
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        action="append",
        default=None,
        help="输出文件路径，可多次指定；格式按扩展名推断"
        "（.db/.sqlite/.sqlite3/.xlsx/.csv/.tsv/.json）。\n"
        "例如: -o logs.db -o logs.xlsx -o logs.csv\n"
        "不传则只输出合并统计",
    )
    p.add_argument(
        "-f",
        "--format",
        default=None,
        choices=["sqlite", "xlsx", "csv", "tsv", "json"],
        help="显式指定输出格式（仅在只有一个 -o 时有效）",
    )
    p.add_argument(
        "-T",
        "--table",
        default=None,
        help=f"SQLite 表名；未指定时用输出文件 stem（如 logs.db -> logs），" f"stem 为空则用 {DEFAULT_TABLE}",
    )
    p.add_argument(
        "--if-exists",
        default="replace",
        choices=["replace", "append", "fail"],
        help="SQLite 表已存在时的处理方式，默认 replace",
    )
    p.add_argument("--recursive", action="store_true", help="是否递归扫描子目录")
    p.add_argument("--add-source", action="store_true", help="是否额外增加一列 _source_file 记录来源文件名")
    p.add_argument("--encoding", default="utf-8", help="输入 csv 编码，默认 utf-8（可试 gbk）")
    p.add_argument(
        "--out-encoding",
        default="utf-8-sig",
        help="输出 csv/tsv/json 的编码，默认 utf-8-sig（Excel 友好）；" "对 xlsx/sqlite 无效",
    )
    p.add_argument("--dry-run", action="store_true", help="只打印将要处理的文件，不执行")
    return p.parse_args()


def find_files(root: Path, ext: str, pattern: str | None, recursive: bool) -> list[Path]:
    """返回匹配的文件列表；目录不存在时返回空列表。

    pattern 为 None 或空字符串时不做文件名过滤。
    """
    if not root.is_dir():
        return []
    it = root.rglob("*") if recursive else root.glob("*")
    pat = pattern or None
    return sorted(
        f
        for f in it
        if f.is_file() and f.suffix.lower().lstrip(".") == ext and (pat is None or fnmatch.fnmatch(f.name, pat))
    )


def read_file(path: Path, ext: str, encoding: str) -> pd.DataFrame:
    """按扩展名读取为 str 类型的 DataFrame，空值统一为 ""。"""
    if ext in SEPARATORS:
        return pd.read_csv(
            path,
            sep=SEPARATORS[ext],
            dtype=str,
            keep_default_na=False,
            encoding=encoding,
        )
    return pd.read_excel(path, dtype=str, keep_default_na=False)


def load_files(
    files: list[Path],
    ext: str,
    columns: list[str] | None,
    encoding: str,
    add_source: bool,
) -> pd.DataFrame:
    """读取并合并文件。columns 为 None 时使用所有列（多文件取并集）。"""
    frames: list[pd.DataFrame] = []

    for f in files:
        try:
            df = read_file(f, ext, encoding)
        except Exception as e:
            print(f"[跳过] 读取失败 {f}: {e}", file=sys.stderr)
            continue

        if columns is not None:
            missing = [c for c in columns if c not in df.columns]
            if missing:
                print(f"[跳过] {f} 缺少列 {missing}，实际列: {list(df.columns)}", file=sys.stderr)
                continue
            sub = df.loc[:, columns].copy()
        else:
            sub = df.copy()

        if add_source:
            if "_source_file" in sub.columns:
                print(f"[警告] {f} 已含 _source_file 列，将被覆盖", file=sys.stderr)
            sub["_source_file"] = f.name

        frames.append(sub)
        print(f"[读取] {f}  行数={len(sub)}  列数={sub.shape[1]}")

    if not frames:
        cols = (columns or []) + (["_source_file"] if add_source else [])
        return pd.DataFrame(columns=cols)

    # 多文件列不一致时按列名对齐，缺失处为 NaN
    return pd.concat(frames, ignore_index=True, sort=False)


# ---------- 各格式写出 ----------


def write_sqlite(df: pd.DataFrame, path: Path, table: str, if_exists: IfExists) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    try:
        df.to_sql(table, conn, if_exists=if_exists, index=False)
        row = conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()
        total = int(row[0]) if row else 0
    finally:
        conn.close()
    return f"表 `{table}`，当前总行数 {total}"


def write_xlsx(df: pd.DataFrame, path: Path) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_excel(path, index=False)
    except ImportError as e:
        raise RuntimeError(f"写出 xlsx 需要 openpyxl，请先安装: pip install openpyxl（{e}）") from e
    return f"共 {len(df)} 行"


def write_csv_like(df: pd.DataFrame, path: Path, sep: str, encoding: str) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, sep=sep, index=False, encoding=encoding)
    return f"共 {len(df)} 行"


def write_json(df: pd.DataFrame, path: Path, encoding: str) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    records = df.to_dict(orient="records")
    with path.open("w", encoding=encoding) as fp:
        json.dump(records, fp, ensure_ascii=False, indent=2)
    return f"共 {len(records)} 条记录"


def detect_format(output: Path, override: str | None) -> OutputFormat:
    """根据 --format 或扩展名决定输出格式。"""
    if override:
        return cast(OutputFormat, override)
    fmt = OUTPUT_EXTS.get(output.suffix.lower())
    if fmt is None:
        raise SystemExit(
            f"[错误] 无法根据扩展名 '{output.suffix}' 推断输出格式，"
            f"请用 --format 指定（可选: sqlite, xlsx, csv, tsv, json）"
        )
    return fmt


def resolve_table(output: Path, table_arg: str | None) -> str:
    """确定 SQLite 输出使用的表名。

    优先级：-T 指定 > 输出文件 stem > DEFAULT_TABLE
    """
    if table_arg:
        return table_arg
    return output.stem or DEFAULT_TABLE


def write_all(
    df: pd.DataFrame,
    jobs: list[tuple[Path, OutputFormat]],
    *,
    table_arg: str | None,
    if_exists: IfExists,
    encoding: str,
) -> int:
    """依次执行所有输出任务；返回失败数量。

    单个输出失败不影响后续输出。
    """
    failures = 0
    for out, fmt in jobs:
        try:
            if fmt == "sqlite":
                tname = resolve_table(out, table_arg)
                detail = write_sqlite(df, out, tname, if_exists)
            elif fmt == "xlsx":
                detail = write_xlsx(df, out)
            elif fmt == "csv":
                detail = write_csv_like(df, out, ",", encoding)
            elif fmt == "tsv":
                detail = write_csv_like(df, out, "\t", encoding)
            elif fmt == "json":
                detail = write_json(df, out, encoding)
            else:
                raise RuntimeError(f"未知格式: {fmt}")
        except Exception as e:
            print(f"[失败] 写出 {out}（{fmt}）: {e}", file=sys.stderr)
            failures += 1
            continue
        print(f"[完成] 格式={fmt} 已写出 {out}（{detail}）")
    return failures


def main() -> None:
    args = parse_args()

    # -c 未传 -> None（表示使用所有列）；传了则拆分、去空、去重
    raw = [c.strip() for c in (args.columns or "").split(",") if c.strip()]
    columns: list[str] | None = list(dict.fromkeys(raw)) or None

    if not args.dir.is_dir():
        sys.exit(f"[错误] 目录不存在: {args.dir}")

    outputs: list[Path] = args.output or []

    # --format 只能用于单输出，多输出时用扩展名区分
    if args.format and len(outputs) > 1:
        sys.exit("[错误] --format 只能用于单个 -o；多输出时请用扩展名区分")

    # 预解析所有输出的格式（尽早暴露无法推断的扩展名）
    jobs: list[tuple[Path, OutputFormat]] = []
    for out in outputs:
        jobs.append((out, detect_format(out, args.format)))

    # 表名合法性只在确实要写 sqlite 时校验
    if any(fmt == "sqlite" for _, fmt in jobs) and args.table and '"' in args.table:
        sys.exit("[错误] 表名不能包含双引号")

    files = find_files(args.dir, args.type, args.pattern, args.recursive)
    if not files:
        desc = f"匹配 '{args.pattern}' 的" if args.pattern else ""
        sys.exit(f"[错误] 在 {args.dir} 下没有找到{desc} .{args.type} 文件")

    print(f"[信息] 扫描目录: {args.dir}")
    if args.pattern:
        print(f"[信息] 文件名过滤: {args.pattern}")
    print(f"[信息] 找到 {len(files)} 个文件:")
    for f in files:
        print("   -", f)

    if args.dry_run:
        print("[dry-run] 不执行读取与写出")
        return

    df = load_files(files, args.type, columns, args.encoding, args.add_source)
    print(f"[信息] 合并后共 {len(df)} 行, 列: {list(df.columns)}")

    if df.empty:
        print("[警告] 合并结果为空，未写出")
        return

    if not jobs:
        print("[信息] 未指定 -o/--output，跳过写出")
        return

    failures = write_all(
        df,
        jobs,
        table_arg=args.table,
        if_exists=cast(IfExists, args.if_exists),
        encoding=args.out_encoding,
    )
    if failures:
        sys.exit(f"[警告] 共 {len(jobs)} 个输出，其中 {failures} 个失败")


if __name__ == "__main__":
    main()
