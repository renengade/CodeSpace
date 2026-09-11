#!/usr/bin/env python3
"""纯标准库 Excel(.xlsx) 读写：结构分析 + 可复现随机抽样 + 选列 + 结果导出"""

import random
import re
import zipfile
from datetime import date, datetime, timedelta
from xml.etree import ElementTree as ET

MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'m': MAIN_NS, 'r': REL_NS}

BUILTIN_DATE_FMT_IDS = {str(i) for i in range(14, 23)} | {str(i) for i in range(45, 48)} \
    | {str(i) for i in range(27, 37)} | {str(i) for i in range(50, 59)}

ET.register_namespace('', MAIN_NS)
ET.register_namespace('r', REL_NS)

_COL_RE = re.compile(r'([A-Z]+)(\d+)')


def _col_index(ref):
    """单元格引用 'AB12' -> 列索引 0-based"""
    m = _COL_RE.match(ref)
    if m is None:
        raise ValueError(f'非法单元格引用: {ref}')
    col = 0
    for ch in m.group(1):
        col = col * 26 + (ord(ch) - 64)
    return col - 1


def _serial_to_datetime(serial):
    """Excel 序列日期 -> date/datetime（处理 1900 闰年 bug）"""
    epoch = datetime(1899, 12, 30)
    if 0 < serial < 60:
        serial += 1
    days, frac = int(serial), serial - int(serial)
    dt = epoch + timedelta(days=days)
    if frac > 0:
        ms = round(frac * 86400000)
        return dt + timedelta(milliseconds=min(ms, 86399999))
    return dt.date()


class XlsxReader:
    """只读 .xlsx：解包、共享字符串、日期格式识别、按 sheet 取行"""

    def __init__(self, path):
        self.path = path
        self.zf = zipfile.ZipFile(path)
        self._shared_strings = self._load_shared_strings()
        self._date_styles = self._load_date_styles()
        self._sheet_paths = self._load_sheet_map()
        self._cache = {}

    def _read_xml(self, member):
        try:
            return ET.fromstring(self.zf.read(member))
        except KeyError:
            return None

    def _load_shared_strings(self):
        root = self._read_xml('xl/sharedStrings.xml')
        if root is None:
            return []
        return [''.join(si.itertext()) for si in root.findall('m:si', NS)]

    def _load_date_styles(self):
        root = self._read_xml('xl/styles.xml')
        if root is None:
            return set()
        num_fmts = {}
        for nf in root.findall('.//m:numFmts/m:numFmt', NS):
            num_fmts[nf.get('numFmtId')] = nf.get('formatCode', '')
        date_styles = set()
        for i, xf in enumerate(root.findall('.//m:cellXfs/m:xf', NS)):
            fmt_id = xf.get('numFmtId')
            # 剥离引号内文本与 [颜色] 标记，避免 "yds"、[Red] 等误判为日期
            code = re.sub(r'"[^"]*"|\[[^\]]*\]', '',
                          num_fmts.get(fmt_id, ''))
            if fmt_id in BUILTIN_DATE_FMT_IDS or re.search(r'[ymdhs]', code, re.IGNORECASE):
                date_styles.add(str(i))
        return date_styles

    def _load_sheet_map(self):
        root = self._read_xml('xl/workbook.xml')
        if root is None:
            raise KeyError('缺少 xl/workbook.xml，不是有效的 xlsx')
        rels = self._read_xml('xl/_rels/workbook.xml.rels')
        rid_to_target = {}
        if rels is not None:
            rid_to_target = {
                rel.get('Id'): rel.get('Target') or ''
                for rel in rels
                if (rel.get('Type') or '').endswith('/worksheet')
            }
        sheet_map = {}
        for sh in root.findall('m:sheets/m:sheet', NS):
            rid = sh.get(f'{{{REL_NS}}}id')
            target = rid_to_target.get(rid, '').lstrip('/')
            if target and not target.startswith('xl/'):
                target = 'xl/' + target
            if target:
                sheet_map[sh.get('name')] = target
        return sheet_map

    def _parse_cell(self, cell, style_id):
        t = cell.get('t')
        v = cell.find('m:v', NS)
        if t == 'inlineStr':
            is_ = cell.find('m:is', NS)
            return ''.join(is_.itertext()) if is_ is not None else None
        if v is None:
            return None
        text = v.text or ''
        if t == 's':
            try:
                return self._shared_strings[int(text)]
            except (ValueError, IndexError):
                return text
        if t == 'b':
            return text == '1'
        if t == 'str':
            return text
        if t == 'd':
            return datetime.fromisoformat(text.replace('Z', ''))
        if t == 'e':
            return f'#ERR({text})'
        if style_id in self._date_styles:
            try:
                return _serial_to_datetime(float(text))
            except ValueError:
                return None
        try:
            return int(text) if '.' not in text and 'e' not in text.lower() else float(text)
        except ValueError:
            return text

    def sheet_names(self):
        return list(self._sheet_paths)

    def get_rows(self, sheet_name):
        if sheet_name in self._cache:
            return self._cache[sheet_name]
        rows = []
        # 流式解析：逐行处理完立即释放，避免整棵 XML 树驻留内存
        sheet_data = None
        with self.zf.open(self._sheet_paths[sheet_name]) as f:
            context = ET.iterparse(f, events=('start', 'end'))
            for event, elem in context:
                tag = elem.tag
                if event == 'start' and tag == f'{{{MAIN_NS}}}sheetData':
                    sheet_data = elem
                elif (event == 'end' and tag == f'{{{MAIN_NS}}}row'
                      and sheet_data is not None):
                    cells = {}
                    for c in elem:
                        ref = c.get('r')
                        if not ref:
                            continue
                        try:
                            cells[_col_index(ref)] = self._parse_cell(c, c.get('s'))
                        except ValueError:
                            continue
                    values = [cells.get(i) for i in range(max(cells) + 1)] if cells else []
                    if any(v is not None for v in values):
                        rows.append(values)
                    sheet_data.remove(elem)
                    elem.clear()
        self._cache[sheet_name] = rows
        return rows

    def sheet_structure(self, sheet_name, infer_limit=200):
        rows = self.get_rows(sheet_name)
        if not rows:
            return {'columns': [], 'types': {}, 'row_count': 0}
        header = [str(h) if h is not None else f'col_{i+1}'
                  for i, h in enumerate(rows[0])]
        col_types = [set() for _ in header]
        for r in rows[1:1 + infer_limit]:
            for i, v in enumerate(r):
                if i < len(col_types) and v is not None:
                    col_types[i].add(type(v).__name__)
        type_names = []
        for ts in col_types:
            if not ts:
                type_names.append('empty')
            elif ts <= {'int'}:
                type_names.append('int')
            elif ts <= {'int', 'float'}:
                type_names.append('float')
            elif ts <= {'bool'}:
                type_names.append('bool')
            elif ts <= {'date', 'datetime'}:
                type_names.append('date/datetime')
            elif ts <= {'str'}:
                type_names.append('str')
            else:
                type_names.append('mixed:' + '+'.join(sorted(ts)))
        return {'columns': header, 'types': dict(zip(header, type_names)),
                'row_count': len(rows) - 1}


def random_sample_from_reader(reader, sheet_name, n=5, seed=42, columns=None):
    rows = reader.get_rows(sheet_name)
    if len(rows) < 2:
        return []
    header = rows[0]
    data = rows[1:]
    rng = random.Random(seed)
    sample = rng.sample(data, max(0, min(n, len(data))))
    if columns:
        missing = [c for c in columns if c not in header]
        if missing:
            raise KeyError(f'列不存在: {missing}，可用列: {header}')
        idx = [header.index(c) for c in columns]
        return [{c: row[i] for c, i in zip(columns, idx)} for row in sample]
    return [dict(zip(header, row)) for row in sample]


# ---------------- xlsx 写入（导出抽样结果） ----------------

_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '</Types>'
)

_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>'
)

_WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>'
)

_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<numFmts count="2">'
    '<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>'
    '<numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/>'
    '</numFmts>'
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="2">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '</fills>'
    '<borders count="1"><border/></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '<cellXfs count="3">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '</cellXfs>'
    '</styleSheet>'
)


def _col_letters(idx):
    idx += 1
    s = ''
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _datetime_to_serial(dt):
    """Python datetime/date -> Excel 序列值（处理 1900 闰年 bug）"""
    epoch = datetime(1899, 12, 30)
    if isinstance(dt, datetime):
        if dt.tzinfo:
            dt = dt.replace(tzinfo=None)
    else:
        dt = datetime(dt.year, dt.month, dt.day)
    days = (dt - epoch).days
    frac = (dt - epoch - timedelta(days=days)).total_seconds() / 86400
    if days < 61:
        days -= 1
    return days + frac


def _build_worksheet_xml(header, rows):
    ws = ET.Element(f'{{{MAIN_NS}}}worksheet')
    data = ET.SubElement(ws, f'{{{MAIN_NS}}}sheetData')

    def add_row(r_idx, values):
        row = ET.SubElement(data, f'{{{MAIN_NS}}}row')
        row.set('r', str(r_idx))
        for c_idx, v in enumerate(values):
            if v is None:
                continue
            cell = ET.SubElement(row, f'{{{MAIN_NS}}}c')
            cell.set('r', _col_letters(c_idx) + str(r_idx))
            val = ET.SubElement(cell, f'{{{MAIN_NS}}}v')
            if isinstance(v, bool):
                cell.set('t', 'b')
                val.text = '1' if v else '0'
            elif isinstance(v, int):
                val.text = str(v)
            elif isinstance(v, float):
                val.text = repr(v)
            elif isinstance(v, datetime):
                cell.set('s', '2')
                val.text = repr(_datetime_to_serial(v))
            elif isinstance(v, date):
                cell.set('s', '1')
                val.text = repr(_datetime_to_serial(v))
            else:
                cell.set('t', 'inlineStr')
                is_ = ET.SubElement(cell, f'{{{MAIN_NS}}}is')
                t = ET.SubElement(is_, f'{{{MAIN_NS}}}t')
                t.text = str(v)

    add_row(1, header)
    for i, row in enumerate(rows, 2):
        add_row(i, [row.get(h) for h in header])
    return ET.tostring(ws, encoding='utf-8', xml_declaration=True)


def write_sample_xlsx(path, sheet_name, sample):
    """把抽样结果(list[dict])导出为 .xlsx，仅用标准库"""
    header = list(sample[0].keys()) if sample else []
    safe = re.sub(r'[\\/:*?\[\]]', '_', sheet_name)[:31] or 'sample'

    wb = ET.Element(f'{{{MAIN_NS}}}workbook')
    sheets = ET.SubElement(wb, f'{{{MAIN_NS}}}sheets')
    sh = ET.SubElement(sheets, f'{{{MAIN_NS}}}sheet')
    sh.set('name', safe)
    sh.set('sheetId', '1')
    sh.set(f'{{{REL_NS}}}id', 'rId1')
    workbook_xml = ET.tostring(wb, encoding='utf-8', xml_declaration=True)

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', _CONTENT_TYPES)
        zf.writestr('_rels/.rels', _ROOT_RELS)
        zf.writestr('xl/workbook.xml', workbook_xml)
        zf.writestr('xl/_rels/workbook.xml.rels', _WORKBOOK_RELS)
        zf.writestr('xl/styles.xml', _STYLES_XML)
        zf.writestr('xl/worksheets/sheet1.xml',
                    _build_worksheet_xml(header, sample))


if __name__ == '__main__':
    import argparse
    import os
    import sys

    ap = argparse.ArgumentParser(
        description='纯标准库 xlsx 读取：结构总览 + 随机抽样 + 导出 xlsx',
        epilog=(
            '用法示例:\n'
            '1. 非交互模式 python sampling.py 数据.xlsx -s Sheet1 -n 10 -S 42 -c A列,B列 -o -p 结果.xlsx\n'
            '\n'
            '2. 交互模式 不传 -s/-n/-S/-c 时进入交互模式，按提示依次输入 sheet/样本量/种子/抽样列，最后询问是否导出 xlsx。'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument('file', help='.xlsx 文件路径')
    ap.add_argument('-s', '--sheet', help='sheet 名（跳过 sheet 选择）')
    ap.add_argument('-n', '--num', type=int, help='抽样条数（跳过样本量输入）')
    ap.add_argument('-S', '--seed', type=int, help='随机种子（跳过种子输入）')
    ap.add_argument('-c', '--cols', help='逗号分隔的列名，只返回这些列')
    ap.add_argument('-o', '--output', action='store_true',
                    help='将抽样结果写入 xlsx 文件（默认不输出；指定 --path 时自动开启）')
    ap.add_argument('-p', '--path', help='输出文件路径（xlsx 格式；默认: 本脚本所在目录）')
    args = ap.parse_args()

    def ask(prompt, default):
        if not sys.stdin.isatty():
            return default
        try:
            return input(prompt).strip() or default
        except EOFError:
            return default

    try:
        reader = XlsxReader(args.file)
    except FileNotFoundError:
        raise SystemExit(f'文件不存在: {args.file}')
    except (KeyError, zipfile.BadZipFile, OSError):
        raise SystemExit(f'不是有效的 .xlsx 文件（不支持 .xls 或文件已损坏）: {args.file}')
    names = reader.sheet_names()
    if not names:
        raise SystemExit('文件中没有任何 sheet')

    infos = {name: reader.sheet_structure(name) for name in names}
    print('=' * 62)
    for i, name in enumerate(names, 1):
        info = infos[name]
        print(f'[{i}] sheet: {name}')
        print(f'    数据总量: {info["row_count"]} 行')
        print(f'    列名    : {", ".join(info["columns"]) or "(无)"}')
    print('=' * 62)

    if args.sheet:
        sheet = args.sheet
        if sheet not in names:
            raise SystemExit(f'sheet 不存在: {sheet}，可用: {names}')
    elif len(names) > 1:
        while True:
            choice = ask('请选择要抽样的 sheet [1]: ', '1')
            if choice.isdigit() and 1 <= int(choice) <= len(names):
                sheet = names[int(choice) - 1]
                break
            print(f'无效选择，请输入 1~{len(names)}')
    else:
        sheet = names[0]

    info = infos[sheet]
    total = info['row_count']

    n = args.num
    if n is None:
        while True:
            raw = ask(f'请输入样本量 [5]（共 {total} 行）: ', '5')
            if raw.isdigit() and int(raw) > 0:
                n = int(raw)
                break
            print('请输入正整数')
        if n > total:
            n = total
            print(f'样本量超出数据总量，已调整为 {n}')

    seed = args.seed
    if seed is None:
        while True:
            raw = ask('请输入随机种子 [42]: ', '42')
            if raw.isdigit():
                seed = int(raw)
                break
            print('请输入整数')

    cols = args.cols
    if cols is None:
        raw = ask('请输入需要的列（逗号分隔，留空=全部列）: ', '')
        cols = [c.strip() for c in raw.split(',')] if raw else None
    else:
        cols = [c.strip() for c in cols.split(',')]

    for name in names:
        if name != sheet:
            reader._cache.pop(name, None)
    sample = random_sample_from_reader(reader, sheet, n=n, seed=seed, columns=cols)
    print(f'\n抽样结果（{sheet}，n={n}，seed={seed}）:')
    for row in sample:
        print(row)

    write_out = bool(args.output or args.path)
    if not write_out:
        raw = ask('是否将抽样结果导出？[y/n]: ', 'n').lower()
        write_out = raw in ('y', 'yes', '是', '1')
    if write_out:
        safe_sheet = re.sub(r'[\\/:*?"<>|]', '_', sheet)
        base = os.path.splitext(os.path.basename(args.file))[0]
        default_name = f'{base}_{safe_sheet}_sample.xlsx'
        if args.path and os.path.isdir(args.path):
            out_path = os.path.join(args.path, default_name)
        elif args.path:
            out_path = args.path
        else:
            out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    default_name)
        if not out_path.lower().endswith('.xlsx'):
            out_path += '.xlsx'
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        write_sample_xlsx(out_path, sheet, sample)
        print(f'\n抽样结果已写入: {out_path}')
