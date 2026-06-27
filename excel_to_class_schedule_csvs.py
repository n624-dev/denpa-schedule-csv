import csv
import re
import shutil
import sys
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

DEFAULT_SHEET_NAME = '時間割変更'
REQUIRED_HEADERS = ['学 年', '学科・クラス', '月日']
YEAR_COL = '学 年'
CLASS_COL = '学科・クラス'
EXCLUDED_FROM_ALL = {'AI'}

NS_MAIN = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
NS_RELS = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}


def get_input_path() -> Path:
    """
    入力ファイルを決める。
    1. コマンドライン引数があれば、それを優先する
       例: python excel_to_class_schedule_csvs.py henkou.xlsx
       Windowsではxlsx/csvをこの.pyにドラッグ&ドロップしても argv[1] に入る
    2. 引数がなければ、入力してもらう
    """
    if len(sys.argv) >= 2:
        raw = sys.argv[1]
    else:
        raw = input('変換したいExcelまたはCSVファイルのパスを入力してください（ドラッグ&ドロップ可）: ')

    raw = raw.strip().strip('"').strip("'")
    if not raw:
        raise ValueError('ファイルのパスが入力されていません。')

    path = Path(raw).expanduser()
    if not path.exists():
        raise FileNotFoundError(f'ファイルが見つかりません: {path}')
    if not path.is_file():
        raise ValueError(f'ファイルではありません: {path}')
    if path.suffix.lower() not in ('.xlsx', '.csv'):
        raise ValueError(f'.xlsx または .csv ファイルを指定してください: {path.name}')

    return path.resolve()


# =============================
# Excel -> 通常CSV
# =============================

def column_index_from_cell_ref(cell_ref: str) -> int:
    """A1 -> 0, B1 -> 1, AA1 -> 26"""
    letters = ''.join(ch for ch in cell_ref if ch.isalpha()).upper()
    if not letters:
        return -1

    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
    except KeyError:
        return []

    strings: list[str] = []
    for si in root.findall('x:si', NS_MAIN):
        parts = [t.text or '' for t in si.findall('.//x:t', NS_MAIN)]
        strings.append(''.join(parts))
    return strings


def load_sheet_path(zf: zipfile.ZipFile, preferred_sheet_name: str) -> tuple[str, str]:
    workbook_root = ET.fromstring(zf.read('xl/workbook.xml'))
    rels_root = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))

    rels: dict[str, str] = {}
    for rel in rels_root.findall('rel:Relationship', NS_RELS):
        rel_id = rel.attrib.get('Id')
        target = rel.attrib.get('Target')
        if rel_id and target:
            if not target.startswith('/'):
                target = 'xl/' + target
            else:
                target = target.lstrip('/')
            rels[rel_id] = target

    sheets: list[tuple[str, str]] = []
    for sheet in workbook_root.findall('x:sheets/x:sheet', NS_MAIN):
        name = sheet.attrib.get('name', '')
        rel_id = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        if rel_id in rels:
            sheets.append((name, rels[rel_id]))

    if not sheets:
        raise ValueError('Excel内のシートを読み取れませんでした。')

    for name, path in sheets:
        if name == preferred_sheet_name:
            return name, path

    return sheets[0]


def read_cell_value(cell, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get('t')

    if cell_type == 'inlineStr':
        parts = [t.text or '' for t in cell.findall('.//x:t', NS_MAIN)]
        return ''.join(parts)

    v = cell.find('x:v', NS_MAIN)
    if v is None or v.text is None:
        return ''

    text = v.text
    if cell_type == 's':
        try:
            return shared_strings[int(text)]
        except (ValueError, IndexError):
            return text
    if cell_type == 'b':
        return 'TRUE' if text == '1' else 'FALSE'

    return text


def normalize_number_text(value: str) -> str:
    """1.0 のような整数相当の数値を 1 にする。"""
    v = value.strip()
    if re.fullmatch(r'-?\d+\.0+', v):
        return v.split('.', 1)[0]
    return v


def excel_serial_to_date_text(value: str) -> str:
    """Excelのシリアル値 46119 を 2026/4/7 のようにする。"""
    v = value.strip()
    if not re.fullmatch(r'\d+(\.\d+)?', v):
        return v

    serial = float(v)
    if serial <= 0:
        return v

    dt = datetime(1899, 12, 30) + timedelta(days=serial)
    return f'{dt.year}/{dt.month}/{dt.day}'


def read_sheet_rows(xlsx_path: Path, sheet_name: str = DEFAULT_SHEET_NAME) -> tuple[str, list[list[str]]]:
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings = load_shared_strings(zf)
        actual_sheet_name, sheet_path = load_sheet_path(zf, sheet_name)
        root = ET.fromstring(zf.read(sheet_path))

    rows: list[list[str]] = []
    for row in root.findall('.//x:sheetData/x:row', NS_MAIN):
        values_by_col: dict[int, str] = {}
        max_col = -1
        for cell in row.findall('x:c', NS_MAIN):
            cell_ref = cell.attrib.get('r', '')
            col = column_index_from_cell_ref(cell_ref)
            if col < 0:
                continue
            max_col = max(max_col, col)
            values_by_col[col] = read_cell_value(cell, shared_strings)

        if max_col >= 0:
            rows.append([values_by_col.get(i, '') for i in range(max_col + 1)])
        else:
            rows.append([])

    return actual_sheet_name, rows


def find_header_row_index(rows: list[list[str]]) -> int:
    for index, row in enumerate(rows):
        stripped = [str(value).strip() for value in row]
        if all(header in stripped for header in REQUIRED_HEADERS):
            return index
    raise ValueError(f'見出し行が見つかりません。必要な見出し: {", ".join(REQUIRED_HEADERS)}')


def convert_xlsx_to_csv(xlsx_path: Path) -> Path:
    sheet_name, rows = read_sheet_rows(xlsx_path)
    header_index = find_header_row_index(rows)

    headers = [str(value).strip() if str(value).strip() != '時限' else str(value) for value in rows[header_index]]
    col_count = len(headers)
    date_col = next((i for i, header in enumerate(headers) if header.strip() == '月日'), None)

    output_rows: list[list[str]] = [headers]

    for row in rows[header_index + 1:]:
        fixed = list(row[:col_count]) + [''] * max(0, col_count - len(row))
        fixed = [normalize_number_text(str(value)) for value in fixed]

        # 完全な空行は無視する。下に追加された新規データはそのまま拾う。
        if not any(value.strip() for value in fixed):
            continue

        if date_col is not None and date_col < len(fixed):
            fixed[date_col] = excel_serial_to_date_text(fixed[date_col])

        output_rows.append(fixed)

    out_csv = xlsx_path.with_suffix('.csv')
    with out_csv.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(output_rows)

    print(f'読み取りシート: {sheet_name}')
    print(f'通常CSV: {out_csv}')
    print(f'通常CSV行数: {len(output_rows) - 1}件')
    return out_csv


# =============================
# 通常CSV -> クラス別CSV
# =============================

def expand_years(value: str) -> list[str]:
    """1～3 のような学年範囲を ['1', '2', '3'] にする。"""
    v = (value or '').strip()
    m = re.fullmatch(r'(\d+)\s*[～〜~-]\s*(\d+)', v)
    if m:
        start, end = map(int, m.groups())
        step = 1 if start <= end else -1
        return [str(i) for i in range(start, end + step, step)]
    return [v] if v else []


def split_class_cell(value: str) -> list[str]:
    """1,2 のような複数クラス指定を ['1', '2'] にする。全はここでは展開しない。"""
    v = (value or '').strip()
    if not v or v == '全':
        return []
    return [part.strip() for part in re.split(r'[,，]', v) if part.strip()]


def sort_key(class_key: str) -> tuple[int, str]:
    year, cls = class_key.split(maxsplit=1)
    return (int(year) if year.isdigit() else 999, cls)


def split_csv_by_class(src: Path) -> tuple[Path, Path, list[list[str]]]:
    out_dir = src.parent / f'{src.stem}_class_csvs'
    zip_path = src.parent / f'{src.stem}_class_csvs.zip'

    with src.open('r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames
        rows = list(reader)

    if headers is None:
        raise ValueError('CSVのヘッダーを読み取れませんでした。')
    if YEAR_COL not in headers or CLASS_COL not in headers:
        raise ValueError(f'CSVに必要な列がありません。必要な列: {YEAR_COL}, {CLASS_COL}')

    known_classes_by_year: dict[str, set[str]] = defaultdict(set)
    for row in rows:
        for year in expand_years(row.get(YEAR_COL, '')):
            for cls in split_class_cell(row.get(CLASS_COL, '')):
                known_classes_by_year[year].add(cls)

    by_class: dict[str, list[tuple[int, dict[str, str]]]] = defaultdict(list)

    for idx, row in enumerate(rows):
        years = expand_years(row.get(YEAR_COL, ''))
        class_cell = (row.get(CLASS_COL, '') or '').strip()

        for year in years:
            if class_cell == '全':
                target_classes = sorted(
                    cls for cls in known_classes_by_year.get(year, set())
                    if cls not in EXCLUDED_FROM_ALL
                )
            else:
                target_classes = split_class_cell(class_cell)

            for cls in target_classes:
                key = f'{year} {cls}'
                by_class[key].append((idx, row.copy()))

    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    summary: list[list[str]] = []
    for key in sorted(by_class.keys(), key=sort_key):
        safe_name = key.replace(' ', '_').replace('/', '_').replace('\\', '_')
        path = out_dir / f'{safe_name}.csv'
        entries = sorted(by_class[key], key=lambda item: item[0])

        with path.open('w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            for _, row in entries:
                writer.writerow(row)

        summary.append([key, str(len(entries)), path.name])

    with (out_dir / '_summary.csv').open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['クラス', '件数', 'ファイル名'])
        writer.writerows(summary)

    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for file in sorted(out_dir.iterdir()):
            zf.write(file, arcname=file.name)

    return out_dir, zip_path, summary


def main() -> None:
    src = get_input_path()

    if src.suffix.lower() == '.xlsx':
        csv_path = convert_xlsx_to_csv(src)
    else:
        csv_path = src
        print(f'通常CSV変換はスキップしました: {csv_path}')

    out_dir, zip_path, summary = split_csv_by_class(csv_path)

    print('')
    print(f'{len(summary)}個のクラス別CSVを作成しました。')
    print(f'出力フォルダ: {out_dir}')
    print(f'ZIP: {zip_path}')
    print('')
    print('クラス, 件数, ファイル名')
    for class_name, count, filename in summary:
        print(f'{class_name}, {count}, {filename}')


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'エラー: {e}', file=sys.stderr)
        sys.exit(1)
