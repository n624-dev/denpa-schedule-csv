import argparse
import csv
import json
import re
import shutil
import sys
import unicodedata
import zipfile
from collections import defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

DEFAULT_SHEET_NAME = '時間割変更'
REQUIRED_HEADERS = ['学 年', '学科・クラス', '月日']
YEAR_COL = '学 年'
CLASS_COL = '学科・クラス'
DATE_COL = '月日'
EXCLUDED_FROM_ALL = {'AI'}

NORMALIZED_HEADERS = [
    'change_date',
    'class_name',
    'period',
    'before_subject',
    'after_subject',
    'teacher',
    'room',
    'note',
    'raw_text',
    'canonical_text',
]

NS_MAIN = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
NS_RELS = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}


class ScheduleCsvError(Exception):
    """時間割CSV変換の基底例外。"""


class UnsupportedInputFormatError(ScheduleCsvError, ValueError):
    """未対応の入力形式。"""


class SheetNotFoundError(ScheduleCsvError, ValueError):
    """指定シートが見つからない。"""


class RequiredHeaderNotFoundError(ScheduleCsvError, ValueError):
    """必須ヘッダーが見つからない。"""


class ConversionError(ScheduleCsvError):
    """変換処理全般のエラー。"""


@dataclass(frozen=True)
class ClassCsvFile:
    class_name: str
    path: Path
    rows: int


@dataclass(frozen=True)
class ConversionResult:
    input_path: Path
    output_dir: Path
    normalized_csv_path: Path
    classes_dir: Path
    summary_csv_path: Path
    manifest_path: Path
    zip_path: Path | None
    sheet_name: str | None
    class_files: list[ClassCsvFile]
    warnings: list[str]


def normalize_text(value: object) -> str:
    """全角半角・空白・改行を通知Bot向けに安定化する。"""
    text = unicodedata.normalize('NFKC', str(value or ''))
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r'[ \t\f\v]+', ' ', text)
    text = re.sub(r'\n+', '\n', text)
    return text.strip()


def normalize_token(value: object) -> str:
    """クラス名や列判定向けに空白を除去して大文字化する。"""
    return normalize_text(value).replace(' ', '').upper()


def normalize_number_text(value: str) -> str:
    """1.0 のような整数相当の数値を 1 にする。"""
    v = normalize_text(value)
    if re.fullmatch(r'-?\d+\.0+', v):
        return v.split('.', 1)[0]
    return v


def excel_serial_to_date_text(value: str) -> str:
    """Excelのシリアル値 46119 を 2026/4/7 のようにする。"""
    v = normalize_text(value)
    if not re.fullmatch(r'\d+(\.\d+)?', v):
        return v

    serial = float(v)
    if serial <= 0:
        return v

    dt = datetime(1899, 12, 30) + timedelta(days=serial)
    return f'{dt.year}/{dt.month}/{dt.day}'


def normalize_date_value(value: object, default_year: int | None = None) -> str:
    """月日を YYYY-MM-DD に正規化する。年なしなら default_year を使う。"""
    text = excel_serial_to_date_text(normalize_text(value))
    if not text:
        return ''

    text = text.replace('年', '/').replace('月', '/').replace('日', '')
    text = text.replace('.', '/').replace('-', '/')
    text = re.sub(r'\s+', '', text)

    match = re.fullmatch(r'(\d{4})/(\d{1,2})/(\d{1,2})', text)
    if match:
        year, month, day = map(int, match.groups())
        return f'{year:04d}-{month:02d}-{day:02d}'

    match = re.fullmatch(r'(\d{1,2})/(\d{1,2})', text)
    if match:
        year = default_year or datetime.now().year
        month, day = map(int, match.groups())
        return f'{year:04d}-{month:02d}-{day:02d}'

    return normalize_text(value)


def normalize_class_part(value: object) -> str:
    return normalize_token(value)


def normalize_class_name(year: object, class_value: object) -> str:
    year_text = normalize_token(year)
    class_text = normalize_class_part(class_value)
    if not year_text or not class_text:
        return ''
    return f'{year_text}_{class_text}'


def denormalize_class_key(class_name: str) -> str:
    """1_1 -> 1 1 のように既存summaryに近い表記へ戻す。"""
    normalized = normalize_token(class_name)
    if '_' not in normalized:
        return normalized
    year, cls = normalized.split('_', 1)
    return f'{year} {cls}'


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

    return validate_input_path(raw)


def validate_input_path(raw_path: str | Path) -> Path:
    raw = str(raw_path).strip().strip('"').strip("'")
    if not raw:
        raise ValueError('ファイルのパスが入力されていません。')

    path = Path(raw).expanduser()
    if not path.exists():
        raise FileNotFoundError(f'ファイルが見つかりません: {path}')
    if not path.is_file():
        raise ValueError(f'ファイルではありません: {path}')
    if path.suffix.lower() not in ('.xlsx', '.csv'):
        raise UnsupportedInputFormatError(f'.xlsx または .csv ファイルを指定してください: {path.name}')

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


def load_sheet_path(
    zf: zipfile.ZipFile,
    preferred_sheet_name: str,
    strict_sheet: bool = False,
) -> tuple[str, str]:
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
        raise SheetNotFoundError('Excel内のシートを読み取れませんでした。')

    for name, path in sheets:
        if name == preferred_sheet_name:
            return name, path

    if strict_sheet:
        available = ', '.join(name for name, _ in sheets)
        raise SheetNotFoundError(
            f'シート「{preferred_sheet_name}」が見つかりません。利用可能なシート: {available}'
        )

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


def read_sheet_rows(
    xlsx_path: Path,
    sheet_name: str = DEFAULT_SHEET_NAME,
    strict_sheet: bool = False,
) -> tuple[str, list[list[str]]]:
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings = load_shared_strings(zf)
        actual_sheet_name, sheet_path = load_sheet_path(zf, sheet_name, strict_sheet=strict_sheet)
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
        stripped = [normalize_text(value) for value in row]
        if all(header in stripped for header in REQUIRED_HEADERS):
            return index
    raise RequiredHeaderNotFoundError(f'見出し行が見つかりません。必要な見出し: {", ".join(REQUIRED_HEADERS)}')


def extract_table_from_sheet_rows(rows: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    header_index = find_header_row_index(rows)
    headers = [normalize_text(value) if normalize_text(value) != '時限' else str(value) for value in rows[header_index]]
    col_count = len(headers)
    date_col = next((i for i, header in enumerate(headers) if header.strip() == DATE_COL), None)

    output_rows: list[list[str]] = []
    for row in rows[header_index + 1:]:
        fixed = list(row[:col_count]) + [''] * max(0, col_count - len(row))
        fixed = [normalize_number_text(str(value)) for value in fixed]

        # 完全な空行は無視する。下に追加された新規データはそのまま拾う。
        if not any(value.strip() for value in fixed):
            continue

        if date_col is not None and date_col < len(fixed):
            fixed[date_col] = excel_serial_to_date_text(fixed[date_col])

        output_rows.append(fixed)

    return headers, output_rows


def write_csv(path: Path, headers: list[str], rows: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


def convert_xlsx_to_csv(
    xlsx_path: Path,
    output_dir: Path | None = None,
    strict_sheet: bool = False,
) -> Path:
    sheet_name, rows = read_sheet_rows(xlsx_path, strict_sheet=strict_sheet)
    headers, output_rows = extract_table_from_sheet_rows(rows)

    if output_dir is None:
        out_csv = xlsx_path.with_suffix('.csv')
    else:
        output_dir.mkdir(parents=True, exist_ok=True)
        out_csv = output_dir / f'{xlsx_path.stem}.csv'

    write_csv(out_csv, headers, output_rows)

    print(f'読み取りシート: {sheet_name}')
    print(f'通常CSV: {out_csv}')
    print(f'通常CSV行数: {len(output_rows)}件')
    return out_csv


# =============================
# 通常CSV -> クラス別CSV
# =============================

def expand_years(value: str) -> list[str]:
    """1～3 のような学年範囲を ['1', '2', '3'] にする。"""
    v = normalize_text(value)
    m = re.fullmatch(r'(\d+)\s*[～〜~-]\s*(\d+)', v)
    if m:
        start, end = map(int, m.groups())
        step = 1 if start <= end else -1
        return [str(i) for i in range(start, end + step, step)]
    return [normalize_token(v)] if v else []


def split_class_cell(value: str) -> list[str]:
    """1,2 のような複数クラス指定を ['1', '2'] にする。全はここでは展開しない。"""
    v = normalize_text(value)
    if not v or normalize_token(v) == '全':
        return []
    return [normalize_class_part(part) for part in re.split(r'[,，、]', v) if normalize_class_part(part)]


def sort_key(class_key: str) -> tuple[int, str]:
    normalized = normalize_token(class_key).replace(' ', '_')
    if '_' in normalized:
        year, cls = normalized.split('_', 1)
    else:
        year, cls = normalized, ''
    return (int(year) if year.isdigit() else 999, cls)


def safe_csv_name(class_name: str) -> str:
    return normalize_token(class_name).replace('/', '_').replace('\\', '_')


def split_csv_by_class(
    src: Path,
    out_dir: Path | None = None,
    zip_path: Path | None = None,
    overwrite: bool = True,
    create_zip: bool = True,
) -> tuple[Path, Path | None, list[list[str]]]:
    out_dir = out_dir or (src.parent / f'{src.stem}_class_csvs')
    default_zip_path = src.parent / f'{src.stem}_class_csvs.zip'
    resolved_zip_path = zip_path or default_zip_path

    with src.open('r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames
        rows = list(reader)

    if headers is None:
        raise RequiredHeaderNotFoundError('CSVのヘッダーを読み取れませんでした。')
    if YEAR_COL not in headers or CLASS_COL not in headers:
        raise RequiredHeaderNotFoundError(f'CSVに必要な列がありません。必要な列: {YEAR_COL}, {CLASS_COL}')

    known_classes_by_year: dict[str, set[str]] = defaultdict(set)
    for row in rows:
        for year in expand_years(row.get(YEAR_COL, '')):
            for cls in split_class_cell(row.get(CLASS_COL, '')):
                known_classes_by_year[year].add(cls)

    by_class: dict[str, list[tuple[int, dict[str, str]]]] = defaultdict(list)

    for idx, row in enumerate(rows):
        years = expand_years(row.get(YEAR_COL, ''))
        class_cell = normalize_text(row.get(CLASS_COL, ''))

        for year in years:
            if normalize_token(class_cell) == '全':
                target_classes = sorted(
                    cls for cls in known_classes_by_year.get(year, set())
                    if cls not in EXCLUDED_FROM_ALL
                )
            else:
                target_classes = split_class_cell(class_cell)

            for cls in target_classes:
                key = normalize_class_name(year, cls)
                if key:
                    by_class[key].append((idx, row.copy()))

    if out_dir.exists():
        if not overwrite:
            raise FileExistsError(f'出力先フォルダが既に存在します: {out_dir}')
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    summary: list[list[str]] = []
    for key in sorted(by_class.keys(), key=sort_key):
        path = out_dir / f'{safe_csv_name(key)}.csv'
        entries = sorted(by_class[key], key=lambda item: item[0])

        with path.open('w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            for _, row in entries:
                writer.writerow(row)

        summary.append([denormalize_class_key(key), str(len(entries)), path.name])

    with (out_dir / '_summary.csv').open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['クラス', '件数', 'ファイル名'])
        writer.writerows(summary)

    if resolved_zip_path.exists() and (overwrite or create_zip):
        if not overwrite and create_zip:
            raise FileExistsError(f'ZIPが既に存在します: {resolved_zip_path}')
        resolved_zip_path.unlink()

    if create_zip:
        with zipfile.ZipFile(resolved_zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
            for file in sorted(out_dir.iterdir()):
                zf.write(file, arcname=file.name)
        return out_dir, resolved_zip_path, summary

    return out_dir, None, summary


# =============================
# 通知Bot向け標準化出力
# =============================

def find_value(row: dict[str, str], candidates: list[str]) -> str:
    normalized_candidates = {normalize_token(candidate) for candidate in candidates}
    for key, value in row.items():
        if normalize_token(key) in normalized_candidates:
            return normalize_text(value)
    return ''


def make_raw_text(row: dict[str, str], headers: list[str]) -> str:
    parts = []
    for header in headers:
        value = normalize_text(row.get(header, ''))
        if value:
            parts.append(f'{header}:{value}')
    return ' | '.join(parts)


def make_canonical_text(values: list[object]) -> str:
    return ' | '.join(normalize_text(value) for value in values if normalize_text(value))


def normalize_period(value: object) -> str:
    text = normalize_text(value)
    text = re.sub(r'(時限|限目|限)$', '', text)
    return normalize_number_text(text)


def apply_change_content_fallback(
    row: dict[str, str],
    before_subject: str,
    after_subject: str,
    note: str,
) -> tuple[str, str, str]:
    """Map 変更内容/科目(担当教員) without inventing an arrow pair."""
    change_content = find_value(row, ["変更内容", "変更種別", "種別"])
    subject_with_teacher = find_value(
        row,
        ["科目(担当教員)", "科目（担当教員）", "科目・担当教員", "科目"],
    )
    if change_content and not note:
        note = change_content
    if not subject_with_teacher or before_subject or after_subject:
        return before_subject, after_subject, note
    if normalize_token(change_content) == "休講":
        return subject_with_teacher, "", note
    return "", subject_with_teacher, note


def build_normalized_records(
    rows: list[dict[str, str]],
    headers: list[str],
    default_year: int | None = None,
) -> list[dict[str, str]]:
    known_classes_by_year: dict[str, set[str]] = defaultdict(set)
    for row in rows:
        for year in expand_years(row.get(YEAR_COL, '')):
            for cls in split_class_cell(row.get(CLASS_COL, '')):
                known_classes_by_year[year].add(cls)

    records: list[dict[str, str]] = []
    for row in rows:
        years = expand_years(row.get(YEAR_COL, ''))
        class_cell = normalize_text(row.get(CLASS_COL, ''))
        raw_text = make_raw_text(row, headers)

        date_text = normalize_date_value(row.get(DATE_COL, ''), default_year=default_year)
        period = normalize_period(find_value(row, ['時限', '校時', '時間', '限']))
        before_subject = find_value(row, ['変更前', '変更前科目', '変更前 科目', '旧科目', '変更元'])
        after_subject = find_value(row, ['変更後', '変更後科目', '変更後 科目', '新科目', '変更先'])
        teacher = find_value(row, ['教員', '担当', '担当教員', '担任', '教官'])
        room = find_value(row, ['教室', '場所'])
        note = find_value(row, ['備考', '連絡', 'メモ', 'その他'])
        before_subject, after_subject, note = apply_change_content_fallback(
            row, before_subject, after_subject, note
        )

        for year in years:
            if normalize_token(class_cell) == '全':
                target_classes = sorted(
                    cls for cls in known_classes_by_year.get(year, set())
                    if cls not in EXCLUDED_FROM_ALL
                )
            else:
                target_classes = split_class_cell(class_cell)

            for cls in target_classes:
                class_name = normalize_class_name(year, cls)
                if not class_name:
                    continue
                canonical_text = make_canonical_text([
                    date_text,
                    class_name,
                    period,
                    before_subject,
                    after_subject,
                    teacher,
                    room,
                    note,
                    raw_text,
                ])
                records.append({
                    'change_date': date_text,
                    'class_name': class_name,
                    'period': period,
                    'before_subject': before_subject,
                    'after_subject': after_subject,
                    'teacher': teacher,
                    'room': room,
                    'note': note,
                    'raw_text': raw_text,
                    'canonical_text': canonical_text,
                })

    return records


def read_source_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames
        if headers is None:
            raise RequiredHeaderNotFoundError('CSVのヘッダーを読み取れませんでした。')
        rows = list(reader)

    if YEAR_COL not in headers or CLASS_COL not in headers or DATE_COL not in headers:
        raise RequiredHeaderNotFoundError(
            f'CSVに必要な列がありません。必要な列: {YEAR_COL}, {CLASS_COL}, {DATE_COL}'
        )
    return headers, rows


def read_source_rows_from_input(
    input_path: Path,
    strict_sheet: bool,
) -> tuple[str | None, list[str], list[dict[str, str]]]:
    suffix = input_path.suffix.lower()

    if suffix == '.xlsx':
        sheet_name, sheet_rows = read_sheet_rows(input_path, strict_sheet=strict_sheet)
        headers, table_rows = extract_table_from_sheet_rows(sheet_rows)
        rows = [dict(zip(headers, row)) for row in table_rows]
        return sheet_name, headers, rows

    if suffix == '.csv':
        headers, rows = read_source_csv(input_path)
        return None, headers, rows

    raise UnsupportedInputFormatError(f'.xlsx または .csv ファイルを指定してください: {input_path.name}')


def write_dict_csv(path: Path, headers: list[str], rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def convert_to_class_csvs(
    input_path: str | Path,
    output_dir: str | Path | None = None,
    strict_sheet: bool = True,
    default_year: int | None = None,
    overwrite: bool = True,
    create_zip: bool = True,
) -> ConversionResult:
    """通知Botなど外部プログラムから安全に呼べる標準化変換API。"""
    path = validate_input_path(input_path)
    out = Path(output_dir).expanduser().resolve() if output_dir is not None else path.parent.resolve()

    out.mkdir(parents=True, exist_ok=True)

    normalized_csv_path = out / 'normalized.csv'
    classes_dir = out / 'classes'
    summary_csv_path = out / 'summary.csv'
    manifest_path = out / 'manifest.json'
    expected_zip_path = out / 'classes.zip'
    zip_path = expected_zip_path if create_zip else None

    if overwrite:
        for file_path in (normalized_csv_path, summary_csv_path, manifest_path, expected_zip_path):
            if file_path.exists():
                file_path.unlink()
        if classes_dir.exists():
            shutil.rmtree(classes_dir)
    else:
        existing_targets = [normalized_csv_path, summary_csv_path, manifest_path, classes_dir]
        if create_zip:
            existing_targets.append(expected_zip_path)
        existing = [p for p in existing_targets if p.exists()]
        if existing:
            raise FileExistsError(f'出力先に既存ファイルがあります: {", ".join(str(p) for p in existing)}')

    sheet_name, source_headers, source_rows = read_source_rows_from_input(path, strict_sheet=strict_sheet)
    records = build_normalized_records(source_rows, source_headers, default_year=default_year)

    write_dict_csv(normalized_csv_path, NORMALIZED_HEADERS, records)

    by_class: dict[str, list[dict[str, str]]] = defaultdict(list)
    for record in records:
        by_class[record['class_name']].append(record)

    classes_dir.mkdir(parents=True, exist_ok=True)
    class_files: list[ClassCsvFile] = []
    for class_name in sorted(by_class.keys(), key=sort_key):
        rows = by_class[class_name]
        class_path = classes_dir / f'{safe_csv_name(class_name)}.csv'
        write_dict_csv(class_path, NORMALIZED_HEADERS, rows)
        class_files.append(ClassCsvFile(class_name=class_name, path=class_path, rows=len(rows)))

    with summary_csv_path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['class_name', 'rows', 'filename'])
        for item in class_files:
            writer.writerow([item.class_name, item.rows, item.path.name])

    if create_zip:
        with zipfile.ZipFile(expected_zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
            zf.write(normalized_csv_path, arcname=normalized_csv_path.name)
            zf.write(summary_csv_path, arcname=summary_csv_path.name)
            for item in class_files:
                zf.write(item.path, arcname=f'classes/{item.path.name}')

    warnings: list[str] = []
    if path.suffix.lower() == '.csv':
        warnings.append('CSV入力のためsheet_nameはありません。')
    if default_year is None:
        warnings.append('年なし日付は実行時の現在年で補完します。')

    manifest = {
        'input_path': str(path),
        'output_dir': str(out),
        'sheet_name': sheet_name,
        'strict_sheet': strict_sheet,
        'generated_at': datetime.now().astimezone().isoformat(timespec='seconds'),
        'normalized_csv': normalized_csv_path.name,
        'classes_dir': classes_dir.name,
        'summary_csv': summary_csv_path.name,
        'zip_path': zip_path.name if zip_path is not None else None,
        'class_files': [
            {
                'class_name': item.class_name,
                'path': f'{classes_dir.name}/{item.path.name}',
                'rows': item.rows,
            }
            for item in class_files
        ],
        'warnings': warnings,
    }
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + '\n',
        encoding='utf-8',
    )

    return ConversionResult(
        input_path=path,
        output_dir=out,
        normalized_csv_path=normalized_csv_path,
        classes_dir=classes_dir,
        summary_csv_path=summary_csv_path,
        manifest_path=manifest_path,
        zip_path=zip_path,
        sheet_name=sheet_name,
        class_files=class_files,
        warnings=warnings,
    )


def result_to_jsonable(result: ConversionResult) -> dict:
    data = asdict(result)
    for key in ['input_path', 'output_dir', 'normalized_csv_path', 'classes_dir', 'summary_csv_path', 'manifest_path', 'zip_path']:
        if data[key] is not None:
            data[key] = str(data[key])
    data['class_files'] = [
        {
            'class_name': item.class_name,
            'path': str(item.path),
            'rows': item.rows,
        }
        for item in result.class_files
    ]
    return data


def print_conversion_result(result: ConversionResult) -> None:
    print(f'標準CSV: {result.normalized_csv_path}')
    print(f'クラス別CSVフォルダ: {result.classes_dir}')
    print(f'サマリーCSV: {result.summary_csv_path}')
    if result.zip_path is not None:
        print(f'ZIP: {result.zip_path}')
    else:
        print('ZIP: 作成しませんでした')
    print(f'manifest: {result.manifest_path}')
    if result.sheet_name:
        print(f'読み取りシート: {result.sheet_name}')
    print('')
    print('class_name, rows, filename')
    for item in result.class_files:
        print(f'{item.class_name}, {item.rows}, {item.path.name}')
    if result.warnings:
        print('')
        print('warnings:')
        for warning in result.warnings:
            print(f'- {warning}')


def parse_args(argv: list[str]) -> argparse.Namespace:
    if argv and argv[0] == 'convert':
        argv = argv[1:]

    parser = argparse.ArgumentParser(
        description='時間割変更Excel/CSVを通知Bot向け標準CSVとクラス別CSVに変換します。'
    )
    parser.add_argument('input_path', nargs='?', help='変換する .xlsx または .csv ファイル')
    parser.add_argument('--output-dir', help='出力先ディレクトリ。未指定なら入力ファイルと同じディレクトリ')
    parser.add_argument('--default-year', type=int, help='年なし月日を補完する年。未指定なら現在年')
    parser.add_argument(
        '--strict-sheet',
        action=argparse.BooleanOptionalAction,
        default=True,
        help='Excelで「時間割変更」シートが無い場合にエラーにする。デフォルトは有効',
    )
    parser.add_argument(
        '--overwrite',
        action=argparse.BooleanOptionalAction,
        default=True,
        help='既存出力を上書きする。デフォルトは有効',
    )
    parser.add_argument(
        '--zip',
        dest='create_zip',
        action=argparse.BooleanOptionalAction,
        default=True,
        help='ZIPを作成する。デフォルトは有効。無効化する場合は --no-zip',
    )
    parser.add_argument(
        '--legacy',
        action='store_true',
        help='旧形式の通常CSVとクラス別CSVを作成する',
    )
    parser.add_argument(
        '--json',
        action='store_true',
        help='変換結果をJSONで標準出力へ出す',
    )
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args(sys.argv[1:])

    if args.input_path:
        src = validate_input_path(args.input_path)
    else:
        src = get_input_path()

    if args.legacy:
        if src.suffix.lower() == '.xlsx':
            csv_path = convert_xlsx_to_csv(
                src,
                output_dir=Path(args.output_dir).expanduser() if args.output_dir else None,
                strict_sheet=args.strict_sheet,
            )
        else:
            csv_path = src
            print(f'通常CSV変換はスキップしました: {csv_path}')

        out_dir = Path(args.output_dir).expanduser() / f'{csv_path.stem}_class_csvs' if args.output_dir else None
        zip_path = Path(args.output_dir).expanduser() / f'{csv_path.stem}_class_csvs.zip' if args.output_dir else None
        class_dir, zip_file, summary = split_csv_by_class(
            csv_path,
            out_dir=out_dir,
            zip_path=zip_path,
            overwrite=args.overwrite,
            create_zip=args.create_zip,
        )

        print('')
        print(f'{len(summary)}個のクラス別CSVを作成しました。')
        print(f'出力フォルダ: {class_dir}')
        if zip_file is not None:
            print(f'ZIP: {zip_file}')
        else:
            print('ZIP: 作成しませんでした')
        print('')
        print('クラス, 件数, ファイル名')
        for class_name, count, filename in summary:
            print(f'{class_name}, {count}, {filename}')
        return

    result = convert_to_class_csvs(
        input_path=src,
        output_dir=args.output_dir,
        strict_sheet=args.strict_sheet,
        default_year=args.default_year,
        overwrite=args.overwrite,
        create_zip=args.create_zip,
    )

    if args.json:
        print(json.dumps(result_to_jsonable(result), ensure_ascii=False, indent=2))
    else:
        print_conversion_result(result)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'エラー: {e}', file=sys.stderr)
        sys.exit(1)
