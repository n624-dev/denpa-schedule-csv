import sys
from pathlib import Path

from excel_to_class_schedule_csvs import convert_xlsx_to_csv


def get_input_xlsx_path() -> Path:
    """
    入力XLSXを決める。
    1. コマンドライン引数があれば、それを優先する
       例: python excel_to_schedule_csv.py henkou.xlsx
       WindowsではXLSXをこの.pyにドラッグ&ドロップしても argv[1] に入る
    2. 引数がなければ、入力してもらう
    """
    if len(sys.argv) >= 2:
        raw = sys.argv[1]
    else:
        raw = input('CSVに変換したいExcelファイルのパスを入力してください（ドラッグ&ドロップ可）: ')

    raw = raw.strip().strip('"').strip("'")
    if not raw:
        raise ValueError('Excelファイルのパスが入力されていません。')

    path = Path(raw).expanduser()
    if not path.exists():
        raise FileNotFoundError(f'Excelファイルが見つかりません: {path}')
    if not path.is_file():
        raise ValueError(f'ファイルではありません: {path}')
    if path.suffix.lower() != '.xlsx':
        raise ValueError(f'.xlsx ファイルを指定してください: {path.name}')

    return path.resolve()


def main() -> None:
    src = get_input_xlsx_path()
    convert_xlsx_to_csv(src)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'エラー: {e}', file=sys.stderr)
        sys.exit(1)
