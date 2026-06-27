import sys
from pathlib import Path

from excel_to_class_schedule_csvs import split_csv_by_class


def get_input_csv_path() -> Path:
    """
    入力CSVを決める。
    1. コマンドライン引数があれば、それを優先する
       例: python split_schedule_csv_by_class.py henkou.csv
       WindowsではCSVをこの.pyにドラッグ&ドロップしても argv[1] に入る
    2. 引数がなければ、入力してもらう
    """
    if len(sys.argv) >= 2:
        raw = sys.argv[1]
    else:
        raw = input('分割したいCSVファイルのパスを入力してください（ドラッグ&ドロップ可）: ')

    raw = raw.strip().strip('"').strip("'")
    if not raw:
        raise ValueError('CSVファイルのパスが入力されていません。')

    path = Path(raw).expanduser()
    if not path.exists():
        raise FileNotFoundError(f'CSVファイルが見つかりません: {path}')
    if not path.is_file():
        raise ValueError(f'ファイルではありません: {path}')
    if path.suffix.lower() != '.csv':
        raise ValueError(f'.csv ファイルを指定してください: {path.name}')

    return path.resolve()


def main() -> None:
    src = get_input_csv_path()
    out_dir, zip_path, summary = split_csv_by_class(src)

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
