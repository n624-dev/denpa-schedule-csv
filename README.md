# denpa-schedule-csv

[English README](README.en.md)

電波ポータルの時間割変更ExcelをCSV化し、クラスごとのCSVに自動分割するためのPythonスクリプト集です。

外部プログラムから安全に使えるように、標準化CSV・クラス別CSV・`manifest.json` を指定ディレクトリへ出力できます。

## できること

- Excel（`.xlsx`）の「時間割変更」シートをCSVに変換
- 外部連携向けの標準化CSVを作成
- 標準化済みCSVをクラス別CSVに分割
- `manifest.json` で生成結果を機械的に取得
- 出力先ディレクトリを指定可能
- ZIP出力を任意で無効化可能
- Windowsのドラッグ＆ドロップ実行に対応
- `1,2` のような複数クラス指定に対応
- `1～3` のような学年範囲に対応
- `全` は同じ学年の各クラスに展開
- ただし `AI` は `全` に含めない
- クラス名を `3_IT`, `1_1` のように正規化
- 日付を `YYYY-MM-DD` に正規化

## ファイル

| ファイル | 用途 |
|---|---|
| `excel_to_schedule_csv.py` | 旧形式: Excelから通常CSVを作る |
| `split_schedule_csv_by_class.py` | 旧形式: 通常CSVからクラス別CSVを作る |
| `excel_to_class_schedule_csvs.py` | 推奨: ExcelまたはCSVから標準化CSV・クラス別CSV・manifestまで一括作成する |

基本的には `excel_to_class_schedule_csvs.py` を使えばOKです。

## 推奨CLI

### Excelから標準化CSV・クラス別CSVを一括作成

```bash
python excel_to_class_schedule_csvs.py convert henkou.xlsx --output-dir work/out --default-year 2026
```

`convert` は省略できます。

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --default-year 2026
```

出力例:

```text
work/out/
├─ normalized.csv
├─ classes/
│  ├─ 1_1.csv
│  ├─ 3_IT.csv
│  └─ 5_CN.csv
├─ summary.csv
├─ classes.zip
└─ manifest.json
```

### ZIP出力

デフォルトでは `classes.zip` を作成します。

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --zip
```

ZIPが不要な場合は `--no-zip` を指定します。

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --no-zip
```

`--no-zip` を指定した場合、`classes.zip` は作成されず、`manifest.json` の `zip_path` は `null` になります。

```text
work/out/
├─ normalized.csv
├─ classes/
│  ├─ 1_1.csv
│  ├─ 3_IT.csv
│  └─ 5_CN.csv
├─ summary.csv
└─ manifest.json
```

利用側は、ZIPではなく `manifest.json` の `class_files` を読むことで、ZIPの有無に関係なくCSVを取り込めます。

### strict-sheet

デフォルトでは `--strict-sheet` が有効です。  
Excel内に `時間割変更` シートがない場合はエラーにします。

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --strict-sheet
```

従来のように、見つからなければ先頭シートへフォールバックしたい場合は次を指定します。

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --no-strict-sheet
```

外部連携用途では `--strict-sheet` のまま使ってください。

### 年なし日付

`月日` が `6/28` のように年なしの場合、`--default-year` を使って補完できます。

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --default-year 2026
```

`--default-year` を省略した場合は、実行時の現在年で補完します。

### JSONで結果を受け取る

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --json
```

## 標準化CSVのスキーマ

`normalized.csv` と `classes/*.csv` は同じヘッダーです。

```csv
change_date,class_name,period,before_subject,after_subject,teacher,room,note,raw_text,canonical_text
```

| 列 | 内容 |
|---|---|
| `change_date` | `YYYY-MM-DD` に正規化した日付 |
| `class_name` | `3_IT`, `1_1` のように正規化したクラス名 |
| `period` | 時限 |
| `before_subject` | 変更前科目 |
| `after_subject` | 変更後科目 |
| `teacher` | 教員・担当者 |
| `room` | 教室・場所 |
| `note` | 備考 |
| `raw_text` | 元行の非空セルを連結した文字列 |
| `canonical_text` | 外部プログラムで `change_id` を作るための正規化済み文字列 |

`period` や `before_subject` などは、元Excelの列名から推定できる範囲で埋めます。  
抽出できない場合でも、`raw_text` と `canonical_text` には行全体の情報が入ります。

## manifest.json

`manifest.json` には生成結果の一覧が入ります。

例:

```json
{
  "input_path": "/path/to/henkou.xlsx",
  "output_dir": "/path/to/work/out",
  "sheet_name": "時間割変更",
  "strict_sheet": true,
  "generated_at": "2026-06-27T21:00:00+09:00",
  "normalized_csv": "normalized.csv",
  "classes_dir": "classes",
  "summary_csv": "summary.csv",
  "zip_path": "classes.zip",
  "class_files": [
    {
      "class_name": "3_IT",
      "path": "classes/3_IT.csv",
      "rows": 12
    }
  ],
  "warnings": []
}
```

`--no-zip` を指定した場合は、`zip_path` が `null` になります。

```json
{
  "zip_path": null
}
```

利用側では、`manifest.json` の `class_files` を読めば、取り込むべきCSVを安全に判断できます。

## Python APIとして使う

```python
from pathlib import Path
from excel_to_class_schedule_csvs import convert_to_class_csvs

result = convert_to_class_csvs(
    input_path=Path("downloads/job_xxx/input/henkou.xlsx"),
    output_dir=Path("downloads/job_xxx/output"),
    strict_sheet=True,
    default_year=2026,
    overwrite=True,
    create_zip=False,
)

print(result.normalized_csv_path)
print(result.classes_dir)
print(result.manifest_path)
print(result.zip_path)  # create_zip=False の場合は None

for item in result.class_files:
    print(item.class_name, item.rows, item.path)
```

外部プログラムで使う場合は、添付ファイルごとに専用作業ディレクトリを作り、その中の `output/` を `output_dir` に指定してください。

例:

```text
downloads/
└─ uid12345_ab12cd34/
   ├─ input/
   │  └─ henkou.xlsx
   └─ output/
      ├─ normalized.csv
      ├─ classes/
      ├─ summary.csv
      ├─ classes.zip
      └─ manifest.json
```

`create_zip=False` の場合、`classes.zip` は作成されません。

## 旧形式で出力する

以前の通常CSVと `*_class_csvs/` 形式が必要な場合は `--legacy` を使えます。

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --legacy
```

出力例:

```text
henkou.csv
henkou_class_csvs/
henkou_class_csvs.zip
```

`--legacy --output-dir work/out` も使えます。

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --legacy --output-dir work/out
```

旧形式でも `--no-zip` を指定できます。

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --legacy --no-zip
```

## 入力CSVに必要な列

少なくとも次の列が必要です。

- `学 年`
- `学科・クラス`
- `月日`

`学 年` は `学` + 半角スペース + `年` です。  
`学年` のようにスペースを削除した名前とは別扱いです。

## 分割ルール

例:

| 学 年 | 学科・クラス | 分割先 |
|---|---|---|
| `1` | `1` | `1_1.csv` |
| `1` | `1,2` | `1_1.csv`, `1_2.csv` |
| `1～3` | `全` | 1〜3年の各クラス。ただしAIは除外 |
| `1` | `AI` | `1_AI.csv` のみ |
| `2` | `AI` | `2_AI.csv` のみ |

`1 AI` や `2 AI` は `1 全` や `2 全` には含めません。

クラス名は全角半角・大文字小文字を正規化します。

例:

| 入力 | 出力 |
|---|---|
| `3IT` | `3_IT` |
| `3 it` | `3_IT` |
| `３ＩＴ` | `3_IT` |
| `1 1` | `1_1` |
| `１＿１` | `1_1` |

## 対応形式

初期実装では `.xlsx` と `.csv` に対応しています。

`.xls` は未対応です。  
`.xls` が必要な場合は、LibreOfficeなどで `.xlsx` に変換してから使ってください。

## Python環境

標準ライブラリだけで動作します。`openpyxl` などの追加インストールは不要です。

推奨:

```bash
python --version
```

Python 3.10以降を想定しています。

## ライセンス

このリポジトリのソースコードは MIT License で公開します。

ただし、入力Excel、生成CSV、実際の時間割変更データなどの学校由来データは、このライセンスの対象外です。
