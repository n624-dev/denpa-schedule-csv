# denpa-schedule-csv

電波ポータルの時間割変更ExcelをCSV化し、クラスごとのCSVに自動分割するためのPythonスクリプト集です。

## できること

- Excel（`.xlsx`）の「時間割変更」シートを通常CSVに変換
- 通常CSVをクラス別CSVに分割
- Excelから直接、クラス別CSVとZIPを作成
- Windowsのドラッグ＆ドロップ実行に対応
- `1,2` のような複数クラス指定に対応
- `1～3` のような学年範囲に対応
- `全` は同じ学年の各クラスに展開
- ただし `AI` は `全` に含めない

## ファイル

| ファイル | 用途 |
|---|---|
| `excel_to_schedule_csv.py` | Excelから通常CSVを作る |
| `split_schedule_csv_by_class.py` | 通常CSVからクラス別CSVを作る |
| `excel_to_class_schedule_csvs.py` | ExcelまたはCSVからクラス別CSVまで一括作成する |

基本的には `excel_to_class_schedule_csvs.py` を使えばOKです。

## 使い方

### Excelからクラス別CSVを一括作成

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx
```

Windowsでは、`henkou.xlsx` を `excel_to_class_schedule_csvs.py` にドラッグ＆ドロップしても使えます。

出力例:

```text
henkou.csv
henkou_class_csvs/
henkou_class_csvs.zip
```

### Excelから通常CSVだけ作成

```bash
python excel_to_schedule_csv.py henkou.xlsx
```

出力例:

```text
henkou.csv
```

### 通常CSVからクラス別CSVを作成

```bash
python split_schedule_csv_by_class.py henkou.csv
```

出力例:

```text
henkou_class_csvs/
henkou_class_csvs.zip
```

## 入力CSVに必要な列

少なくとも次の列が必要です。

- `学 年`
- `学科・クラス`
- `月日`

クラス別分割では、主に `学 年` と `学科・クラス` を使います。

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

## Python環境

標準ライブラリだけで動作します。`openpyxl` などの追加インストールは不要です。

推奨:

```bash
python --version
```

Python 3.10以降を想定しています。
