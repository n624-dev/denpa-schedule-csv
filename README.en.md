# denpa-schedule-csv

[日本語版 README](README.md)

A Python script collection for converting timetable change Excel files into standardized CSV files and class-specific CSV files.

It is designed so that external programs can safely consume generated outputs including standardized CSV files, class-specific CSV files, and `manifest.json`.

## Features

- Convert `.xlsx` timetable change sheets to CSV
- Generate standardized CSV files for external integrations
- Split standardized CSV files into class-specific CSV files
- Generate `manifest.json` for machine-readable output metadata
- Specify the output directory
- Optionally disable ZIP output
- Support drag-and-drop execution on Windows
- Support multiple class values such as `1,2`
- Support grade ranges such as `1～3`
- Expand `全` to each class in the same grade
- Exclude `AI` from `全`
- Normalize class names such as `3_IT` and `1_1`
- Normalize dates to `YYYY-MM-DD`

## Files

| File | Purpose |
|---|---|
| `excel_to_schedule_csv.py` | Legacy: convert Excel files to regular CSV files |
| `split_schedule_csv_by_class.py` | Legacy: split regular CSV files by class |
| `excel_to_class_schedule_csvs.py` | Recommended: convert Excel/CSV files and generate standardized CSV, class CSV files, and manifest in one step |

In most cases, use `excel_to_class_schedule_csvs.py`.

## Recommended CLI

### Convert an Excel file and generate all outputs

```bash
python excel_to_class_schedule_csvs.py convert henkou.xlsx --output-dir work/out --default-year 2026
```

The `convert` subcommand can be omitted.

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --default-year 2026
```

Example output:

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

### ZIP output

By default, `classes.zip` is created.

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --zip
```

If ZIP output is not needed, pass `--no-zip`.

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --no-zip
```

When `--no-zip` is used, `classes.zip` is not created and `zip_path` in `manifest.json` is set to `null`.

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

Importing programs can read `class_files` in `manifest.json` instead of relying on the ZIP file, so imports can work regardless of whether ZIP output is enabled.

### strict-sheet

By default, `--strict-sheet` is enabled.
If the Excel file does not contain a sheet named `時間割変更`, the conversion fails.

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --strict-sheet
```

To fall back to the first sheet when the target sheet is not found, use:

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --no-strict-sheet
```

For external integrations, keep `--strict-sheet` enabled.

### Dates without a year

If a date value is written without a year, such as `6/28`, `--default-year` can be used to fill in the year.

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --default-year 2026
```

If `--default-year` is omitted, the current year at runtime is used.

### JSON output

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --output-dir work/out --json
```

## Standardized CSV schema

`normalized.csv` and `classes/*.csv` use the same headers.

```csv
change_date,class_name,period,before_subject,after_subject,teacher,room,note,raw_text,canonical_text
```

| Column | Description |
|---|---|
| `change_date` | Date normalized to `YYYY-MM-DD` |
| `class_name` | Normalized class name such as `3_IT` or `1_1` |
| `period` | Class period |
| `before_subject` | Subject before the change |
| `after_subject` | Subject after the change |
| `teacher` | Teacher or staff member |
| `room` | Classroom or location |
| `note` | Notes |
| `raw_text` | Concatenated non-empty cells from the original row |
| `canonical_text` | Normalized text used by external programs to generate a `change_id` |

Columns such as `period` and `before_subject` are filled when they can be inferred from the source Excel headers.
Even when these fields cannot be extracted, `raw_text` and `canonical_text` still contain the whole row information.

For source files with `変更内容` and `科目(担当教員)` columns, the change type is stored in `note`.
Cancelled subjects are stored in `before_subject`; makeup or changed subjects are stored in
`after_subject`, allowing consumers to render the source facts without inventing a before/after pair.

## manifest.json

`manifest.json` contains a list of generated outputs.

Example:

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

When `--no-zip` is used, `zip_path` becomes `null`.

```json
{
  "zip_path": null
}
```

Consumers can read `class_files` in `manifest.json` to safely determine which CSV files should be imported.

## Python API usage

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
print(result.zip_path)  # None when create_zip=False

for item in result.class_files:
    print(item.class_name, item.rows, item.path)
```

For external program usage, create a dedicated working directory for each attachment and pass its `output/` directory as `output_dir`.

Example:

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

When `create_zip=False` is used, `classes.zip` is not created.

## Legacy output

Use `--legacy` if you need the previous regular CSV and `*_class_csvs/` output format.

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --legacy
```

Example output:

```text
henkou.csv
henkou_class_csvs/
henkou_class_csvs.zip
```

`--legacy --output-dir work/out` can also be used.

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --legacy --output-dir work/out
```

`--no-zip` can also be used with legacy output.

```bash
python excel_to_class_schedule_csvs.py henkou.xlsx --legacy --no-zip
```

## Required input CSV columns

At minimum, the following columns are required.

- `学 年`
- `学科・クラス`
- `月日`

`学 年` must include a half-width space between `学` and `年`.
It is treated differently from `学年` without the space.

## Splitting rules

Examples:

| 学 年 | 学科・クラス | Output |
|---|---|---|
| `1` | `1` | `1_1.csv` |
| `1` | `1,2` | `1_1.csv`, `1_2.csv` |
| `1～3` | `全` | Each class from grades 1 to 3, excluding AI |
| `1` | `AI` | `1_AI.csv` only |
| `2` | `AI` | `2_AI.csv` only |

`1 AI` and `2 AI` are not included in `1 全` or `2 全`.

Class names are normalized for full-width/half-width characters and letter case.

Examples:

| Input | Output |
|---|---|
| `3IT` | `3_IT` |
| `3 it` | `3_IT` |
| `３ＩＴ` | `3_IT` |
| `1 1` | `1_1` |
| `１＿１` | `1_1` |

## Supported formats

The initial implementation supports `.xlsx` and `.csv` files.

`.xls` is not supported.
If `.xls` support is needed, convert the file to `.xlsx` with LibreOffice or another tool first.

## Python environment

This project uses only the Python standard library. Extra packages such as `openpyxl` are not required.

Recommended:

```bash
python --version
```

Python 3.10 or later is expected.

## License

This project is licensed under the MIT License.

The license applies only to the source code in this repository.
Input Excel files, generated CSV files, and actual school timetable change data are not included in this license.
