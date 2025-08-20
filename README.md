# Excel Formula Extractor (xlwings)

A Python library to extract Excel formulas, validations, hyperlinks, notes, and rich cell metadata using xlwings. Provides a simple Python API and command-line tools.

## Install

- Requirements: Windows, Microsoft Excel, Python 3.8+
- Install deps in your project: `pip install xlwings`
- This repo includes a ready-to-package library under `excel_extractor/`.
- After publishing: `pip install excel-extractor`

### Cross-platform (openpyxl engine)

- On macOS/Linux or Windows without Excel, use the `openpyxl` engine:
  - CLI: add `--engine openpyxl`
  - Programmatic: use `from excel_extractor import OpenpyxlExcelExtractor`
  - Note: display text and live calc values are limited because openpyxl does not evaluate formulas.

## CLI

- Basic (all formulas on active sheet):
```bash
python -m excel_extractor "Workbook.xlsx"  # defaults to xlwings on Windows, openpyxl elsewhere
```
- Force engine:
```bash
python -m excel_extractor "Workbook.xlsx" --engine openpyxl
python -m excel_extractor "Workbook.xlsx" --engine xlwings
```
- Specific worksheet:
```bash
python -m excel_extractor "Workbook.xlsx" --sheet "Sheet1"
```
- Range only:
```bash
python -m excel_extractor "Workbook.xlsx" --range "A1:D10"
```
- Formula dependencies for a cell:
```bash
python -m excel_extractor "Workbook.xlsx" --dependencies "B5"
```
- Full details (formatting, validations, hyperlinks, notes):
```bash
python -m excel_extractor "Workbook.xlsx" --full
```
- Full details for all sheets:
```bash
python -m excel_extractor "Workbook.xlsx" --full --all-sheets
```
- Text output instead of JSON:
```bash
python -m excel_extractor "Workbook.xlsx" --format text
```

- Convert a previously generated `*_full_details.json` into per-sheet CSV/JSON and index:
```bash
python -m excel_extractor.convert_excel_json "Workbook_full_details.json" --out exports --ndjson
```

## Python API

```python
from excel_extractor import ExcelFormulaExtractor, OpenpyxlExcelExtractor

# 1) Windows + Excel (xlwings)
with ExcelFormulaExtractor("Workbook.xlsx") as extractor:
    data = extractor.extract_sheet_full_details("Sheet1")

# 2) Cross-platform (openpyxl)
with OpenpyxlExcelExtractor("Workbook.xlsx") as extractor:
    data = extractor.extract_sheet_full_details("Sheet1")
```

## Public API (summary)

- Class `ExcelFormulaExtractor(excel_file_path: str)` (xlwings)
- Class `OpenpyxlExcelExtractor(excel_file_path: str)` (openpyxl)
  - Context manager: opens/quits workbook automatically
  - `get_worksheet_info(sheet_name: Optional[str]) -> dict`
  - `extract_formulas_from_range(start_cell: str, end_cell: Optional[str]) -> list[dict]`
  - `extract_all_formulas(sheet_name: Optional[str]) -> dict`
  - `extract_sheet_full_details(sheet_name: Optional[str]) -> dict`
  - `extract_workbook_full_details() -> dict`
  - `extract_formula_dependencies(cell_address: str) -> dict`
  - `export_to_json(data: dict, output_file: str) -> bool`
  - `export_to_text(data: dict, output_file: str) -> bool`

- Console scripts (after packaging):
  - `excel-extractor` → same as `python -m excel_extractor`
  - `excel-extractor-convert` → same as `python -m excel_extractor.convert_excel_json`

- Programmatic converter (optional):
  - `excel_extractor.tools.convert_full_details_json(input_json: Path, output_dir: Path, make_ndjson: bool) -> None`

## Notes

- Full-detail extraction returns, per cell: value, formula, display text (xlwings only), basic formatting (number format, font, alignment), fill color, hyperlink, note/comment, data validation (including resolved list items when possible), and merge info.
- xlwings automation requires local Excel. Set the app visible for debugging by editing the code path that creates `xw.App(visible=False)`.

## License

MIT License. See `LICENSE`. 