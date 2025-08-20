#!/usr/bin/env python3
import json
import csv
import re
from pathlib import Path
from typing import Any, Dict, List, Tuple, Optional


def sanitize_filename(name: str) -> str:
	name = name.strip()
	name = re.sub(r"[\\/:*?\"<>|]", "_", name)
	name = re.sub(r"\s+", "_", name)
	return name[:120] or "sheet"


def ensure_dir(path: Path) -> None:
	path.mkdir(parents=True, exist_ok=True)


def load_json(path: Path) -> Dict[str, Any]:
	with path.open("r", encoding="utf-8") as f:
		return json.load(f)


def detect_structure(data: Dict[str, Any]) -> str:
	if "workbook" in data and isinstance(data["workbook"], dict) and "sheets" in data["workbook"]:
		return "workbook"
	if "sheet" in data and "cells" in data:
		return "sheet"
	return "unknown"


def extract_sheets(data: Dict[str, Any]) -> List[Dict[str, Any]]:
	kind = detect_structure(data)
	if kind == "workbook":
		return data.get("workbook", {}).get("sheets", []) or []
	if kind == "sheet":
		return [data]
	return []


def summarize_sheet(sheet_entry: Dict[str, Any]) -> Tuple[str, Dict[str, int]]:
	sheet_meta = sheet_entry.get("sheet", {})
	name = sheet_meta.get("name", "Unknown")
	cells: List[Dict[str, Any]] = sheet_entry.get("cells", []) or []
	formula_count = sum(1 for c in cells if c.get("formula"))
	validation_count = sum(1 for c in cells if c.get("data_validation"))
	dropdown_count = sum(1 for c in cells if (c.get("data_validation") or {}).get("type_name") == "xlValidateList")
	return name, {
		"cells": len(cells),
		"formulas": formula_count,
		"validations": validation_count,
		"dropdowns": dropdown_count,
	}


def write_sheet_csvs(out_dir: Path, sheet_name: str, cells: List[Dict[str, Any]]) -> Dict[str, str]:
	files: Dict[str, str] = {}
	base = sanitize_filename(sheet_name)
	all_csv = out_dir / f"{base}-cells.csv"
	form_csv = out_dir / f"{base}-formulas.csv"
	val_csv = out_dir / f"{base}-validations.csv"

	columns = [
		"address","row","column","column_letter",
		"value","display_text","formula",
		"format.number_format","format.font_name","format.font_size","format.font_bold","format.font_italic",
		"hyperlink.address","note",
		"data_validation.type_name","data_validation.formula1","data_validation.formula2","data_validation.list_items",
		"format.merged","format.merge_area"
	]

	def get_nested(d: Dict[str, Any], path: str) -> Any:
		cur: Any = d
		for part in path.split('.'):
			if not isinstance(cur, dict):
				return None
			cur = cur.get(part)
		return cur

	def write_csv(path: Path, rows: List[Dict[str, Any]]):
		with path.open("w", newline="", encoding="utf-8") as f:
			writer = csv.writer(f)
			writer.writerow(columns)
			for r in rows:
				line: List[Any] = []
				for col in columns:
					val = get_nested(r, col) if "." in col else r.get(col)
					if isinstance(val, list):
						val = ", ".join(str(x) for x in val)
					line.append(val)
				writer.writerow(line)

	# All cells
	write_csv(all_csv, cells)
	files["cells_csv"] = str(all_csv)

	# Formulas only
	formula_rows = [c for c in cells if c.get("formula")]
	write_csv(form_csv, formula_rows)
	files["formulas_csv"] = str(form_csv)

	# Validations only
	validation_rows = [c for c in cells if c.get("data_validation")]
	write_csv(val_csv, validation_rows)
	files["validations_csv"] = str(val_csv)

	return files


def write_sheet_json(out_dir: Path, sheet_name: str, sheet_entry: Dict[str, Any]) -> str:
	base = sanitize_filename(sheet_name)
	path = out_dir / f"{base}.json"
	with path.open("w", encoding="utf-8") as f:
		json.dump(sheet_entry, f, ensure_ascii=False, indent=2, default=str)
	return str(path)


def write_ndjson(out_dir: Path, workbook_name: str, sheets: List[Dict[str, Any]]) -> str:
	path = out_dir / f"{workbook_name}-cells.ndjson"
	with path.open("w", encoding="utf-8") as f:
		for sheet_entry in sheets:
			name = sheet_entry.get("sheet", {}).get("name", "Unknown")
			for cell in sheet_entry.get("cells", []) or []:
				row = {"sheet": name}
				row.update(cell)
				f.write(json.dumps(row, ensure_ascii=False, default=str) + "\n")
	return str(path)


def write_index_md(out_dir: Path, workbook_path: Path, sheets: List[Dict[str, Any]], file_map: Dict[str, Dict[str, str]]) -> str:
	index = out_dir / "INDEX.md"
	lines: List[str] = []
	lines.append(f"# Excel Extraction Index\n")
	lines.append(f"- **source_file**: `{workbook_path.name}`\n")
	lines.append("")
	lines.append("## Sheets\n")
	lines.append("| Sheet | Cells | Formulas | Validations | Dropdowns | CSV (cells) | CSV (formulas) | CSV (validations) | JSON |\n")
	lines.append("|---|---:|---:|---:|---:|---|---|---|---|\n")
	for sheet_entry in sheets:
		name, stats = summarize_sheet(sheet_entry)
		key = sanitize_filename(name)
		files = file_map.get(key, {})
		lines.append(
			f"| {name} | {stats['cells']} | {stats['formulas']} | {stats['validations']} | {stats['dropdowns']} | "
			f"[{Path(files.get('cells_csv','')).name}]({Path(files.get('cells_csv','')).name}) | "
			f"[{Path(files.get('formulas_csv','')).name}]({Path(files.get('formulas_csv','')).name}) | "
			f"[{Path(files.get('validations_csv','')).name}]({Path(files.get('validations_csv','')).name}) | "
			f"[{key}.json]({key}.json) |\n"
		)
	lines.append("")
	lines.append("Generated by convert_excel_json.py\n")
	with index.open("w", encoding="utf-8") as f:
		f.write("\n".join(lines))
	return str(index)


def convert(input_json: Path, output_dir: Path, make_ndjson: bool) -> None:
	data = load_json(input_json)
	sheets = extract_sheets(data)
	if not sheets:
		raise SystemExit("No sheets found in JSON. Ensure you passed a _full_details.json file.")

	workbook_stem = input_json.stem.replace(" ", "_")
	out_root = output_dir / workbook_stem
	ensure_dir(out_root)

	file_map: Dict[str, Dict[str, str]] = {}
	for sheet_entry in sheets:
		name = sheet_entry.get("sheet", {}).get("name", "Unknown")
		sheet_dir = out_root / sanitize_filename(name)
		ensure_dir(sheet_dir)
		csv_files = write_sheet_csvs(sheet_dir, name, sheet_entry.get("cells", []) or [])
		json_file = write_sheet_json(sheet_dir, name, sheet_entry)
		key = sanitize_filename(name)
		file_map[key] = {**csv_files, "json": json_file}

	ndjson_path = write_ndjson(out_root, workbook_stem, sheets) if make_ndjson else None
	index_path = write_index_md(out_root, input_json, sheets, file_map)

	print("Conversion complete.")
	print(f"Output directory: {out_root}")
	print(f"Index: {index_path}")
	if ndjson_path:
		print(f"NDJSON: {ndjson_path}") 