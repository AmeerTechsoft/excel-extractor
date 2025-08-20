#!/usr/bin/env python3
"""
Command-line interface for the excel_extractor package.
Usage:
  python -m excel_extractor <excel_file> [options]
"""

import argparse
import platform
from pathlib import Path
from typing import Any, Dict

from .extractor import ExcelFormulaExtractor
from .openpyxl_extractor import OpenpyxlExcelExtractor


def main() -> None:
	parser = argparse.ArgumentParser(description='Extract Excel formulas using xlwings or openpyxl')
	parser.add_argument('excel_file', help='Path to Excel file')
	parser.add_argument('--sheet', '-s', help='Worksheet name (optional)')
	parser.add_argument('--output', '-o', help='Output file path (optional)')
	parser.add_argument('--format', '-f', choices=['json', 'text'], default='json', 
				   help='Output format (default: json)')
	parser.add_argument('--range', '-r', help='Cell range (e.g., A1:D10)')
	parser.add_argument('--dependencies', '-d', help='Extract dependencies for specific cell')
	parser.add_argument('--full', action='store_true', help='Extract full details (all cells, formatting, validations, hyperlinks, comments)')
	parser.add_argument('--all-sheets', action='store_true', help='Process all sheets (ignored for range/dependencies modes)')
	parser.add_argument('--engine', choices=['xlwings', 'openpyxl'], help='Backend engine to use')
	args = parser.parse_args()

	# Choose engine: default xlwings on Windows, openpyxl elsewhere
	default_engine = 'xlwings' if platform.system().lower().startswith('win') else 'openpyxl'
	engine = args.engine or default_engine

	ExtractorCls = ExcelFormulaExtractor if engine == 'xlwings' else OpenpyxlExcelExtractor

	with ExtractorCls(args.excel_file) as extractor:
		if args.full:
			if args.all_sheets:
				result: Dict[str, Any] = extractor.extract_workbook_full_details()
			else:
				result = extractor.extract_sheet_full_details(args.sheet)
		elif args.dependencies:
			result = extractor.extract_formula_dependencies(args.dependencies)
		elif args.range:
			if ':' in args.range:
				start, end = args.range.split(':')
				result = extractor.extract_formulas_from_range(start, end)
			else:
				result = extractor.extract_formulas_from_range(args.range)
		else:
			result = extractor.extract_all_formulas(args.sheet)

		if args.output:
			output_file = args.output
		else:
			base_name = Path(args.excel_file).stem
			if args.format == 'json':
				suffix = "_full_details.json" if args.full else "_formulas.json"
				output_file = f"{base_name}{suffix}"
			else:
				output_file = f"{base_name}_formulas.txt"

		if args.format == 'json':
			extractor.export_to_json(result, output_file)
		else:
			extractor.export_to_text(result, output_file)

	print("\nExtraction completed successfully!")
	print(f"Output file: {output_file}")


if __name__ == "__main__":
	main() 