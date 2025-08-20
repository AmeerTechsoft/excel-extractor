#!/usr/bin/env python3
"""
Excel Formula Extractor using xlwings
Extracts formulas, calculations, and cell references from Excel files
Outputs results in JSON or text format
"""

import xlwings as xw
import json
import os
import sys
from typing import Dict, List, Any, Optional
from pathlib import Path
import argparse
from datetime import datetime


class ExcelFormulaExtractor:
	"""Extract formulas and calculations from Excel files using xlwings"""
	
	def __init__(self, excel_file_path: str):
		"""
		Initialize the extractor with an Excel file path
		
		Args:
			excel_file_path (str): Path to the Excel file
		"""
		self.excel_file_path = Path(excel_file_path)
		self.app = None
		self.workbook = None
		self.worksheet = None
		
		if not self.excel_file_path.exists():
			raise FileNotFoundError(f"Excel file not found: {excel_file_path}")
	
	def __enter__(self):
		"""Context manager entry"""
		self.open_workbook()
		return self
	
	def __exit__(self, exc_type, exc_val, exc_tb):
		"""Context manager exit"""
		self.close_workbook()
	
	def open_workbook(self):
		"""Open the Excel workbook using xlwings"""
		try:
			# Start Excel application
			self.app = xw.App(visible=False)
			self.workbook = self.app.books.open(str(self.excel_file_path))
			print(f"Successfully opened: {self.excel_file_path.name}")
		except Exception as e:
			print(f"Error opening workbook: {e}")
			raise
	
	def close_workbook(self):
		"""Close the workbook and Excel application"""
		try:
			if self.workbook:
				self.workbook.close()
			if self.app:
				self.app.quit()
		except Exception as e:
			print(f"Error closing workbook: {e}")
	
	def get_worksheet_info(self, sheet_name: Optional[str] = None) -> Dict[str, Any]:
		"""
		Get basic information about a worksheet
		
		Args:
			sheet_name (str, optional): Name of the worksheet. If None, uses active sheet.
			
		Returns:
			Dict containing worksheet information
		"""
		try:
			if sheet_name:
				self.worksheet = self.workbook.sheets[sheet_name]
			else:
				self.worksheet = self.workbook.sheets.active
			
			# Get used range
			used_range = self.worksheet.used_range
			
			info = {
				"sheet_name": self.worksheet.name,
				"used_range": f"{used_range.address}",
				"rows": used_range.rows.count,
				"columns": used_range.columns.count,
				"total_cells": used_range.rows.count * used_range.columns.count
			}
			
			return info
		except Exception as e:
			print(f"Error getting worksheet info: {e}")
			return {}
	
	def extract_formulas_from_range(self, start_cell: str = "A1", end_cell: Optional[str] = None) -> List[Dict[str, Any]]:
		"""
		Extract formulas from a specific range
		
		Args:
			start_cell (str): Starting cell reference (e.g., "A1")
			end_cell (str, optional): Ending cell reference. If None, extracts from start_cell only.
			
		Returns:
			List of dictionaries containing cell information and formulas
		"""
		try:
			if end_cell:
				range_obj = self.worksheet.range(f"{start_cell}:{end_cell}")
			else:
				range_obj = self.worksheet.range(start_cell)
			
			formulas = []
			
			for cell in range_obj:
				cell_info = self._extract_cell_info(cell)
				if cell_info:
					formulas.append(cell_info)
			
			return formulas
		except Exception as e:
			print(f"Error extracting formulas from range: {e}")
			return []
	
	def extract_all_formulas(self, sheet_name: Optional[str] = None) -> Dict[str, Any]:
		"""
		Extract all formulas from the entire worksheet
		
		Args:
			sheet_name (str, optional): Name of the worksheet. If None, uses active sheet.
			
		Returns:
			Dictionary containing all extracted formulas and metadata
		"""
		try:
			# Get worksheet info
			sheet_info = self.get_worksheet_info(sheet_name)
			
			# Get used range
			used_range = self.worksheet.used_range
			
			# Extract formulas from all cells
			all_formulas = []
			for row in range(1, used_range.rows.count + 1):
				for col in range(1, used_range.columns.count + 1):
					cell = self.worksheet.cells(row, col)
					cell_info = self._extract_cell_info(cell)
					if cell_info:
						all_formulas.append(cell_info)
			
			result = {
				"file_path": str(self.excel_file_path),
				"extraction_timestamp": datetime.now().isoformat(),
				"worksheet_info": sheet_info,
				"total_formulas_found": len(all_formulas),
				"formulas": all_formulas
			}
			
			return result
		except Exception as e:
			print(f"Error extracting all formulas: {e}")
			return {}
	
	def _extract_cell_info(self, cell) -> Optional[Dict[str, Any]]:
		"""
		Extract information from a single cell
		
		Args:
			cell: xlwings cell object
			
		Returns:
			Dictionary containing cell information or None if no formula
		"""
		try:
			# Get cell address
			address = cell.address
			
			# Get cell value
			value = cell.value
			
			# Get formula
			formula = cell.formula
			
			# Get cell format info
			format_info = {
				"number_format": cell.number_format,
				"font_name": cell.font.name,
				"font_size": cell.font.size,
				"font_bold": cell.font.bold,
				"font_italic": cell.font.italic
			}
			
			# Only return cells that have formulas or are part of calculations
			if formula and formula.startswith('='):
				return {
					"address": address,
					"formula": formula,
					"value": value,
					"format": format_info,
					"row": cell.row,
					"column": cell.column,
					"column_letter": self._column_to_letter(cell.column)
				}
			elif isinstance(value, (int, float)) and value != 0:
				# Include numeric values that might be calculation results
				return {
					"address": address,
					"formula": None,
					"value": value,
					"format": format_info,
					"row": cell.row,
					"column": cell.column,
					"column_letter": self._column_to_letter(cell.column),
					"note": "Numeric value (potential calculation result)"
				}
			
			return None
		except Exception as e:
			print(f"Error extracting cell info for {cell.address}: {e}")
			return None
	
	def _column_to_letter(self, column_number: int) -> str:
		"""Convert column number to Excel column letter"""
		result = ""
		while column_number > 0:
			column_number, remainder = divmod(column_number - 1, 26)
			result = chr(65 + remainder) + result
		return result
	
	# --- New: Full-detail extraction helpers and APIs ---
	def _get_cell_display_text(self, cell) -> Optional[str]:
		try:
			return cell.api.Text
		except Exception:
			return None
	
	def _get_cell_fill_color(self, cell) -> Optional[Dict[str, Any]]:
		try:
			color_val = cell.api.Interior.Color
			if color_val is None:
				return None
			# Excel returns BGR as an int; provide raw and decomposed RGB
			b = (int(color_val) >> 16) & 255
			g = (int(color_val) >> 8) & 255
			r = int(color_val) & 255
			return {"excel_bgr": int(color_val), "rgb": {"r": r, "g": g, "b": b}}
		except Exception:
			return None
	
	def _get_cell_hyperlink(self, cell) -> Optional[Dict[str, Any]]:
		try:
			hyperlinks = cell.api.Hyperlinks
			if hyperlinks and hyperlinks.Count > 0:
				link = hyperlinks.Item(1)
				return {
					"address": link.Address,
					"sub_address": getattr(link, "SubAddress", None),
					"text_to_display": getattr(link, "TextToDisplay", None)
				}
		except Exception:
			pass
		return None
	
	def _get_cell_note(self, cell) -> Optional[str]:
		# Try modern note API first via xlwings, fallback to COM Comment
		try:
			note_val = getattr(cell, "note", None)
			if isinstance(note_val, str) and note_val.strip() != "":
				return note_val
		except Exception:
			pass
		try:
			comment = cell.api.Comment
			if comment is not None:
				try:
					return comment.Text()
				except Exception:
					return None
		except Exception:
			pass
		return None
	
	def _get_cell_validation(self, cell) -> Optional[Dict[str, Any]]:
		try:
			v = cell.api.Validation
			# Accessing Type can raise if no validation
			v_type = v.Type
			if v_type is None:
				return None
			type_name_map = {
				0: "xlValidateInputOnly",
				1: "xlValidateWholeNumber",
				2: "xlValidateDecimal",
				3: "xlValidateList",
				4: "xlValidateDate",
				5: "xlValidateTime",
				6: "xlValidateTextLength",
				7: "xlValidateCustom"
			}
			validation: Dict[str, Any] = {
				"type": int(v_type),
				"type_name": type_name_map.get(int(v_type), None),
				"alert_style": int(getattr(v, "AlertStyle", 0)) if getattr(v, "AlertStyle", None) is not None else None,
				"operator": int(getattr(v, "Operator", 0)) if getattr(v, "Operator", None) is not None else None,
				"ignore_blank": bool(getattr(v, "IgnoreBlank", False)),
				"in_cell_dropdown": bool(getattr(v, "InCellDropdown", False)),
				"formula1": getattr(v, "Formula1", None),
				"formula2": getattr(v, "Formula2", None)
			}
			# If it's a list validation, try to resolve list items
			# xlValidateList == 3
			if int(v_type) == 3:
				resolved_list = self._resolve_validation_list_items(cell, validation.get("formula1"))
				if resolved_list is not None:
					validation["list_items"] = resolved_list
			return validation
		except Exception:
			return None

	def _flatten_to_list(self, vals: Any) -> List[Any]:
		flattened: List[Any] = []
		try:
			if isinstance(vals, list):
				for row in vals:
					if isinstance(row, list):
						for itm in row:
							if itm is not None:
								flattened.append(itm)
					else:
						if row is not None:
							flattened.append(row)
			elif vals is not None:
				flattened = [vals]
		except Exception:
			pass
		return flattened

	def _values_from_range_on_sheet(self, sheet, addr: str) -> Optional[List[Any]]:
		try:
			vals = sheet.range(addr).value
			return self._flatten_to_list(vals)
		except Exception:
			return None

	def _evaluate_in_excel(self, expr: str, sheet) -> Any:
		try:
			return sheet.api.Evaluate(expr)
		except Exception:
			try:
				return self.app.api.Evaluate(expr)
			except Exception:
				return None

	def _try_resolve_named_range(self, name: str) -> Optional[List[Any]]:
		try:
			target = name.strip().strip("'")
			for nm in self.workbook.names:
				n = getattr(nm, "name", None)
				if not n:
					continue
				if n.replace(" ", "").lower() == target.replace(" ", "").lower():
					refers_to = getattr(nm, "refers_to", None)
					if not refers_to:
						continue
					f = str(refers_to)
					if f.startswith("="):
						f = f[1:]
					if "!" in f or ":" in f:
						sheet_name = None
						addr = f
						if "!" in f:
							sheet_name, addr = f.split("!", 1)
							sheet_name = sheet_name.strip().strip("'")
						sheet = self.workbook.sheets[sheet_name] if sheet_name else self.worksheet
						return self._values_from_range_on_sheet(sheet, addr)
					# Try evaluate on each sheet context
					for ws in self.workbook.sheets:
						res = self._evaluate_in_excel(f, ws)
						try:
							addr = getattr(res, "Address", None)
						except Exception:
							addr = None
						if addr:
							try:
								vals = ws.range(addr).value
								return self._flatten_to_list(vals)
							except Exception:
								continue
			return None
		except Exception:
			return None

	def _try_resolve_table_column(self, structured_ref: str) -> Optional[List[Any]]:
		# Handle references like Table1[Column] used in validation lists
		try:
			text = structured_ref.strip().strip("'")
			if "[" in text and "]" in text:
				table_name = text.split("[", 1)[0]
				col_name = text.split("[", 1)[1].split("]", 1)[0]
				for ws in self.workbook.sheets:
					try:
						los = ws.api.ListObjects
						if not los or los.Count == 0:
							continue
						for i in range(1, los.Count + 1):
							lo = los.Item(i)
							if getattr(lo, "Name", None) == table_name:
								try:
									col = lo.ListColumns.Item(col_name)
									dbr = getattr(col, "DataBodyRange", None)
									if dbr is not None:
										addr = getattr(dbr, "Address", None)
										if addr:
											vals = ws.range(addr).value
											return self._flatten_to_list(vals)
								except Exception:
									continue
					except Exception:
						continue
			return None
		except Exception:
			return None

	def _resolve_validation_list_items(self, cell, formula1: Optional[str]) -> Optional[List[Any]]:
		if not formula1:
			return None
		try:
			f = str(formula1).strip()
			# Remove leading '=' if present
			if f.startswith("="):
				f = f[1:]
			# Remove wrapping quotes for literal lists
			if f.startswith('"') and f.endswith('"') and len(f) >= 2:
				f = f[1:-1]
			# If comma-separated literal list like "A,B,C"
			if "," in f and "!" not in f and ":" not in f and "[" not in f:
				return [item.strip() for item in f.split(",")]
			# Structured reference to a table column
			if "[" in f and "]" in f:
				resolved = self._try_resolve_table_column(f)
				if resolved is not None:
					return resolved
			# Range reference possibly with sheet, e.g., Sheet1!$A$1:$A$10 or 'Sheet Name'!$A$1:$A$10
			if "!" in f or ":" in f:
				sheet_name = None
				addr = f
				if "!" in f:
					sheet_name, addr = f.split("!", 1)
					sheet_name = sheet_name.strip().strip("'")
				target_sheet = self.worksheet if sheet_name is None else self.workbook.sheets[sheet_name]
				vals = self._values_from_range_on_sheet(target_sheet, addr)
				if vals is not None:
					return vals
			# Named range resolution
			named_vals = self._try_resolve_named_range(f)
			if named_vals is not None:
				return named_vals
			# Last resort: evaluate expression in Excel context
			res = self._evaluate_in_excel(f, self.worksheet)
			try:
				addr = getattr(res, "Address", None)
			except Exception:
				addr = None
			if addr:
				vals = self._values_from_range_on_sheet(self.worksheet, addr)
				if vals is not None:
					return vals
		except Exception:
			return None
		return None
	
	def _get_cell_basic_format(self, cell) -> Dict[str, Any]:
		# Expand existing formatting info
		format_info: Dict[str, Any] = {}
		try:
			format_info["number_format"] = cell.number_format
		except Exception:
			format_info["number_format"] = None
		try:
			format_info["font_name"] = cell.font.name
			format_info["font_size"] = cell.font.size
			format_info["font_bold"] = cell.font.bold
			format_info["font_italic"] = cell.font.italic
		except Exception:
			pass
		try:
			format_info["font_color_rgb"] = self._extract_font_color_rgb(cell)
		except Exception:
			pass
		try:
			format_info["fill_color"] = self._get_cell_fill_color(cell)
		except Exception:
			pass
		try:
			format_info["horizontal_alignment"] = getattr(cell.api, "HorizontalAlignment", None)
			format_info["vertical_alignment"] = getattr(cell.api, "VerticalAlignment", None)
		except Exception:
			pass
		try:
			format_info["locked"] = bool(getattr(cell.api, "Locked", False))
			format_info["formula_hidden"] = bool(getattr(cell.api, "FormulaHidden", False))
		except Exception:
			pass
		try:
			merge_cells = bool(getattr(cell.api, "MergeCells", False))
			format_info["merged"] = merge_cells
			if merge_cells:
				format_info["merge_area"] = getattr(cell.api.MergeArea, "Address", None)
		except Exception:
			pass
		return format_info
	
	def _extract_font_color_rgb(self, cell) -> Optional[Dict[str, int]]:
		try:
			color_val = cell.api.Font.Color
			if color_val is None:
				return None
			b = (int(color_val) >> 16) & 255
			g = (int(color_val) >> 8) & 255
			r = int(color_val) & 255
			return {"r": r, "g": g, "b": b}
		except Exception:
			return None
	
	def _extract_cell_full_details(self, cell) -> Dict[str, Any]:
		# Comprehensive per-cell record
		value = None
		formula = None
		try:
			value = cell.value
		except Exception:
			value = None
		try:
			formula = cell.formula
		except Exception:
			formula = None
		details: Dict[str, Any] = {
			"address": cell.address,
			"row": cell.row,
			"column": cell.column,
			"column_letter": self._column_to_letter(cell.column),
			"value": value,
			"formula": formula if formula else None,
			"display_text": self._get_cell_display_text(cell),
			"format": self._get_cell_basic_format(cell),
			"hyperlink": self._get_cell_hyperlink(cell),
			"note": self._get_cell_note(cell),
			"data_validation": self._get_cell_validation(cell)
		}
		return details
	
	def extract_sheet_full_details(self, sheet_name: Optional[str] = None) -> Dict[str, Any]:
		"""Extract full details for a single worksheet (all used cells)."""
		try:
			# Select worksheet
			sheet_info = self.get_worksheet_info(sheet_name)
			# Safe used range retrieval
			used_address = None
			rows_count = 0
			cols_count = 0
			try:
				used_range = self.worksheet.used_range
				used_address = used_range.address
				rows_count = used_range.rows.count
				cols_count = used_range.columns.count
			except Exception:
				try:
					used_range_api = self.worksheet.api.UsedRange
					used_address = getattr(used_range_api, "Address", None)
					rows_count = getattr(getattr(used_range_api, "Rows", None), "Count", 0) or 0
					cols_count = getattr(getattr(used_range_api, "Columns", None), "Count", 0) or 0
				except Exception:
					used_address = None
					rows_count = 0
					cols_count = 0
			cells: List[Dict[str, Any]] = []
			if rows_count and cols_count:
				for row in range(1, rows_count + 1):
					for col in range(1, cols_count + 1):
						try:
							cell = self.worksheet.cells(row, col)
							cells.append(self._extract_cell_full_details(cell))
						except Exception:
							# Continue on per-cell failures
							continue
			# Tables (ListObjects)
			tables: List[Dict[str, Any]] = []
			try:
				list_objects = self.worksheet.api.ListObjects
				if list_objects and list_objects.Count > 0:
					for i in range(1, list_objects.Count + 1):
						lo = list_objects.Item(i)
						try:
							columns = []
							try:
								if getattr(lo, "ListColumns", None) is not None:
									for j in range(1, lo.ListColumns.Count + 1):
										col_obj = lo.ListColumns.Item(j)
										columns.append(getattr(col_obj, "Name", None))
							except Exception:
								columns = []
							tables.append({
								"name": getattr(lo, "Name", None),
								"range": getattr(getattr(lo, "Range", None), "Address", None),
								"data_body_range": getattr(getattr(lo, "DataBodyRange", None), "Address", None),
								"header_row_range": getattr(getattr(lo, "HeaderRowRange", None), "Address", None),
								"totals_row_range": getattr(getattr(lo, "TotalsRowRange", None), "Address", None),
								"show_totals": bool(getattr(lo, "ShowTotals", False)),
								"columns": columns
							})
						except Exception:
							continue
			except Exception:
				pass
			# Sheet-level properties
			try:
				visible = getattr(self.worksheet.api, "Visible", None)
				protect_contents = getattr(self.worksheet.api, "ProtectContents", None)
				protect_drawing = getattr(self.worksheet.api, "ProtectDrawing", None)
				protect_scenarios = getattr(self.worksheet.api, "ProtectScenarios", None)
			except Exception:
				visible = None
				protect_contents = None
				protect_drawing = None
				protect_scenarios = None
			# Override used_range in sheet_info with safe one if available
			if used_address:
				sheet_info["used_range"] = used_address
				sheet_info["rows"] = rows_count
				sheet_info["columns"] = cols_count
				sheet_info["total_cells"] = rows_count * cols_count
			sheet_props: Dict[str, Any] = {
				"name": sheet_info.get("sheet_name"),
				"used_range": sheet_info.get("used_range"),
				"rows": sheet_info.get("rows"),
				"columns": sheet_info.get("columns"),
				"visible": visible,
				"protect_contents": protect_contents,
				"protect_drawing": protect_drawing,
				"protect_scenarios": protect_scenarios
			}
			return {
				"sheet": sheet_props,
				"cells": cells,
				"tables": tables
			}
		except Exception as e:
			print(f"Error extracting full details for sheet: {e}")
			# Return at least basic sheet metadata on error
			safe_name = None
			try:
				safe_name = self.worksheet.name
			except Exception:
				safe_name = sheet_name
			return {
				"sheet": {
					"name": safe_name,
					"error": str(e)
				},
				"cells": [],
				"tables": []
			}

	def extract_workbook_full_details(self) -> Dict[str, Any]:
		"""Extract full details across all sheets, plus named ranges."""
		try:
			all_sheets: List[Dict[str, Any]] = []
			for ws in self.workbook.sheets:
				self.worksheet = ws
				try:
					all_sheets.append(self.extract_sheet_full_details(ws.name))
				except Exception as e:
					print(f"Error extracting sheet '{ws.name}': {e}")
					all_sheets.append({
						"sheet": {"name": ws.name, "error": str(e)},
						"cells": [],
						"tables": []
					})
			# Named ranges
			names: List[Dict[str, Any]] = []
			try:
				for nm in self.workbook.names:
					try:
						refers_to = getattr(nm, "refers_to", None)
					except Exception:
						refers_to = None
					names.append({
						"name": getattr(nm, "name", None),
						"refers_to": refers_to
					})
			except Exception:
				pass
			return {
				"file_path": str(self.excel_file_path),
				"extraction_timestamp": datetime.now().isoformat(),
				"workbook": {
					"sheet_count": len(self.workbook.sheets),
					"sheets": all_sheets,
					"names": names
				}
			}
		except Exception as e:
			print(f"Error extracting workbook details: {e}")
			return {}
	
	def extract_formula_dependencies(self, cell_address: str) -> Dict[str, Any]:
		"""
		Extract dependencies for a specific formula cell
		
		Args:
			cell_address (str): Cell address (e.g., "A1")
			
		Returns:
			Dictionary containing formula dependencies
		"""
		try:
			cell = self.worksheet.range(cell_address)
			formula = cell.formula
			
			if not formula or not formula.startswith('='):
				return {"error": "Cell does not contain a formula"}
			
			# Extract cell references from formula
			dependencies = self._parse_formula_dependencies(formula)
			
			# Get values of dependent cells
			dependent_values = {}
			for dep in dependencies:
				try:
					dep_cell = self.worksheet.range(dep)
					dependent_values[dep] = {
						"value": dep_cell.value,
						"formula": dep_cell.formula if dep_cell.formula.startswith('=') else None,
						"address": dep
					}
				except:
					dependent_values[dep] = {"error": "Could not access cell"}
			
			return {
				"cell_address": cell_address,
				"formula": formula,
				"dependencies": dependencies,
				"dependent_values": dependent_values,
				"calculated_value": cell.value
			}
		except Exception as e:
			return {"error": f"Error extracting dependencies: {e}"}
	
	def _parse_formula_dependencies(self, formula: str) -> List[str]:
		"""
		Parse formula to extract cell references
		
		Args:
			formula (str): Excel formula string
			
		Returns:
			List of cell references
		"""
		import re
		
		# Pattern to match Excel cell references (A1, B2, etc.)
		cell_pattern = r'[A-Z]+\d+'
		matches = re.findall(cell_pattern, formula.upper())
		
		# Remove duplicates and sort
		unique_refs = sorted(list(set(matches)))
		
		return unique_refs
	
	def export_to_json(self, data: Dict[str, Any], output_file: str) -> bool:
		"""
		Export extracted data to JSON file
		
		Args:
			data (Dict): Data to export
			output_file (str): Output file path
			
		Returns:
			bool: True if successful, False otherwise
		"""
		try:
			with open(output_file, 'w', encoding='utf-8') as f:
				json.dump(data, f, indent=2, ensure_ascii=False, default=str)
			print(f"Data exported to: {output_file}")
			return True
		except Exception as e:
			print(f"Error exporting to JSON: {e}")
			return False
	
	def export_to_text(self, data: Dict[str, Any], output_file: str) -> bool:
		"""
		Export extracted data to text file
		
		Args:
			data (Dict): Data to export
			output_file (str): Output file path
			
		Returns:
			bool: True if successful, False otherwise
		"""
		try:
			with open(output_file, 'w', encoding='utf-8') as f:
				f.write("EXCEL FORMULA EXTRACTION REPORT\n")
				f.write("=" * 50 + "\n\n")
				
				f.write(f"File: {data.get('file_path', 'Unknown')}\n")
				f.write(f"Extracted: {data.get('extraction_timestamp', 'Unknown')}\n")
				f.write(f"Worksheet: {data.get('worksheet_info', {}).get('sheet_name', 'Unknown')}\n")
				f.write(f"Total Formulas: {data.get('total_formulas_found', 0)}\n\n")
				
				f.write("FORMULAS:\n")
				f.write("-" * 20 + "\n")
				
				for formula_info in data.get('formulas', []):
					f.write(f"Cell: {formula_info.get('address', 'Unknown')}\n")
					if formula_info.get('formula'):
						f.write(f"Formula: {formula_info['formula']}\n")
					f.write(f"Value: {formula_info.get('value', 'N/A')}\n")
					f.write(f"Row: {formula_info.get('row', 'N/A')}, Column: {formula_info.get('column_letter', 'N/A')}\n")
					f.write("-" * 10 + "\n")
			
			print(f"Data exported to: {output_file}")
			return True
		except Exception as e:
			print(f"Error exporting to text: {e}")
			return False 