#!/usr/bin/env python3
from pathlib import Path
from typing import Any, Dict, List

from .extractor import ExcelFormulaExtractor  # for package coherence

# Re-export conversion helpers from the script if needed. For now, keep this minimal.

def convert_full_details_json(input_json: Path, output_dir: Path, make_ndjson: bool) -> None:
	from pathlib import Path as _Path
	from typing import Any as _Any, Dict as _Dict, List as _List
	from ._convert_impl import convert as _convert
	_convert(_Path(str(input_json)), _Path(str(output_dir)), bool(make_ndjson)) 