#!/usr/bin/env python3
import argparse
from pathlib import Path
from ._convert_impl import convert


def main():
	parser = argparse.ArgumentParser(description="Convert Excel full-details JSON into per-sheet artifacts for easy browsing.")
	parser.add_argument("input", help="Path to *_full_details.json")
	parser.add_argument("--out", default="exports", help="Output directory (default: exports)")
	parser.add_argument("--ndjson", action="store_true", help="Also write a combined cells.ndjson for grepping")
	args = parser.parse_args()

	convert(Path(args.input), Path(args.out), bool(args.ndjson))


if __name__ == "__main__":
	main() 