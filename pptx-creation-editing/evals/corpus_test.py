"""Compatibility wrapper for the PPTX comprehensive corpus workflow."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPT_PATH = ROOT / "scripts" / "comprehensive_test.py"
SPEC = importlib.util.spec_from_file_location("pptx_comprehensive_test", SCRIPT_PATH)
if SPEC is None or SPEC.loader is None:
    raise ImportError(f"Unable to load comprehensive_test module from {SCRIPT_PATH}")
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)
main = MODULE.main


if __name__ == "__main__":
    sys.exit(main())
