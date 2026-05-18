# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A collection of open-source (Apache 2.0) agent skills for Office document creation and editing. Each skill is a self-contained directory that gets packaged as a `.skill` ZIP file.

Three skills exist:
- **`docx-creation-editing/`** — Word document creation/editing (python-docx + lxml)
- **`pptx-creation-editing/`** — PowerPoint creation/editing (python-pptx + lxml)
- **`xlsx-creation-editing/`** — Excel spreadsheet creation/editing (openpyxl + lxml)

## Running tests

All skills use the Apache POI test corpus. Set `POI_PATH` to the corpus directory:

```bash
export POI_PATH=/home/user/poi/test-data/document      # docx
export POI_PATH=/home/user/poi/test-data/slideshow     # pptx
export POI_PATH=/home/user/poi/test-data/spreadsheet   # xlsx
```

**docx-creation-editing — comprehensive test suite:**
```bash
cd docx-creation-editing
python scripts/comprehensive_test.py
```

**pptx-creation-editing — 5 targeted eval tests:**
```bash
cd pptx-creation-editing
POI_PATH=... python evals/eval_runner.py
```

**pptx-creation-editing — 120 corpus assertions across 15 fixtures:**
```bash
POI_PATH=... python evals/corpus_test.py
```

**Run a single pptx eval by calling the function directly:**
```python
import sys; sys.path.insert(0, 'pptx-creation-editing')
from evals.eval_runner import test_reorder_slides
r = test_reorder_slides(); print(r.passed, r.failures)
```

**xlsx-creation-editing — 5 targeted eval tests:**
```bash
cd xlsx-creation-editing
POI_PATH=/home/user/poi/test-data/spreadsheet python evals/eval_runner.py
```

**xlsx-creation-editing — 120 corpus assertions across 15 fixtures:**
```bash
POI_PATH=/home/user/poi/test-data/spreadsheet python evals/corpus_test.py
```

**Run a single xlsx eval by calling the function directly:**
```python
import sys; sys.path.insert(0, 'xlsx-creation-editing')
from evals.eval_runner import test_embed_image_with_alt_text
r = test_embed_image_with_alt_text(); print(r.passed, r.failures)
```

**Benchmarks** (all skills):
```bash
POI_PATH=... python docx-creation-editing/scripts/benchmark.py
POI_PATH=... python pptx-creation-editing/scripts/benchmark.py --quick
POI_PATH=/home/user/poi/test-data/spreadsheet python xlsx-creation-editing/scripts/benchmark.py --quick
```

## Packaging a skill

```bash
python3 -c "
import zipfile
from pathlib import Path
skill_dir = Path('xlsx-creation-editing')  # or docx-creation-editing / pptx-creation-editing
with zipfile.ZipFile(skill_dir.name + '.skill', 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in sorted(skill_dir.rglob('*')):
        if f.is_file() and '__pycache__' not in str(f):
            zf.write(f, f.relative_to(skill_dir))
"
```

## Architecture

### Skill layout (same pattern for all skills)

```
{skill}-creation-editing/
├── SKILL.md            YAML frontmatter (name, description, license, version, deps) + feature tables
├── LICENSE.txt         Apache 2.0
├── references/
│   ├── {lib}.md         Summarized API reference from official docs (not verbatim copies)
│   └── ooxml-{fmt}.md   ECMA-376 XML structure for shapes, images, notes, etc.
├── scripts/
│   ├── create_*.py     New document creation; returns objects, never hardcodes paths
│   ├── edit_*.py       Open + modify existing files; reorder, image replace, read values
│   ├── validate_*.py   ZIP structure, required OPC parts, XML well-formedness, broken rels
│   └── benchmark.py    Wall-clock timing against POI corpus; uses POI_PATH env var
└── evals/
    ├── evals.json       Human-readable test spec (assertions as strings, not runnable)
    ├── eval_runner.py   Executable: runs 5 targeted tests, sys.exit(0/1)
    ├── corpus_test.py   Executable: runs 8 checks × 15 POI fixtures, sys.exit(0/1)
    └── eval_viewer.html Static HTML dashboard showing results
```

### Script module conventions

- All scripts are **importable modules** (functions at top level) and also have a `main()` / CLI entry point.
- Scripts import from sibling scripts: `from scripts.create_pptx import set_notes` — so run them from the skill root, or add it to `sys.path`.
- All file paths use `pathlib.Path`. Temp output uses `tempfile.mkdtemp()`. Never hardcode `/tmp/` or `/home/user/`.
- `POI_PATH` env var controls the corpus location; scripts fall back to a relative default (`../../poi/test-data/...`).

### Where the primary library ends and lxml begins

Each `edit_*.py` script documents this boundary with ECMA-376 section references. Key lxml-only operations:

| Feature | Why lxml needed | ECMA-376 ref |
|---------|----------------|-------------|
| Slide reorder/delete | No python-pptx API; manipulate `<p:sldIdLst>` children | §19.2.1.38 |
| Image alt text (pptx) | `<p:cNvPr descr="…">` — no python-pptx setter | §20.1.2.2.8 |
| Image replacement (pptx) | Update `r:embed` on `<a:blip>` + `get_or_add_image_part()` | §20.1.8.11 |
| Multi-paragraph notes | Remove/insert `<a:p>` children in the notes `txBody` | §19.3.1.43 |
| Anchored images (docx) | python-docx only exposes inline shapes | §20.4 |
| Comment deletion (docx) | Must remove XML triplet + `w:comment` entry | §17.13.4 |
| Track changes (docx) | Construct `w:ins`/`w:del` wrappers; no python-docx API | §17.13.5 |
| Image alt text (xlsx) | `<xdr:cNvPr descr="…">` — no openpyxl setter; post-save ZIP patch | §20.5.2.8 |
| Sheet reorder fallback (xlsx) | Manipulate `<workbook><sheets>` child order if `move_sheet()` absent | §18.2.19 |

### Adding a new skill

Follow the established layout exactly. Build exclusively from:
1. The library's official documentation
2. ECMA-376 Part 1 specification (for XML element details)
3. Apache POI corpus at `POI_PATH` — **fixtures only**, never read the Java source

Document every lxml decision with an ECMA-376 section reference in the docstring.

## Dependencies

```bash
pip install python-pptx python-docx openpyxl lxml pillow
```

No `pyproject.toml` or `requirements.txt` yet — dependencies are listed in each `SKILL.md` frontmatter.
