# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A collection of open-source (Apache 2.0) agent skills for Office document creation and editing. Each skill is a self-contained directory that gets packaged as a `.skill` ZIP file.

Currently two skills exist:
- **`docx-skill/`** — Word document creation/editing (python-docx + lxml)
- **`pptx-skill/`** — PowerPoint creation/editing (python-pptx + lxml)

## Running tests

Both skills use the Apache POI test corpus. Set `POI_PATH` to the corpus directory:

```bash
export POI_PATH=/home/user/poi/test-data/slideshow   # pptx
export POI_PATH=/home/user/poi/test-data/document    # docx
```

**pptx-skill — 5 targeted eval tests:**
```bash
cd pptx-skill
POI_PATH=... python evals/eval_runner.py
```

**pptx-skill — 120 corpus assertions across 15 fixtures:**
```bash
POI_PATH=... python evals/corpus_test.py
```

**docx-skill — comprehensive test suite:**
```bash
cd docx-skill
python scripts/comprehensive_test.py
```

**Run a single pptx eval by calling the function directly:**
```python
import sys; sys.path.insert(0, 'pptx-skill')
from evals.eval_runner import test_reorder_slides
r = test_reorder_slides(); print(r.passed, r.failures)
```

**Benchmarks** (both skills):
```bash
POI_PATH=... python pptx-skill/scripts/benchmark.py --quick
POI_PATH=... python docx-skill/scripts/benchmark.py
```

## Packaging a skill

```bash
python3 -c "
import zipfile
from pathlib import Path
skill_dir = Path('pptx-skill')  # or docx-skill
with zipfile.ZipFile('pptx-creation-editing.skill', 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in sorted(skill_dir.rglob('*')):
        if f.is_file() and '__pycache__' not in str(f):
            zf.write(f, f.relative_to(skill_dir))
"
```

## Architecture

### Skill layout (same pattern for both skills)

```
{skill}-skill/
├── SKILL.md            YAML frontmatter (name, description, license, version, deps) + feature tables
├── LICENSE.txt         Apache 2.0
├── references/
│   ├── python-{lib}.md  Summarized API reference from official docs (not verbatim copies)
│   └── ooxml-{fmt}.md   ECMA-376 XML structure for shapes, images, notes, etc.
├── scripts/
│   ├── create_*.py     New document/presentation creation; returns objects, never hardcodes paths
│   ├── edit_*.py       Open + modify existing files; slide/page management, notes, image replace
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

### Where python-pptx/python-docx ends and lxml begins

Each `edit_*.py` script documents this boundary. Key lxml-only operations:

| Feature | Why lxml needed |
|---------|----------------|
| Slide reorder/delete | No python-pptx API; manipulate `<p:sldIdLst>` children |
| Image alt text (pptx) | `<p:cNvPr descr="…">` — no python-pptx setter |
| Image replacement | Update `r:embed` on `<a:blip>` + `slide.part.get_or_add_image_part()` |
| Multi-paragraph notes | Remove/insert `<a:p>` children in the notes `txBody` directly |
| Anchored (floating) images in docx | python-docx only exposes inline shapes |
| Comment deletion (docx) | Must remove XML triplet + `w:comment` entry via lxml |
| Track changes (docx) | Construct `w:ins`/`w:del` wrappers; no python-docx API |

### Adding a new skill

Follow the established layout exactly. Build exclusively from:
1. The library's official documentation
2. ECMA-376 Part 1 specification (for XML element details)
3. Apache POI corpus at `POI_PATH` — **fixtures only**, never read the Java source

Document every lxml decision with an ECMA-376 section reference in the docstring.

## Dependencies

```bash
pip install python-pptx python-docx lxml pillow
```

No `pyproject.toml` or `requirements.txt` yet — dependencies are listed in each `SKILL.md` frontmatter.
