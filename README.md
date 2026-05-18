# open-skills

Open-source (Apache 2.0) agent skills for document creation and editing.

## Skills

### `pptx-creation-editing`

Create, edit, and validate PowerPoint (`.pptx`) files.

**Features**
- **Slides** — add, remove, reorder (lxml `<p:sldIdLst>` manipulation), bulk creation from JSON spec
- **Shapes & images** — text boxes, pictures with alt text, tables; image replacement by swapping OPC relationship targets
- **Speaker notes** — read/write via python-pptx API; direct lxml for multi-paragraph/formatted notes
- **Validation** — ZIP integrity, required OPC parts, XML well-formedness, relationship targets

```bash
pip install python-pptx lxml pillow

# Create a demo deck
python pptx-skill/scripts/create_pptx.py --out demo.pptx

# Edit: set notes on slide 0
python pptx-skill/scripts/edit_pptx.py set-notes deck.pptx 0 "My speaker note"

# Validate
python pptx-skill/scripts/validate_pptx.py --file deck.pptx

# Run evals (requires POI test corpus)
POI_PATH=/path/to/poi/test-data/slideshow python pptx-skill/evals/eval_runner.py
```

| Test suite | Result |
|------------|--------|
| Eval tests | 5 / 5 passed |
| POI corpus assertions | 120 / 120 passed |
| Fixtures tested | 15 |

See [`pptx-skill/SKILL.md`](pptx-skill/SKILL.md) for full API reference.

---

### `docx-creation-editing`

Create and edit Word (`.docx`) documents.

See [`docx-skill/`](docx-skill/) for source and [`docx-creation-editing.skill`](docx-creation-editing.skill) for the packaged artifact.

---

## Structure

```
open-skills/
├── pptx-skill/           # pptx-creation-editing source
│   ├── SKILL.md
│   ├── LICENSE.txt
│   ├── references/        # API and OOXML reference docs
│   ├── scripts/           # create_pptx.py, edit_pptx.py, validate_pptx.py, benchmark.py
│   └── evals/             # evals.json, eval_runner.py, corpus_test.py, eval_viewer.html
├── pptx-creation-editing.skill   # packaged ZIP artifact
├── docx-skill/
└── docx-creation-editing.skill
```

## License

Apache 2.0 — see [LICENSE](LICENSE).
