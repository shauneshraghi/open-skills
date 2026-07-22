# Third-party notices

The skills in this repository are licensed under Apache-2.0 (see [LICENSE](./LICENSE)).
They bundle reference documentation from the sources below, which is
redistributed under its own license.

## Microsoft Office VBA documentation

Files under the following paths are authored by Microsoft and redistributed
here unmodified except where noted:

| Path | Files |
|---|---|
| `docx-creation-editing/references/vba-docs/` | Word VBA reference |
| `pptx-creation-editing/references/vba-docs/` | PowerPoint VBA reference |
| `xlsx-creation-editing/references/vba-docs/` | Excel VBA reference |
| `*/references/{winword,powerpnt,excel}-cli.md` | Office command-line switches |

- **Source:** [MicrosoftDocs/VBA-Docs](https://github.com/MicrosoftDocs/VBA-Docs)
- **Published at:** <https://learn.microsoft.com/office/vba/api/overview/>
- **Documentation license:** [CC-BY-4.0](https://creativecommons.org/licenses/by/4.0/)
- **Code samples license:** MIT

© Microsoft Corporation. Used under CC-BY-4.0. Files retain their original
`title` / `ms.date` front matter so each article can be traced to its upstream
source. Some articles have been trimmed to the subset relevant to these skills;
no technical content has been rewritten.

## Excel function reference

`xlsx-creation-editing/references/functions/*.md` (13 files) are condensed from
Microsoft's [Excel functions (by category)](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)
on support.microsoft.com. They list function names grouped by category with
brief descriptions; each file cites the source URL in its header. This is
support-site content under the [Microsoft Terms of Use](https://www.microsoft.com/legal/terms-of-use),
not the CC-BY-4.0 docs repos above.

## Summarized API references (no third-party text)

These are written for this repository and paraphrase upstream documentation
rather than reproduce it; they are covered by this repo's Apache-2.0 license:

- `*/references/python-docx.md`, `python-pptx.md`, `openpyxl.md` — summarized
  from each library's source and docs, with the upstream URL cited.
- `*/references/ooxml-*.md` — patterns extracted from documents authored for
  this repo.

## Apache POI (test corpus, not redistributed)

`*/scripts/comprehensive_test.py` clones [apache/poi](https://github.com/apache/poi)
at run time to use its `test-data/` corpus as a regression baseline. No Apache
POI content is included in this repository. Apache POI is licensed under
Apache-2.0.
