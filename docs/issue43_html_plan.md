# Issue 43 Full-Coverage Fixture and HTML Rendering Plan

This plan tracks GitHub issue #43: create a broad DOCX feature fixture and render DOCX metadata to HTML.

## Current Implementation

- `examples/full_coverage.md`
  - Human-readable Markdown source that enumerates all requested coverage areas.
  - Covers headings, alignment samples, text formatting, fonts, highlighting, lists, quotes, code, horizontal rules, links, images, tables, formulas, notes, task lists, and advanced DOCX-only items.
- `examples/full_coverage.docx`
  - Generated from the Markdown source with the project CLI, then augmented with Word/OpenXML-only features.
  - Includes page headers/footers, first/even/odd header/footer variants, page break, multiple sections, landscape section, columns, complex tables, merged cells, nested table, image-in-table, bookmark, comment, footnote, endnote, and OMML math.
- `md2word.html.DocxHtmlRenderer`
  - Renders from extracted metadata, not by directly reading the DOCX package.
  - Inlines image resources as data URIs.
  - Emits MathJax configuration and wraps formula text for MathJax rendering.
  - Renders paragraph/run styling, headings, tables with colspan/rowspan, lists, VML text boxes, headers/footers, notes, comments, and bookmark anchors.
- CLI
  - `--to-html` renders either a `.docx` or an existing metadata directory/JSON to standalone HTML.

## Verification Commands

```bash
python -m pytest tests/test_docx_html.py tests/test_issue43_fixtures.py -q
python -m pytest -q
python -m md2word.cli.main examples/full_coverage.docx --to-html -o output/full_coverage.html
python -m md2word.cli.main examples/full_coverage.docx --extract-meta -o output/full_coverage_meta
python -m md2word.cli.main output/full_coverage_meta --to-html -o output/full_coverage_from_meta.html
```

## Current Boundaries

- The HTML renderer is metadata-driven and does not attempt to call Word or parse the DOCX directly after extraction.
- Pixel-perfect browser pagination is not guaranteed because browsers do not implement Word layout exactly.
- The renderer preserves the semantic and visual metadata available in `document.json`; unsupported OpenXML details degrade to readable HTML rather than failing the conversion.

## Completion Criteria

- Full-coverage Markdown and DOCX fixtures are tracked in the repository.
- `--to-html` works from both DOCX input and extracted metadata input.
- Tests cover HTML rendering of key metadata features.
- Full test suite passes.
- PR is merged and issue #43 is closed with verification results.
