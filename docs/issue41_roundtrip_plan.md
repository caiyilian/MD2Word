# Issue 41 DOCX Metadata Roundtrip Plan

This plan tracks the staged work for GitHub issue #41: extract a DOCX into human-readable structured metadata plus resources, then restore the DOCX from that metadata.

## Current Baseline

Implemented in the current worktree:

- `md2word.meta.DocxMetaExtractor`
  - Extracts every DOCX ZIP entry.
  - Stores XML parts as structured JSON trees, not raw XML strings.
  - Stores binary resources under `resources/<package-path>`.
  - Stores a ZIP container exact payload cache for unchanged-entry byte-identical restore.
  - Adds a human-readable `document` semantic index.
- `md2word.meta.DocxMetaRenderer`
  - Restores a DOCX from the structured JSON tree and resources.
  - Reuses the exact ZIP payload cache when structured XML/resource content is unchanged.
  - Falls back to rebuilding a valid DOCX ZIP when metadata/resources have been edited.
  - Preserves ZIP entry metadata needed for byte-identical output on generated test documents.
- `verify_docx_metadata_roundtrip`
  - Extracts, restores, and compares SHA-256 hashes.
- CLI
  - `--extract-meta`
  - `--restore-meta`
  - `--roundtrip-meta`

## Metadata Shape

Top-level `document.json`:

- `format`, `version`
- `source`: source filename, size, SHA-256
- `package`: ZIP-level metadata
- `document`: human-readable semantic index
- `entries`: reconstructable package entries

The `document` semantic index currently includes:

- `metadata`: `docProps/core.xml`, `docProps/app.xml`, `docProps/custom.xml`
- `settings`: `word/settings.xml`
- `styles`: `word/styles.xml`
- `numbering`: `word/numbering.xml`
- `fonts`: `word/fontTable.xml`
- `theme`: `word/theme/*.xml`
- `body`: paragraph, run, table, image, VML shape/textbox, math, hyperlink, note/comment references
- `body`: field codes, structured document tags, tracked insert/delete/move revisions
- `charts`, `diagrams`, `embeddings`, `active_x`, `vba_project`, `people`, `glossary`, `custom_xml`, `mail_merge`
- `sections`: section properties
- `headers_footers`
- `footnotes`, `endnotes`, `comments`
- `resources`, `relationships`

## Stage Breakdown

### Stage 1: Package-Level Closed Loop

Status: implemented.

Acceptance:

- Extracted metadata contains structured XML trees and resource files.
- Restored DOCX is openable by `python-docx`.
- Generated DOCX fixtures restore byte-identically.
- `python -m pytest tests/test_docx_meta.py -q` passes.

### Stage 2: High-Priority Semantic Body Coverage

Status: implemented for common body elements.

Covered:

- Run properties: font names, sizes, color, highlight, underline, strike, vertical align, spacing, scale, effects, borders, hidden text, language.
- Paragraph properties: style, alignment, numbering, indent, spacing, shading, borders, tabs, pagination controls, text direction.
- Tables: table properties, grid, row properties, cell widths, gridSpan, vMerge, borders, shading, margins, vertical alignment, nested body content.
- Drawings: inline/anchor type, relationship/resource, extent, wrap, position, crop, transform, rotation/flip, line/border, effects, and picture hyperlinks.
- Sections: page size, margins, columns, page numbering, header/footer references.

Remaining:

- Deeper rendering semantics beyond the current semantic indexes, especially grouped DrawingML, WordArt, ink, chart embedded workbook data, SmartArt rendering behavior, OLE object details, and macro signatures.

### Stage 3: Support-Part Semantic Coverage

Status: implemented for common support parts.

Covered:

- Styles: IDs, names, type, inheritance, links, default flags, run/paragraph/table properties.
- Numbering: abstract numbering definitions, levels, formats, text templates, indentation, concrete numbering instances, overrides.
- Settings: document protection, view/zoom, default tab stop, revision tracking flags, math settings.
- Font table: names, charset, family, pitch, Panose, embedded font references.
- Theme: color scheme and major/minor font scheme.
- Document properties: core, app, custom properties.

Remaining:

- Full theme effect/style matrix.
- Full compatibility settings and document variables.
- Linked custom properties and advanced docProps typing.

### Stage 4: Real-Document Verification

Status: implemented for `examples/test.docx`.

Acceptance:

- `examples/test.docx` is tracked as the representative real-document fixture.
- `python -m md2word.cli.main examples/test.docx --roundtrip-meta -o output/test_roundtrip`
  - Source SHA-256: `04534fe5936877d039f8de9bbbcbfb13edb6df4a9ee3df0adf9b0c91f38eb284`
  - Restored SHA-256: `04534fe5936877d039f8de9bbbcbfb13edb6df4a9ee3df0adf9b0c91f38eb284`
  - Result: byte-identical.
- `python docx2img.py --compare examples/test.docx output/test_roundtrip/test.restored.docx -o output/test_compare`
  - Result: all 30 pages match pixel-perfect.

### Stage 5: Advanced OpenXML Feature Indexes

Status: partially implemented for VML shapes/textboxes, fields, content controls, tracked text revisions, and advanced package part inventories.

Sub-stages:

- Fields and TOC/PAGE/REF/SEQ field code extraction: implemented for paragraph field sequences and run-level `fldChar` / `instrText`.
- Content controls (`w:sdt`) with tags, aliases, locks, bindings, and control type: implemented for block and inline SDT metadata.
- Revisions (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`): implemented for revision containers and inserted/deleted text metadata.
- Revision property changes (`w:rPrChange`, `w:pPrChange`, `w:tblPrChange`): pending.
- VML/DrawingML shapes, text boxes, WordArt, ink, canvas, groups: VML shape/textbox/fill/stroke/text parsing and DrawingML picture crop/transform/line/effect/hyperlink parsing implemented; deeper WordArt, ink, canvas, and grouped DrawingML semantics pending.
- Charts and embedded chart data relationships: chart part inventory implemented with chart type, title text, series count, chart groups, series titles, category/value references and cached points, axes, legend, and relationships.
- SmartArt diagram data: implemented for data model points/connections, layout nodes, quick style labels, and color labels across data/layout/style/color parts; deeper rendering behavior pending.
- OLE embeddings and package attachments: resource inventory implemented.
- Macro project detection and signature metadata: `vbaProject.bin` detection implemented; signature metadata pending.
- Mail merge settings and data-source relationships: settings-level mail merge summary implemented; external data source interpretation pending.
- People and authorship metadata: `word/people.xml` summary implemented.
- ActiveX controls: XML and binary part inventory implemented.
- Glossary/building blocks: glossary document part summary implemented.
- Custom XML parts: custom XML inventory implemented.

### Stage 6: Editing from Metadata

Status: pending.

The current renderer restores from the reconstructable structured XML tree. A future stage can support intentional edits to the semantic `document` index and then apply them back into package entries.

Acceptance:

- Modify semantic metadata for a constrained feature, such as paragraph text or run color.
- Render the changed metadata into a valid DOCX.
- Verify changed output visually and structurally.

## Verification Commands

```bash
python -m pytest -q
python -m md2word.cli.main input.docx --extract-meta -o output/input_meta
python -m md2word.cli.main output/input_meta --restore-meta -o output/restored.docx
python -m md2word.cli.main input.docx --roundtrip-meta -o output/input_roundtrip
python docx2img.py --compare input.docx output/input_roundtrip/input.restored.docx -o output/input_compare
```

## Completion Criteria for Issue 41

The issue should only be closed when:

- The metadata format covers every required OpenXML part listed in `docx_features_full_coverage.md`, either semantically or as structured package metadata.
- Roundtrip restore is byte-identical for controlled fixtures.
- Visual comparison passes for representative real DOCX documents.
- Tests cover each priority class from the full coverage checklist.
- A PR has been opened, merged, and the GitHub issue has been updated with the commit SHA, PR link, and verification results.
