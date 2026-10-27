# Diagnostics

Diagnostics use stable code prefixes such as `OOXML_`, `PPTX_`, `DOCX_`, `PDF_`, `FONT_`, and `IMAGE_`.

Diagnostics are emitted through `OoxPdfOptions.DiagnosticSink` and through the CLI `--diagnostics` JSON file. Warnings mean conversion continued with a fallback or omission. Errors mean conversion failed or the output cannot be trusted.

## OOXML

Top-level package or dialect issues:

- `OOXML_STRICT_DIALECT`: Strict OOXML (ISO 29500) content was detected; only the transitional dialect is supported and content may be missing.
- `OOXML_MUST_UNDERSTAND`: a `mc:MustUnderstand` namespace-prefix list requires an unsupported or unresolved namespace. Conversion continues with a warning and partial compatibility handling, not full MCE conformance. Requirements use their declaration scope; unselected alternate branches do not participate. Main parts warn once per document; slides, shared masters/layouts, and DOCX story and infrastructure parts warn once per part with their part names.

## PPTX

Unsupported feature warnings:

- `PPTX_UNSUPPORTED_ANIMATION`: slide timing or animation content was detected and ignored.
- `PPTX_UNSUPPORTED_AUDIO`: audio content was detected and ignored.
- `PPTX_UNSUPPORTED_CHART`: an unsupported chart kind was detected, a chart part was missing/unresolvable, or a supported chart referenced formula-only data without cached numeric values (see also `PPTX_CHART_MISSING_CACHED_DATA`). Native rendering covers bar/column, line, area, pie/doughnut, scatter, bubble, and radar charts with cached values.
- `PPTX_UNSUPPORTED_OLE_OBJECT`: embedded OLE content was detected and ignored.
- `PPTX_UNSUPPORTED_SMARTART`: a SmartArt diagram has no readable saved drawing part, so its shapes were not painted. Diagrams with a drawing render natively without this diagnostic. When the data model has node text, that text is kept in the PDF text layer, invisibly.
- `PPTX_UNSUPPORTED_TRANSITION`: slide transition content was detected and ignored.
- `PPTX_UNSUPPORTED_VIDEO`: video content was detected and ignored.
- `PPTX_CHART_MISSING_CACHED_DATA`: a supported chart referenced formula-only data without chart-side cached numeric values; workbook provenance is preserved but not used for layout.
- `PPTX_UNSUPPORTED_CHART_AXIS_TITLE_LAYOUT`: a chart axis title with manual positioning that cannot be honored (default-placement titles render natively).
- `PPTX_UNSUPPORTED_CHART_AXIS_TITLE_AXIS_POSITION`: a default-placement chart axis title has an unsupported or missing axis kind/position.
- `PPTX_UNSUPPORTED_CHART_TRENDLINE`: a chart series defines a trendline, which is not rendered.
- `PPTX_UNSUPPORTED_CHART_NUMBER_FORMAT`: a chart number format uses unsupported syntax (locale rules, colors, native digit shapes, `!` escapes, `@` placeholders, or unknown brackets) that is dropped; one warning per construct per chart.
- `PPTX_UNSUPPORTED_GRAPHIC_FRAME`: an unsupported graphic frame was detected and ignored.
- `PPTX_UNSUPPORTED_GRADIENT_FILL`: an unsupported gradient fill was detected and ignored.
- `PPTX_UNSUPPORTED_PATTERN_FILL`: an unsupported pattern fill was detected and ignored.
- `PPTX_UNSUPPORTED_TEXT_ORIENTATION`: an unrecognized text orientation was detected and ignored (horz, vert, vert270, eaVert, mongolianVert, and the wordArt variants render).
- `PPTX_UNSUPPORTED_TEXT_OVERFLOW`: an unsupported ellipsis context (such as non-horizontal text or chart titles) uses local clipping without a marker. Ordinary horizontal shape text renders the marker without this diagnostic.
- `PPTX_UNSUPPORTED_PICTURE_FILL`: an unsupported picture fill was detected and ignored.
- `PPTX_UNSUPPORTED_IMAGE_TILE`: a tiled image fill was detected and ignored.
- `PPTX_UNSUPPORTED_IMAGE_RECOLOR`: an unsupported image recolor was detected.
- `PPTX_UNSUPPORTED_TRANSPARENCY`: unsupported transparency was detected and ignored.
- `PPTX_UNSUPPORTED_EFFECT`: an unsupported effect was detected and ignored.
- `PPTX_UNSUPPORTED_CUSTOM_GEOMETRY`: unsupported custom geometry was detected and ignored.
- `PPTX_UNSUPPORTED_CALLOUT`: a callout shape was detected and ignored.
- `PPTX_UNSUPPORTED_TABLE_STYLE`: an unsupported table style was detected.
- `PPTX_UNSUPPORTED_HYPERLINK`: a shape, frame, text-run, or chart title hyperlink pointed at an unresolvable, empty, unknown-slide, or unsupported-typed relationship and was ignored. Repeated uses of the same relationship on one slide report once.
- `PPTX_UNSUPPORTED_HYPERLINK_ACTION`: a hyperlink carried only an action (no relationship target) and was ignored. Repeated uses of the same action on one slide report once.
- `PPTX_NODE_RENDER_FAILED`: a slide node failed to render.
- `PPTX_PICTURE_GEOMETRY_UNSUPPORTED`: an ordinary picture has custom or unsupported preset geometry; paint retains its rectangular frame.

These warnings are slide-scoped when a slide index is available. Duplicate occurrences of the same unsupported feature on one slide are aggregated into one warning for that slide.

## DOCX

Unsupported feature warnings:

- `DOCX_UNSUPPORTED_CHARACTER_UNIT_INDENT`: character-unit paragraph indents are ignored; twip-based indents render.
- `DOCX_UNSUPPORTED_CHART`: a DrawingML chart payload in body or related-story text was detected and ignored, with story provenance; one warning per document.
- `DOCX_UNSUPPORTED_IMAGE_TRANSFORM`: picture rotation has unusable or unresolved source/line dimensions; unrotated paint is retained, including supported crop/flips.
- `DOCX_UNSUPPORTED_IMAGE_CROP`: the picture source rectangle is malformed, non-finite or empty; the uncropped image is retained.
- `DOCX_UNSUPPORTED_COMMENTS`: comment markup was detected and ignored, usually in compatibility/default final output when markup printing was not requested.
- `DOCX_UNSUPPORTED_COMPLEX_FIELD`: a complex field was malformed in a way that prevents cached-result or placeholder rendering, nested inside another field instruction, or had no supported dynamic placeholder and no cached result to render.
- `DOCX_UNSUPPORTED_EQUATION`: Office Math content was detected and ignored.
- `DOCX_APPROXIMATED_EQUATION`: native Office Math was painted with approximate layout; one warning per affected part. This is distinct from omitted unsupported equations.
- `DOCX_UNSUPPORTED_EXTERNAL_IMAGE`: an externally referenced image was detected and ignored because only packaged parts render; one warning per document.
- `DOCX_UNSUPPORTED_FLOATING_DRAWING`: floating DrawingML content was malformed or used an unsupported payload such as a chart, SmartArt, external image, or anchor positioning outside the supported image/text-box placement model.
- `DOCX_UNSUPPORTED_FORMATTING_REVISIONS`: formatting revisions were detected outside a mode that approximates them, so they are ignored; the AllMarkup approximation is `DOCX_APPROXIMATED_FORMATTING_REVISIONS`.
- `DOCX_UNSUPPORTED_MACRO`: a VBA project was detected and ignored.
- `DOCX_UNSUPPORTED_MANUAL_BREAK`: an unsupported manual column-break container was detected; visible body column breaks are supported.
- `DOCX_UNSUPPORTED_MULTI_COLUMN`: a multi-column section was detected and rendered as a single column.
- `DOCX_UNSUPPORTED_OLE_OBJECT`: embedded OLE content was detected and ignored.
- `DOCX_UNSUPPORTED_SECTION_BREAK`: a continuous or unknown paragraph section break was detected; support is partial.
- `DOCX_UNSUPPORTED_SMARTART`: a SmartArt diagram payload in body or related-story text was detected and ignored, with story provenance; one warning per document.
- `DOCX_UNSUPPORTED_TRACKED_CHANGES`: tracked insertion or deletion markup was detected but the selected mode did not request visible markup support.
- `DOCX_UNSUPPORTED_VML`: VML drawing content outside the supported inline image subset was detected and ignored.

Approximation warnings:

- `DOCX_APPROXIMATED_COMMENTS`: comment markup was rendered through the selected markup mode with first-pass markers or balloons.
- `DOCX_APPROXIMATED_ENDNOTE`: endnote references render through first-pass related-story bodies; split-note continuation stays approximate.
- `DOCX_APPROXIMATED_FOOTNOTE`: footnote references render through first-pass related-story bodies; split-note continuation stays approximate.
- `DOCX_APPROXIMATED_FORMATTING_REVISIONS`: formatting revisions were detected and surfaced through private-safe property-family provenance, first-pass formatting balloons, change bars, or approximation diagnostics.
- `DOCX_NUMBERING_INDENT`: a numbering indentation variant outside the supported twip-based left/right/first-line/hanging/tab model was detected, such as character-unit list indents.
- `DOCX_STYLE_PARAGRAPH_SPACING`: a paragraph style used a spacing variant outside the supported before/after, line-unit, automatic, contextual, exact, auto, and at-least line-spacing model.
- `DOCX_TABLE_TEXT_DIRECTION`: table cell text-direction variants outside the supported model are approximated.
- `DOCX_TABLE_BORDER_STYLE`: a table or table style used a border style outside the supported solid, thick, double, triple, dotted, dashed, dash-small-gap, dash-dot-stroked, dot-dash, dot-dot-dash, thin/thick compound, wave/double-wave, 3D emboss/engrave, outset, inset, nil, and none set.
- `DOCX_APPROXIMATED_TRACKED_CHANGES`: tracked changes were rendered through the selected markup mode with final/original filtering, change bars, inline revision styling, grouped first-pass revision balloons, or compact overflow continuations, but without full Word-style review-pane geometry.

These warnings are document-scoped. Duplicate occurrences of the same unsupported feature in one document are aggregated into one warning.

## Fonts

- `FONT_UNSUPPORTED_OUTLINES`: a CFF/OpenType-CFF typeface was detected and not embedded (native CFF embedding is unsupported). DOCX substitutes the document fallback typeface for primary runs, skips CFF per-character fallback candidates, and degrades CFF-only fallbacks to the missing-font path; PPTX regroups uses through its per-glyph fallback. One warning per typeface per conversion.

- `FONT_NO_USABLE_FACE`: document text had no usable embeddable font, so it renders with the built-in Helvetica fallback (standard Type1, never embedded) using deterministic average advances. Page breaks and extractable text are preserved; characters outside WinAnsi appear as "?". In PPTX, fallback spans stay link-free and chart titles render through the same seam. One warning per typeface per conversion.

- `FONT_MISSING_GLYPHS`: codepoints with no glyph in any usable face render as "?" through the primary face (the subset carries "?" exactly when needed). One warning per typeface per conversion, listing the affected codepoints in order with a remainder count.

- Font-pack download, hash, validity, and presence failures throw OoxPdfFontPackException carrying `FONT_PACK_DOWNLOAD_FAILED`, `FONT_PACK_HASH_MISMATCH`, `FONT_PACK_INVALID`, or `FONT_PACK_MISSING` in DiagnosticId instead of emitting a sink diagnostic.

- `COMPLEX_SCRIPT_APPROXIMATION`: text needing joining, reordering, bidirectional reordering, or mark positioning renders without shaping, with glyphs emitted in source order. One warning per behavior family per conversion, with Feature naming the family and Fallback Unshaped glyphs.

## Images and SVG

- `IMAGE_MISSING_RELATIONSHIP`: error, a PPTX picture's image relationship could not resolve to an embedded target. The image is omitted and conversion continues; CLI `--strict` returns 3. The source story part and slide identify the failed reference.
- `IMAGE_MISSING_PART`: error, a referenced image part could not be resolved. Its pixels are omitted; supported DOCX picture frames retain their occupied space. Conversion continues with the diagnostic; CLI `--strict` returns failure.
- `IMAGE_UNSUPPORTED_FORMAT`: error, image bytes could not be decoded or the content type is unsupported, so the image was ignored.
- `IMAGE_CROP_UNSUPPORTED_FORMAT`: a cropped image could not be decoded for Office-style embedding, so PDF clipping is used instead.
- `SVG_UNSUPPORTED_CONTENT`: unparsable SVG pictures and unusable viewBoxes are errors with the picture ignored; unsupported elements, path commands, gradients, and paint are warnings with the affected content partially omitted. One diagnostic per construct kind.
  Numeric zero-opacity root/group containers do not paint or activate warnings
  for their contents. Uniform numeric stop-opacity attributes on referenced
  gradients multiply fill/node alpha. Numeric partial opacity on the outer SVG
  and group containers isolates supported paint, up to 32 nested groups. Partial
  CSS/percentage container alpha, excluded nested SVG/deeper containers and varying,
  CSS or percentage stop opacity retain warnings for ignored behavior. Unused
  definitions and effective unit opacity do not add these warnings; inline
  style overrides presentation attributes.

- `DIAGNOSTIC_OVERFLOW`: the collector retained the first diagnostics and dropped further occurrences; per-ID totals stay available on the collector. Info-only overflow is informational and does not affect CLI strict mode. If any warning/error was observed, the marker remains an error and strict mode still reports it, including when the original warning/error was dropped.

The CLI writes diagnostics JSON when `--diagnostics <file>` is provided. The JSON is an array of diagnostic entries with these fields when available:

- `Id`: stable diagnostic code.
- `Severity`: numeric enum value: `0` = `Info`, `1` = `Warning`, `2` = `Error`.
- `Message`: human-readable explanation.
- `PartName`: OOXML package part where the issue was detected.
- `SlideIndex`: one-based slide index for slide-scoped PPTX diagnostics.
- `PageIndex`: one-based page index when available.
- `Feature`: short unsupported or approximated feature name.
- `Fallback`: short description of the fallback, such as `ignored` or `approximated`.

Exit codes:

- `0`: conversion succeeded.
- `1`: conversion failed and an `OOXML_CONVERSION_FAILED` error diagnostic is emitted when a diagnostics path was supplied.
- `2`: invalid arguments.
- `3`: conversion succeeded, but `--strict` saw at least one warning or error diagnostic.

## Output bounds and host admission

`OoxPdfOptions.ConversionLimits` bounds PDF output and resource payloads, not
cumulative work or total managed memory. The current defaults are:

| Limit | Default |
|---|---|
| `MaxPagesPerConversion` | 10,000 pages |
| `MaxPdfContentBytesPerConversion` | 1 GiB encoded page/group content |
| `MaxOutputBytesPerConversion` | 2 GiB serialized PDF |
| `MaxPdfFontBytesPerConversion` | 256 MiB uncompressed embedded-font/ToUnicode payloads before PDF stream compression |
| `MaxPdfImageBytesPerConversion` | 512 MiB serialized images/masks |
| `MaxResidentPageContentBytesPerConversion` | 64 MiB resident page/group payload |

Archive, XML, image, font and individual expansion bounds apply independently.
Cancellation is cooperative. The retired cumulative work and nonrefunding
retained-byte limits are listed in [caller migration](PublicApiChanges.md).

The optional `MemoryAdmission` host adapter acquires one coarse lease before
conversion construction. It covers retained storage and construction peaks;
refusal throws `OoxPdfLimitExceededException` before input reads or PDF writes.
Ownership is released on success, failure and cancellation, before file
publication. Caller-owned streams/resolvers and shared bounded caches remain
outside that lease. Size shared host capacity from measured representative
workloads and concurrency headroom; this is not comprehensive heap accounting.

Page/group content beyond the resident window spills into a delete-on-close
file. Informational `PDF_PAGE_CONTENT_SPILLED` is emitted once on escalation.
With `ReportResourceUsage`, `CONVERSION_RESOURCE_SUMMARY` reports `pages`,
`pdfPages`, `pdfContentBytes`, `pdfOutputBytes`, `pdfFontBytes`, `pdfImageBytes`
and `pageContentSpilledBytes`. These combine physical PDF output bytes with
resource payload and spill totals, rather than measuring live memory.
`pdfFontBytes` and its limit count uncompressed embedded-font programs and ToUnicode
data before PDF stream compression; PDF dictionaries and Standard-14 ToUnicode
maps are excluded. This total can exceed the complete PDF's size. Use
`pdfOutputBytes`/`MaxOutputBytesPerConversion` for the complete serialized PDF.
Informational diagnostics do not affect CLI `--strict` exit codes. Existing
diagnostic IDs, numeric fields and message format are unchanged.

The summary is emitted before serialization for stream output and after staging
but before atomic publication for file output. Writer-stage totals are zero in
the stream summary. A throwing observer aborts conversion. Supplied streams stay
caller-owned; write failures may leave a prefix. File conversion preserves the
previous destination on rendering, observer, cancellation, limit or publication
failure and cleans its owned staging. Normal return signals successful output;
there is no separate committed-notification diagnostic.

Historical performance measurements and their workload limits are documented in
[the operating envelope](OperatingEnvelope.md). They do not establish current
preview capacity or a comprehensive memory ceiling.

## Code Conventions

Use stable prefixes by subsystem:

- `OOXML_`: package, ZIP, relationship, XML, or top-level conversion issues.
- `PPTX_`: PowerPoint-specific parser or renderer issues.
- `DOCX_`: Word-specific parser or renderer issues.
- `PDF_`: PDF writer issues.
- `FONT_`: font resolution, parsing, or embedding issues.
- `IMAGE_`: image parsing or rendering issues.

Prefer one durable code per observable behavior. Aggregate repeated unsupported features when the exact count is not useful to the caller.
