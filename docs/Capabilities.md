# Capabilities

This is the supported scope of **0.2.0-preview**, including approximations and
explicit fallbacks. A successful conversion or absence of diagnostics does not
establish Office fidelity. Features outside this inventory remain unsupported.
See [preview scope](PreviewRelease.md), [measured qualification](PreviewQualification.md)
and [diagnostics](Diagnostics.md) before relying on a conversion.

## Shared input and compatibility behavior

- Transitional OOXML DOCX/PPTX packages; Strict OOXML is diagnosed and may render
  missing content. Archive, XML, image and font bounds remain enforced.
- Exactly one AlternateContent representation: the first understood Choice,
  otherwise Fallback, otherwise no representation. Supported content inside
  ignorable ProcessContent wrappers is exposed to the ordinary format reader.
  Inherited namespace-prefix/QName declarations resolve at their original scope;
  nested wrappers and namespace wildcards are supported.
- MustUnderstand prefix lists produce a warning for unsupported or unresolved
  requirements in active content. The preview continues with partial handling;
  it does not implement complete MCE validation, preservation or schema processing.
  Other extension content keeps its existing reader fallback.
- Shared font resolution, TrueType embedding, glyph coverage fallback, managed
  image decoding and PDF primitives. Complex-script shaping and CFF embedding
  remain unsupported. Custom font identities must distinguish different programs.
- The bounded PNG decoder supports 1/2/4/8-bit grayscale and indexed images,
  plus 8-bit RGB, gray/alpha and RGBA. Palette transparency, transparent-color
  keys and explicit alpha use the shared image path. Sixteen-bit PNG inputs
  retain the unsupported-image fallback; complete PNG conformance is not promised.
- Cooperative cancellation, bounded caches/page staging, optional coarse host
  memory admission and atomic file publication. These are specific safeguards,
  not whole-heap accounting; see the [operating envelope](OperatingEnvelope.md).

## PPTX

Supported within the following constraints:

- Slide order/size, hidden-slide exclusion and per-slide master-shape suppression;
  common inherited backgrounds, shapes, theme colors and Latin theme fonts.
- Common solid shapes, paths, straight lines, rotation/flips and nested group
  translation/scaling. Advanced/custom geometry and effects remain partial.
- JPEG/PNG pictures and bounded native SVG paint. Crop/fill clipping, source
  transforms and supported rectangle/ellipse/rounded/polygon picture masks share
  paint geometry. Unsupported masks retain diagnosed rectangular frames; bounded
  supported shadows/blur retain their qualified fallbacks.
- Fixed-grid tables with fills, explicit borders, merged continuations/row spans,
  text and calibrated built-in styles. Vertical anchoring and typography remain
  approximate; arbitrary automatic table layout is not promised.
- External URI and internal-slide links on supported shapes, pictures, groups,
  frames, tables, text runs and chart text, using their resulting transformed bounds.
  Unresolved/action-only links are diagnosed and omitted.
- Native cached-data bar/column, line, area, pie/doughnut, scatter, bubble and radar
  charts, with supported titles, legends, axes/ticks, labels and source formatting.
  Axes/data labels share number formatting, including dates, numeric sections,
  scaling, fractions and supported accounting/currency forms. Unsupported syntax,
  formula-only data without cached values and unsupported chart kinds are diagnosed.
  Optional workbook styles and shared strings follow that workbook's own
  relationships; unreferenced parts do not affect cached chart rendering.
- Failed slide nodes rewind partial paint/annotations and report a diagnostic;
  neighboring nodes may still render. Inspect diagnostic severity and fallback.

Approximate or restricted:

- Textboxes support paragraphs/runs, Latin greedy wrapping, font styling and
  alignment. Font metrics, kerning, autofit, distributed text, overflow, rotation,
  columns and table/bullet combinations retain qualified approximations. Nominal
  advances and explicit-zero kerning apply only in the qualified unrotated,
  horizontal, single-column no-wrap/noAutofit frames; no blanket parity is claimed.
- Automatic numbering covers the implemented Arabic, alphabetic and Roman
  period/parenthesis forms with supported levels/starts and sequence resets.
  Empty/plain paragraphs and changed settings affect continuation; richer formats
  and table combinations remain restricted.
- Clustered/stacked bar/column, line-with-markers and plain pie are closer-parity
  targets. Area, scatter/bubble, secondary axes, percent-stacked and leader-label
  layouts remain approximate; trendlines are diagnosed and omitted.
- Automatic chart layout, label strips/baselines, dense intervals, source crossings,
  reversed axes and manual-layout combinations retain their semantic guards.
  Qualified automatic single-line plots share axis/label placement; authored
  coordinates, composed/secondary plots, unsupported scales and unresolved fonts
  keep their fallbacks. Empty manual metadata retains automatic measurement.
- Qualified gallery defaults, authored stroke/fill precedence, line smoothing,
  markers and point-count color regimes remain implemented. Classic/composed
  gallery roles, minor-stroke defaults, automatic tick fitting and chart color
  parity are incomplete. Existing independent Office thresholds still apply.
- Horizontal shape text renders an ellipsis marker when vertical overflow hides
  later lines. Non-horizontal text and chart-title contexts retain diagnosed
  clipping fallbacks. Complex effects and typography may be simplified or omitted
  with diagnostics.

Unsupported: stock/surface/3D charts, SmartArt without a saved drawing part, live OLE, macros/ActiveX,
audio/video, animation/transitions, unrestricted custom effects, bidirectional
layout and general OpenType shaping. Static fallback diagnostics are listed below.

## DOCX

- Optional styles, settings, numbering, fonts/themes and related stories follow
  the document's owning relationships. Unreferenced package parts retain the
  corresponding empty/default behavior rather than configuring active content.
- Header/footer rendering diagnostics follow the same typed internal story
  references as their readers. A relationship alone does not select a story;
  selected unsupported payloads and package-level warnings remain diagnosed.
Supported within the following constraints:

- Body paragraphs/runs, common run/paragraph/document defaults and styles,
  margins/page sizes, paragraph spacing/line-height inheritance and basic page
  breaking. Qualified section/page/column breaks retain their explicit limits.
- Latin font resolution through explicit families, themes and font-table alternates.
  Undeclared prose/list families use the documented Times New Roman fallback when
  available, after an explicit resolver document default. Missing faces/glyphs are
  diagnosed; no general shaping or font-independent Office parity is promised.
- Simple decimal/bullet numbering and twip-based indents/tabs. Numbering font
  selection follows numbering-level values and the paragraph mark/style/defaults,
  independently of first-run styling. Supported inserted marks preserve adjacent
  text/equations; old/new review renumbering and marker revision decoration remain
  incomplete. Deleted/unnumbered marks keep their existing fallbacks.
- Fixed-grid tables with fills, common collapsed border styles/conflicts,
  nil/none suppression, repeated headers, split rows and cell content. Actual page
  occupancy controls following-prose admission. Complex implicit borders,
  vertical merges and nested/fragmented layout retain qualified restrictions.
- Inline raster/SVG pictures in qualified body/cell/header/footer containers,
  including supported crop/flips/rotation and source frames. Picture-only body
  paragraphs use occupied height and collapsed spacing. Floating pictures/textboxes
  have bounded positioning, clipping, z-order and first-pass wrap exclusions;
  arbitrary Word reflow around floating content is unsupported.
- Default headers/footers, simple page fields, supported notes and source-owned
  links/bookmarks. Static-story placement shares physical frames with painting
  and annotations. Complex/continued/page-dependent note compositions and arbitrary
  nested stories retain explicit fallbacks and documented geometry residuals.
- Simple/complex fields with stored results, including supported nested cached
  fields and bounded dynamic placeholders. This does not evaluate general Word
  fields, live cross-document data or arbitrary aligned/wrapping field layouts.
- Final, Original, SimpleMarkup and AllMarkup views, with preserved-layout and
  Word-compatible review geometry kept distinct. Public defaults are Final with
  preserved document layout. See [markup modes](DocxMarkupModes.md) for revision,
  comment, field and margin behavior.

Approximate or restricted:

- Supported native equations share ordered inline measurement with text/pictures.
  Fitting multiple formulas, bookmarks and unchanged-format run splitting preserve
  content in qualified body paragraphs, simple cells and direct headers/footers.
  Prepared display rows reuse math boxes and source frames. Mathematical layout
  emits DOCX_APPROXIMATED_EQUATION; omitted unsupported math is diagnosed separately.
- Bounded literal/fraction/script/root/delimiter/n-ary/function/matrix/cases forms,
  supported revisions, empty arguments and ready-made matrix fences retain
  individual metric/paint ownership. Unknown/deeper structures, continuation or
  unresolved merge owners, unfit objects and arbitrary fragmentation keep fallbacks.
  See [native equations](docx-equations.md) for the authoritative construct and
  placement limits; support for one construct does not imply unrestricted nesting.
- Pagination, line leading, kerning, numbering, table/story geometry and complex
  composition remain approximate. Font/size/document-shape guards protect independently
  qualified behavior; they are not promises for every combination.
- Review balloons, ranges, revision bands, comment collision/layout and note-marker
  revisions remain partial. Remaining typography/geometry and extraction failures
  are recorded in [document review](PreviewVisualReview.md); successful conversion
  does not turn a failed visual threshold into a pass.
- SVG gradients follow Office-visible centered focal geometry and smooth preview
  transparency. Broader SVG composition, transformed picture-link combinations
  and some alpha/color effects retain restrictions and explicit residuals.

Unsupported: full Word balloon/collision pagination, floating charts, SmartArt,
live OLE/macros, unrestricted section/column balancing, arbitrary nested/floating
story reflow, complex scripts/bidirectional shaping and unlisted equation constructs.
Detected unsupported content may be simplified or omitted; diagnostics are not
a complete inventory of every feature in an arbitrary document.

## PDF Output

- Deterministic resource naming/object order for stable input, fonts and options;
  optional fixed creation-date metadata. There is no Deterministic option/mode.
- Subset embedded TrueType/CID fonts with ToUnicode, plus standard Type1/WinAnsi
  fallbacks. Subsets retain required compound glyphs; assembled structural symbols
  may use ActualText. CFF/OpenType-CFF outlines are substituted and diagnosed.
- Supported path/fill/stroke/transform/clipping/text operators, raster images,
  bounded gradients/patterns, transparency Form XObjects and soft masks.
- External/internal Link annotations for supported DOCX/PPTX source links.
- Minimal metadata; no interactive AcroForms, other annotation kinds, outlines,
  embedded files, tagged PDF, audio/video or JavaScript. No PDF/A or PDF/UA claim.
  See the source-backed [PDF profile](PdfProfile.md) for emission and validation scope.

Detailed historical calibration and qualification narratives remain in
[the prior capability record](https://github.com/Lokad/OoxPDF/blob/b4ae23938be7df60bb6b1be09985e73fc2e2a523/docs/Capabilities.md).
Condensing this inventory does not remove those rendering rules, evidence or
thresholds, and does not expand supported combinations.


## Static-content fallback matrix

Each row links an unsupported family to its detection diagnostic, fallback, and covering tests (synthetic in-code fixtures). Severities and dedup scope are defined in Diagnostics.

| Family | Detection | Fallback | Covering tests |
|---|---|---|---|
| PPTX stock, surface, and 3D charts | `PPTX_UNSUPPORTED_CHART` Warning | Ignored | `PptxStockChartKindWarnsUnsupportedChart` (pptx-model) |
| PPTX SmartArt without a saved drawing part | `PPTX_UNSUPPORTED_SMARTART` Warning | Ignored (node text kept invisibly) | `PptxUnsupportedFeaturesEmitDiagnostics` (pptx-core) |
| PPTX OLE objects | `PPTX_UNSUPPORTED_OLE_OBJECT` Warning | Ignored | `PptxUnsupportedFeaturesEmitDiagnostics` (pptx-core) |
| PPTX audio, video, animation, transitions | matching `PPTX_UNSUPPORTED_*` Warning | Ignored | `PptxUnsupportedFeaturesEmitDiagnostics` (pptx-core) |
| PPTX missing image parts | `IMAGE_MISSING_PART` Error | Ignored, conversion continues | `PptxMissingImagePartDiagnosesErrorAndContinues` (pptx-images) |
| DOCX missing picture assets | `IMAGE_MISSING_PART` Error | Retain supported picture frame, omit pixels | DocxImageRecoveryTests (rendering-portable) |
| Undecodable images | `IMAGE_UNSUPPORTED_FORMAT` Error | Ignored | DocxImagesTests, ImagingTests, PptxImagesTests |
| Cropped-image decode | `IMAGE_CROP_UNSUPPORTED_FORMAT` Warning | PDF clipping | PptxImagesTests |
| SVG pictures | `SVG_UNSUPPORTED_CONTENT` Error/Warning | Ignored / Partial / RasterImage | PPTX SVG diagnostics; DOCX native/preview fallback and viewport tests |
| DOCX floating charts | `DOCX_UNSUPPORTED_CHART` Warning | Ignored | `DocxUnsupportedChartInHeaderIsDiagnosed` (docx-core) |
| DOCX SmartArt | `DOCX_UNSUPPORTED_SMARTART` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| Supported DOCX body equations | `DOCX_APPROXIMATED_EQUATION` Warning | Approximated | `DocxNativeStructuredSequencesRetainSourceOrderAndIndependentRules` (docx-core) |
| Unsupported DOCX equation constructs/placements | `DOCX_UNSUPPORTED_EQUATION` Warning | Ignored | `DocxUnsupportedEquationInFooterIsDiagnosed` (docx-core) |
| DOCX OLE objects | `DOCX_UNSUPPORTED_OLE_OBJECT` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| DOCX external images | `DOCX_UNSUPPORTED_EXTERNAL_IMAGE` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| DOCX macros | `DOCX_UNSUPPORTED_MACRO` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| Strict OOXML | `OOXML_STRICT_DIALECT` Warning | May render missing | OoxmlTests dialect probes (ooxml) |
| Font packs | `FONT_PACK_*` throw codes | Conversion fails with DiagnosticId | FontTests pack suites (fonts) |
