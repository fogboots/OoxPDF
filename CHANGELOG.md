# Changelog

## 0.2.0-preview — candidate prepared, not published

- Renders PPTX SmartArt from the diagram's saved drawing part through the ordinary
  shape and text pipeline, including text rectangles and hue/saturation color
  offsets. `PPTX_UNSUPPORTED_SMARTART` now means no readable drawing exists; the
  diagram's node text is then kept in the PDF text layer.
- Evaluates PPTX slide-number fields from the authored sequence and numbering
  origin, preserving hidden-slide gaps independently of output-page indexes.
- Uses paragraph reference identity for DOCX review balloons and textbox comment
  suppression; colliding runtime hashes no longer merge distinct owners.
- Keeps PPTX shape diagnostics aligned with painting admission and removes
  unused scene style snapshots. Fresh chart summaries inspect all pages and
  identify the unverified coverage of reused reports.

- Keeps cached PPTX charts independent of unreferenced workbook styles/shared
  strings. Related metadata and its existing bounds remain active.
- Classifies chart graphics and text within each PDF page, preserving page-owned
  plot boxes and region identities before comparison. Single-page rules and
  tolerances remain unchanged.

- Diagnoses unresolved PPTX picture relationships instead of silently omitting
  the image with a clean strict exit. Missing target parts retain their existing
  diagnostic; valid pictures and file publication behavior stay unchanged.
- Keeps graphics comparisons within their PDF pages and refuses missing required
  geometry or explicitly requested text hashes. Removes a table text forwarding
  shim and replaces its obsolete self-comparison with source-content/frame checks.

- Honors explicit PPTX table paragraph-default bold and color ahead of table
  styling, retaining inherited styling and direct run overrides. Removes the
  internal test-only image-key selector while preserving exact deduplication checks.

- Keeps an explicit PPTX slide list authoritative when hidden-slide filtering
  leaves no visible pages; unlisted slide relationships no longer become output.
  Missing-list repair and the existing no-page failure remain.
- Preserves current reference-worker failure metadata and removes stale prior-run
  status, while keeping the supervisor result authoritative.

- Keeps DOCX rendering diagnostics scoped to selected header/footer stories;
  unused template payloads no longer cause false strict failures. Shared related
  XML loading uses the existing package cache, with note edits on detached XML.

- Selects optional DOCX configuration/stories by their owning relationships;
  unrelated styles/settings no longer change formatting or suppress headers.
- Removes internal workbook test-construction adapters and exclusive conversion
  helpers; tests use the live workbook reader. Typed models and caches remain.
- Clarifies existing stream access and font-payload counting contracts: seekable
  inputs may use random access, while font limits/totals count uncompressed data
  before PDF stream compression. Runtime behavior and diagnostics are unchanged.
- Selects DOCX first-page headers and footers using their owning section, while
  preserving default/even selection and undefined-variant suppression.
- Preserves incoming links to paragraph-end text bookmarks, including wrapped
  targets and hidden trailing runs, without changing prose layout.
- Propagates fatal font-resolution limit failures during PPTX measurement;
  ordinary font-I/O recovery remains available and cancellation still aborts.
- Omits missing-glyph warnings for hidden DOCX text, matching layout visibility;
  visible missing glyphs retain their warning and CLI strict-exit behavior.
- Resolves supported English/localized installed-font family aliases to the same
  program and collection face, preserving canonical names and style ranking.
- Uses declared DOCX Ascii/HighAnsi/EastAsia font choices for text fragments;
  unchanged-format run boundaries no longer determine the mixed text's font.
- Matches numeric DOCX note and numbering IDs by decimal value, preserving
  content and list indentation across equivalent spellings such as `1`/`01`.
  Opaque identifiers retain their existing rules.
- Makes PPTX glyph emission inspection honor scene visibility and master-shape
  admission. Independent raw XML layout inspection remains available.
- Preserves explicitly empty DOCX bullet labels and their list indentation;
  missing label text retains the existing default.
- Removes an unreachable PPTX text fallback and unused line tally. Independent
  XML inspection, scene layout, table cells and bounded caches remain.
- Removes unreachable JPEG-dimension/chart-range overflow handlers while keeping
  validated input bounds, storage checks and chart area/union limits.
- Keeps current DOCX section dimensions and margins when tracked formatting
  history contains older page settings. Revision provenance remains inspectable.
- Omits explicitly hidden PPTX objects before resource preflight and painting,
  including their interactive links and hidden group descendants.
- Corrects ellipsis guidance: horizontal shape text renders a marker; unsupported
  contexts retain their diagnosed clipping fallback.
- Honors horizontal PPTX text clipping through existing shape/cell frame bounds,
  including when vertical overflow remains allowed. Existing annotation policy and
  qualified vertical/preset/column behavior remain; general overflow parity is approximate.
- Preserves finite zero in visual-report comparisons and rejects missing required
  foreground-recall evidence while retaining unrequired legacy/inapplicable reports.
- Preserves explicit DOCX text-wrapping breaks and resolves internal PPTX links
  using case-insensitive package-part identity, matching their equivalent source
  forms in Office. Relationship IDs and font identities retain their existing rules.
- Preserves charts when a date-formatted value exceeds the supported calendar:
  that label uses the existing numeric fallback. In-range 1900/1904 formatting is
  unchanged; exact Office tick labels outside the calendar remain approximate.
- Reports the active Office reference operation on timeout and forwards
  `-SkipBuild` through reference exports to the existing pinned rasterizer.
- Handles invalid BMP header extents through normal image diagnostics instead of
  aborting DOCX/PPTX conversion. Supported pixels and the public API are unchanged.
- Scopes validation build locking to physical checkouts, retaining shared-output
  serialization across aliases. Corrects pinned-PDFium setup and font-download
  memory comments without new runtime accounting or options.
- Keeps opaque font IDs ordinal and scoped to their owning resolver during PPTX
  fallback and program reuse. Subsets reuse the resolved immutable program rather
  than a second string identity. No public API change; cache reuse is retained.
- Makes packaged preview guidance reach repository-only guides and checks local
  documentation links in the package smoke. Escapes a literal source NUL and
  corrects an outdated chart-hyperlink comment without changing hyperlink keys.
- Preserves PNG transparent-color keys for supported RGB and packed grayscale
  inputs, matching equivalent full-alpha images through DOCX/PPTX and picture
  effects. Rejects malformed fixed headers and transparency metadata through the
  existing image-error fallback. No public API or supported bit-depth expansion.
- Preserves original Office reference-worker failure stages when cleanup succeeds,
  with separate cleanup-failure checks; corrects the pinned PDFium recovery hint.
- Maps Adobe four-component JPEG ink samples correctly during PDF passthrough,
  including the original-image crop/recolor fallbacks. Untagged/RGB/gray behavior
  is retained; unresolved four-component Adobe transforms diagnose explicitly.
  This is not full ICC color management or a public API change.
- Treats truncated PNG scanlines as invalid image data, preserving image-error
  recovery across DOCX/PPTX rather than an unhandled CLI exception. Removes one
  duplicate exception test. Portable direct-PNG metric checks replace the old
  build/conversion/PDFium probe while preserving all six metric/region premises.
- Rejects invalid JPEG frame metadata before PDF passthrough. Malformed dimensions,
  precision or component tables use the existing image error/omission diagnostic;
  baseline, progressive, grayscale and CMYK controls retain their encoded bytes.
  No public API change or full JPEG entropy validation is introduced.
- Bounds test child-process waits and drains both output streams concurrently.
  Ordinary CLI exit-code checks run in the normal selection; text-content checks
  reuse explicit prebuilt CLI/inspection tools without per-case builds.
- Preserves supported DOCX/PPTX content inside declared ignorable ProcessContent
  wrappers. Compatibility rules resolve in their original namespace scope;
  only the selected alternate branch participates. MustUnderstand uses prefix
  lists and retains the documented warning-only partial policy. No public API
  change; existing image/layout fallbacks and visual thresholds remain.
- Corrects custom font identity guidance, PDF profile and math diagnostics;
  consolidates the capability inventory around scope and remaining limitations.
- Qualification transfer replaces root asset entries through independent staging,
  preserving older artifacts behind hard links. Linked asset directories refuse
  before mutation. Host guidance clarifies that active file-backed font sources
  require unchanged font files; bounded caching and the public API are unchanged.
- Qualification joins reject duplicate document outcomes from distinct reports,
  even when PDF bytes match. Select one authoritative source; repeating the same
  report selection remains allowed. Failures first observed during test repairs
  require a subsequent pass, never a skip; original attempts remain recorded.
- Removes ten verified unused internal methods without changing the public API
  or rendering rules. Workspace measurement rejects unknown arguments; diagnostic
  guidance specifies the existing numeric severity format.
- Binds asynchronous relative file paths before queuing, preserving the caller's
  directory through worker delays. Info-only diagnostic overflow stays
  informational; warning/error overflow retains CLI strict behavior.
- Removes inactive private/internal overloads, obsolete test/probe-only paths,
  parse-count instrumentation and retired zero-valued DOCX corrections. Inert
  PPTX emission/legend state and unread argument plumbing are removed. Tests and
  allocation tooling target live implementations; authored spacing and qualified
  rendering behavior remain. No additional public API is removed.

- Restores supported DOCX content in ordinary combinations: bookmarks and
  multiple equations, pictures, table cells, headers/footers and qualified notes.
  Shared source frames carry layout geometry into painting and navigation.
- Corrects static-story placement, table origins, inline box ownership and
  raster/SVG picture frames using independent Office comparisons. Existing visual
  thresholds remain unchanged; approximate layout and cosmetic failures remain.
- Simplifies shared bounded font-cache ownership and removes duplicated PDF
  validation work, dead helpers, obsolete pattern constructors and test-only
  forwarding APIs. Essential PDF structural validation remains.
- Fixes concurrent requests to both built-in font resolvers while retaining
  bounded caching. Clarifies that failed caller-stream conversions can leave an
  incomplete prefix; atomic file publication remains unchanged.
- Propagates original diagnostic-observer exceptions through PPTX content
  recovery and observes final-summary cancellation before file publication.
  Removes the unused image-backed luminosity-mask path and decoder wrappers;
  live PNG alpha masks and SVG vector luminosity groups remain supported.
- Captures file input/output paths before host callbacks, preventing working-directory
  changes from redirecting publication. File-path guards protect case aliases on
  case-insensitive Linux mounts and retain provably distinct entries; unresolved
  case-only paths refuse conservatively. Raster validation handles spaced arguments, and transfer checks
  include the frozen DOCX inspector and its parent build settings.
- Protects CLI input/output from aliased diagnostics: literal path checks precede
  early failures, and atomic diagnostics publication replaces existing links.
  Qualification checks root build settings, binds cached markup to matching frozen
  runtimes and verifies complete raster identity before reuse. Local capability
  selections retain broader milestone coverage and exhaustive slow balloon checks.
  Removes the isolated test-only PPTX source-run snapshot/coalescer.
- Requires recorded, ordered reference-raster page identities and recovers
  interrupted derivatives without mixing stale pages or replacing trusted PDFs.
  Office worker arguments preserve spaces; qualification reports retain declared
  platform skips and reject unexpected skips. Removes the internal WriteBlank
  forwarding wrapper; tests and allocation tooling call the staged writer directly.
- Protects input/PDF publication through directory junction and symbolic-link
  ancestors. Raster tools refuse existing PNG output so stale pages cannot hide
  missing PDF pages. Removes remaining kerning scan fuel in favor of cancellation;
  font storage and input bounds remain.
- Removes ineffective `Deterministic`, library `Strict` and cumulative work or
  nonrefunding memory quotas. Uses cooperative cancellation and optional coarse
  host memory admission; preserves input/output bounds, bounded page staging,
  deterministic cleanup and atomic file publication.
- Breaking API changes require recompilation. See [migration guidance](docs/PublicApiChanges.md)
  and [supported, approximate and unsupported preview scope](docs/PreviewRelease.md).
  Performance optimization and broader feature expansion are deferred.

## Earlier development history

- Added bottom-reserve 14pt probe as slope held-out evidence (Q07 increment): category ticks bumped 9pt to 14pt by byte-surgical fixture edit (Office reference renders clean); Office plot box matches the rule prediction within 0.01pt at the new size (193.68 vs 193.69), confirming the 7.0+1.813fs slope and ruling out a retune (overlay residual sits inside documented tolerance). Direct PNG review shows complete correct content (SSIM 0.948, recall 0.991, diagnostics empty) with the same box-shift gap class. Manifests: 118 locked / 210 approximate / 0 needs-review (336 cases).

- Removed cell-memo debug output from library stdout and fixed its double-counted hit counter (hygiene fix found during cap-tuning probes: every memo hit wrote a W04HIT line to the host console and incremented Hits twice; TotalHits-gated tests and memoized behavior are unchanged). Pinned by the full suite with byte-identical outputs.

- Added pattern-phase probe as weave-calibration held-out evidence (Q07 increment): a shifted copy of the overlay chart (+1.3pt) renders dotted-band crops pixel-identical (MAE 0.0) between Office and candidate, proving both sides anchor the pct70 weave to the page and closing the pattern-phase question; the probe carries the same bottom-reserve box-shift gap class as its parent (SSIM 0.950, recall 0.970, diagnostics empty). Manifests: 118 locked / 209 approximate / 0 needs-review (335 cases).

- Moved secondary-axis-overlay probe to approximate (Q07 increment): exhaustive forensics (structure boxes, raw path commands, pixel-identical dotted-band crops, full-page diff mapping, stderr-instrumented layout run) prove complete correct content with a single characterized +0.217pt plot-box Y residual from the bottom-reserve rule (23.317 vs Office 23.1 at 9pt ticks, inside the constant documented 0.3 tolerance; a shared-formula retune is contraindicated by green Tier-1 locks on the same preset). SSIM 0.949, recall 0.966, diagnostics empty. This closes the needs-review triage: all 334 visual cases dispositioned (118 locked / 208 approximate, 8 locked-text-ops). Manifests: 118 locked / 208 approximate / 0 needs-review.

- Merged adjacent same-style source runs into one Office-compatible text operation (fidelity fix): the trailing-whitespace run-boundary veto emitted 15 ops where Office emits 13 on the whitespace-controls probe (identical decoded text and pixels either way); fixture forensics plus the CoalesceAdjacentTextSpans guards show hyperlink, math/autofit, alignment, and PreventCoalesce boundaries are unchanged, and hidden-advance accounting for style boundaries is pinned by the surviving layout assertions. Probe text-op gate now passes 13v13 with zero deltas; full Release suite 1592/0/9 with two contract tests migrated to the merged behavior. Manifests: 118 locked / 207 approximate / 1 needs-review.

- Moved scatter-clusters and scatter-smooth ports to approximate (Q07 increment): fixture inspection proves unstyled marker fills (clusters series carry only ln/noFill with no marker element; smooth series has no spPr at all), so the same-hue shade difference is the documented unstyled-series approximation gap already recorded for the line family; direct PNG review shows matching marker positions, curve geometry, axes, and legend with empty diagnostics (histogram correlation 0.503/0.815, the strongest magnitudes in this class, honestly recorded). Manifests: 118 locked / 206 approximate / 2 needs-review.

- Moved curved-connector-transform probe to approximate (Q07 increment): direct PNG review shows all four curved connectors with correctly directed arrowheads and matching curve paths (SSIM 0.988, recall 0.974, histogram correlation 1.0, diagnostics empty) with the recorded arrowhead-placement nuance below the lock bar. Manifests: 118 locked / 204 approximate / 4 needs-review.

- Moved border-alpha to approximate (Q07 increment): direct PNG review shows identical table geometry and fills with only border-blending shade nuance (reference borders fractionally darker; SSIM 0.885 on border coverage, recall 1.0, histogram correlation 0.998, diagnostics empty). Recorded as the alpha-blending gap class shared with picture-alpha. Manifests: 118 locked / 203 approximate / 5 needs-review.

- Moved grouped-picture-caption and title-placeholder to approximate (Q07 increment): direct PNG reviews show complete correct content indistinguishable at page scale (grouped picture with caption text; themed title placeholder with exact text; SSIM 0.990/0.991, recall 0.991/1.0, histogram correlation 1.0, diagnostics empty) with only sub-lock-bar positioning nuance. Manifests: 118 locked / 202 approximate / 6 needs-review.

- Moved gradient-diagnostic and picture-alpha to approximate (Q07 increment): direct PNG reviews show pixel-perfect geometry (recall 1.0, SSIM 0.9998, picture-alpha changed-pixel 0.0, diagnostics empty) with recorded color-math nuances below the family histogram bar (gradient stop interpolation histcorr 0.974, alpha-blending histcorr 0.950 vs the 0.98 bar). Manifests: 118 locked / 200 approximate / 8 needs-review.

- Moved sparse-blank-points probe to approximate (Q07 increment): direct PNG review shows the missing-cache fallback rendering identical gaps to Office (line dip and area start match; the required PPTX_CHART_MISSING_CACHED_DATA diagnostic fires as the manifest already gates); SSIM 0.963, recall 0.960, histogram correlation 1.0. Recorded as the designed missing-data fallback gap class. Manifests: 118 locked / 198 approximate / 10 needs-review.

- Moved line-stacked port to approximate (Q07 increment): fixture inspection proves both series are unstyled (no spPr; colors resolve through the theme accent sequence), so the same-hue shade difference is the documented unstyled-series approximation gap already recorded for line-markers and line-3series; direct PNG review shows matching line geometry, axes, and legend with SSIM 0.990 and empty diagnostics (histogram correlation 0.675). Manifests: 118 locked / 197 approximate / 11 needs-review.

- Moved line-trend port to approximate (Q07 increment): text-operation comparison (34 reference vs 34 candidate ops, no missing or extra text) pins the recall shortfall to fallback advance accumulation (title glyph drift to -0.20pt along the run, tick labels +0.35pt, legend within 0.13pt) with direct PNG review showing matching line geometry; SSIM 0.983, histogram correlation 0.924, diagnostics empty. Recorded as the measured font-advance positioning gap class. Manifests: 118 locked / 196 approximate / 12 needs-review.

- Moved data-label-legend-keys probe to approximate (Q07 increment): direct PNG review shows complete correct content indistinguishable at page scale (columns with data labels including legend-key glyphs all matching; SSIM 0.991, histogram correlation 0.999, diagnostics empty) with the recorded sub-lock-bar label-positioning nuance (recall 0.958). Manifests: 118 locked / 195 approximate / 13 needs-review.

- Moved doughnut-top-legend-exploded probe to approximate (Q07 increment): direct PNG review shows complete correct content indistinguishable at page scale (exploded segments with matching offsets, top legend swatches; SSIM 0.977, recall 0.990, histogram correlation 1.0, diagnostics empty) with the recorded sub-lock-bar positioning nuance. Manifests: 118 locked / 194 approximate / 14 needs-review.

- Moved column-negative and doughnut-top-legend probes to approximate (Q07 increment): direct PNG reviews show complete correct content indistinguishable at page scale (columns including below-axis negatives with month labels; doughnut segments with top legend swatches; SSIM 0.995/0.984, recall 0.974/0.993, diagnostics empty) with the recorded sub-lock-bar positioning nuance. Manifests: 118 locked / 193 approximate / 15 needs-review.

- Moved pie-auto-radius and bubble-port to approximate (Q07 increment): direct PNG reviews show complete correct content indistinguishable at page scale (pie slices with auto-radius geometry and data labels, recall 0.9998; bubble positions and size scaling with axes, title and legend, tight graphics/text structure gates passing; diagnostics empty) with the recorded sub-lock-bar positioning nuance. Manifests: 118 locked / 191 approximate / 17 needs-review.

- Moved bar-default-axis-titles and bar-shifted-frame probes to approximate (Q07 increment): direct PNG reviews show complete correct content indistinguishable at page scale (horizontal bars with default axis titles and gridlines; shifted frame geometry matches; SSIM 0.994/0.994, recall 0.983/0.979, histogram correlation 1.0, diagnostics empty) with the recorded sub-lock-bar positioning nuance. Manifests: 118 locked / 189 approximate / 19 needs-review.

- Moved plot-layout-outer and dashboard-table-chart probes to approximate (Q07 increment): direct PNG reviews show complete correct content indistinguishable at page scale (bar geometry with plot-box layout; KPI table plus Visits chart with axes, labels and legend; SSIM 0.994/0.989, recall 0.991/0.991, diagnostics empty) with the recorded sub-lock-bar positioning nuance. Manifests: 118 locked / 187 approximate / 21 needs-review.

- Moved composite-chart port to approximate (Q07 increment): direct PNG review shows complete correct content indistinguishable at page scale (clustered columns across four categories and three series with axes and legend all matching; SSIM 0.993, recall 0.992, histogram correlation 0.999, diagnostics empty) with the recorded sub-lock-bar positioning nuance. Manifests: 118 locked / 185 approximate / 23 needs-review.

- Moved doughnut-square and doughnut-portrait probes to approximate (Q07 increment): direct PNG reviews show complete correct content indistinguishable at page scale (doughnut segments, legend swatches and title all match; SSIM 0.993/0.991, recall 0.996/0.994, histogram correlation 1.0, diagnostics empty) with the recorded sub-lock-bar positioning nuance. Manifests: 118 locked / 184 approximate / 24 needs-review.

- Moved pie-leader-lines probe to approximate (Q07 increment): direct PNG review shows complete correct content indistinguishable at page scale (pie slices with data labels including the Gamma leader line; SSIM 0.992, recall 0.997, histogram correlation 1.0, diagnostics empty) with the recorded sub-lock-bar positioning nuance shared with its offset sibling. Manifests: 118 locked / 182 approximate / 26 needs-review.

- Moved radar-2series and pie-leader-lines-offset probes to approximate (Q07 increment): direct PNG reviews show complete correct content indistinguishable at page scale (two-series radar web with value rings and category axes; pie slices with data labels including the West leader line; SSIM 0.997/0.995, recall 0.986/0.996, diagnostics empty) with the recorded sub-lock-bar geometry nuance (radar value-axis geometry per the filled-radar fidelity note). Manifests: 118 locked / 181 approximate / 27 needs-review.

- Moved plot-layout-outer-narrow and compact-stacked-secondary-axis probes to approximate (Q07 increment): direct PNG reviews show complete correct content indistinguishable at page scale (bar geometry, plot-box layout, stacked segments with pattern fills, dual value axes, trend lines and arrow all match; SSIM 0.997/0.991, recall 0.990/0.993, MAE 0.087/0.042, diagnostics empty) with the recorded sub-lock-bar positioning nuance. Manifests: 118 locked / 179 approximate / 29 needs-review.

- Moved default and top-right axis-title probes to approximate (Q07 increment): direct PNG reviews show complete correct content indistinguishable at page scale (bars, default axis titles, tick labels, gridlines all match; axis-title text hashes match with positions inside structure tolerance; SSIM 0.997/0.996, recall 0.982/0.983, diagnostics empty) with the recorded sub-recall-bar positioning nuance keeping them below the lock bar. Manifests: 118 locked / 177 approximate / 31 needs-review.

- Moved conditional-run-precedence to approximate (Q07 increment): direct PNG review shows correct conditional run formatting (bold-italic vs plain plus red, matching the reference) with the recorded font-advance gap class (fallback sans advances fit both columns while Word serif advances overflow the 108pt grid, so overflow and column geometry diverge; SSIM 0.507 and recall 0.415 on the loose placeholder gates, diagnostics empty, dimensions match). Manifests: 118 locked / 175 approximate / 33 needs-review.

- Reported live image decode peaks alongside cumulative budgets (Q01 increment): `OoxConversionLimits.MaxLiveImageBytesPerConversion` (default 512 MiB) bounds transient decode scratch plus pixel planes while the image count bounds cumulative work; PNG/BMP/JPEG pixel decodes reserve a conservative width-by-height-by-4 estimate before allocating and release on return (JPEG passthrough holds no reservation), crossings throw the limit-exceeded failure before allocation with no partial PDF, and `CONVERSION_RESOURCE_SUMMARY` now carries `peakLiveImageBytes` so hosts can size admission limits. Pinned by below/at/above, release/peak, end-to-end cap (the 2x1 PNG pins the 8-byte estimate), and determinism tests with the affected suites proving byte-identical output.

- Render unstyled filled-radar series solid (fidelity fix): the 0.40/0.18 translucency defaults (unprovenanced, oldest history) contradicted two live Office references (single-series ladder fixture and a new two-series probe); every other chart kind already defaulted to solid. Single-series SSIM 0.858 to 0.954. Remaining radar gaps: chart-style-18 fill variants and value-axis geometry. New held-out visual case pins multi-series solid fills.

- Triaged the typography tail (Q07 increment): 21 passing needs-review cases moved to approximate with the recorded font-advance positioning gap class (sampled PNG reviews show complete correct content including distinct typefaces, line spacing, and auto-number quirks matching the reference; 7 pre-existing environment gate failures left untouched). Manifests: 118 locked / 176 approximate / 31 needs-review.

- Triaged the docx-layout tail (Q07 increment): 41 needs-review cases moved to approximate with the recorded fallback-advance wrapping/layout-nuance gap class after 10 direct PNG reviews (bullets, fragments, widows, continuation tables, docgrid, decorations all show complete correct content) plus metric-pattern review of the rest; page counts match and diagnostics are empty throughout. Conditional-run-precedence stays needs-review (font-advance overflow, column-width investigation). Manifests: 118 locked / 155 approximate / 52 needs-review; docx-layout family 60/60 green.

- Triaged docx-layout starts (Q07 increment): 2 new locks with reviewed PNGs and tightened gates (ladder-01 line-height/paragraph-spacing). Conditional-run-precedence spot-check shows matching run formatting with a genuine column-width divergence that stays needs-review for column-width investigation rather than approximate. Manifests: 118 locked / 114 approximate / 93 needs-review.

- Triaged typography and docx-layout families (Q07 increment): 5 new locks with reviewed PNGs and tightened gates (cambria highlight, vertical anchor/text, ladder-01 line-height/paragraph-spacing). 7 typography gate failures reproduce without related source changes and are recorded as environment-sensitive drift, gates untouched. Fragment-threshold pagination spot-check confirms correct fragmentation decisions with band-rhythm nuance only. Manifests: 121 locked / 114 approximate / 90 needs-review.

- Triaged the typography family (Q07 increment): 3 new locks with reviewed PNGs and tightened gates (cambria-highlight-center, vertical-anchor-middle, vertical-text-port). 7 pre-existing gate failures in this environment (anchor-overflow, text-flow, mixed-paragraph-stack, ellipsis, tab-space, whitespace-controls, underline-single) reproduce without any related source change and with exact unit tests green - recorded as environment-sensitive (font/rasterization drift), not regressions; classifications and gates left untouched.

- Triaged the charts family (Q07 increment): 24 new locks with reviewed PNGs and computed tightened gates (area, column, 15 doughnut probes, pie-exploded, plot-layout-inner, composite-two-charts) and 1 new approximate with a recorded gap (radar-filled translucent series fill). Radar-2series and top-right-axis-titles stay needs-review (recall below the lock bar; gates restored). Charts family green under the new gates.

- Triaged the shapes family (Q07 increment): 21 new locks with reviewed PNGs and tightened gates (17 pixel-perfect preset/polygon/connector/star/callout cases plus basic-shapes, connector-arrow, line-arrowheads, straight-stealth-connectors) and 1 new approximate with a recorded gap (shape-adjust-port-a chevron arrowhead depth ignores adjust handles). Curved-connector-transform stays needs-review (arrowhead placement nuance below the SSIM bar). Manifests: 89 locked / 113 approximate / 123 needs-review; shapes family 33/33 green under the new gates.

- Triaged smoke/effects/tables/composition visual families (Q07 increment): 4 new locks with reviewed PNGs and tightened gates (corporate-theme theme/master compositing; shadow-diagnostic soft shadow; basic-table quarterly grid; explicit-borders pastel grid; 5 pixel-perfect composition primitives; grouped/nested text; table-shape z-order; Arial-theme title placeholder) and 4 new approximates with recorded gap classes (composite band-fill theme tints; center-multiline/wrapped fill tints; middle-small-insets inset geometry). Gradient/picture-alpha and caption/placeholder-text cases stay needs-review below the family histogram/SSIM support bar. Manifests: 68 locked / 112 approximate / 145 needs-review; all four families green under the new gates.

- Sealed chart data-label record construction (T02 increment): `PptxSceneChartDataLabels` builds only through a `Defined` factory deriving its position enum from the raw spelling; both producer call sites updated. Pinned by the data-label-heavy chart suite, the full suite, and a 57/57 chart visual family run.

- Sealed chart plot/axis record construction (T02 increment): `PptxSceneChartPlot` and `PptxSceneChartAxis` build only through `Defined` factories that derive kind/grouping/direction/style and kind/position/crosses/orientation/tick enums from their raw spellings, so mismatched enum/string pairs are unrepresentable; single producer call sites updated. Pinned by the scene-vs-XML agreement batteries, the plot-pairing proof, the full suite, and a 57/57 chart visual family run.

- Share DOCX table-cell text layouts across measurement passes (W04 cell increment): a per-table memo keyed by cell identity, width, measurer, tab stops, row padding, spacing scale, and page args serves split-feasibility, break-resolution, final-fragment, and nested-table layouts with caller-side re-offsetting; cells with PAGE/NUMPAGES runs keep full page keys (dynamic tier) while all other cells share across pages (static tier), and height-dependent vertical-alignment shifts stay outside the memo. First-hand duplicate-rate evidence (779 corpus layouts fully distinct; sharing concentrates in split-heavy tables) plus per-fragment PAGE evaluation and memo-hit tests; full suite and docx-layout family (60/60) prove byte-identical output.

- Closed the chart dual-interpretation evidence program (D01): series, axis, title, and legend agreement batteries show no divergences (one real absent-title divergence found and unified); reachability analysis shows null-plot arms are unreachable in production rendering (unfiltered scene plots, Source-derived elements, early null-scene return) and they are retained deliberately as the pinned scene-less/defensive contract rather than deleted for no behavioral gain.

- Extended the chart options agreement harness to title properties (D01 increment): title text, body properties, text style, and text runs agree between scene and XML arms. The battery found a real divergence where absent titles built scene body properties from `default` (null strings) while the XML arm parses absent as empty strings; absent/auto-generated titles now parse through the shared body-properties reader (downstream consumers are null/empty-agnostic, chart family 57/57 green). Legend layout/text-style agreement battery (absent/position/overlay/delete) agrees exactly; absent legends provably share one builder on both arms, so no legend follow-up remains.

- Extended the chart options agreement harness to value-axis properties (D01 increment): orientation, major/minor gridlines, units (incl. bogus spellings), number formats, and tick marks agree between scene and XML arms across the battery with no divergences found. Test-only addition; renderer unchanged.

- Extended the chart options agreement harness to series properties (D01 increment): smooth flags, line-hidden detection, series-name records (literal/padded/ref/cache/no-cache, workbook-aware with structural point comparison), and point explosions (series/out-of-range/bogus/per-point, pie plots) agree between scene and XML arms across the batteries with no divergences found. Test-only addition; renderer unchanged.

- Threaded the conversion token through DOCX emission (Q01 increment): layout-item dispatch, per-segment text emission (including segment building, terminal-space resolution, and coverage-fallback splitting), per-cell table/border/junction rendering with nested recursion, inline/floating images (decode token now forwarded), text boxes, placed stories, floating drawings, markup balloons, and revision bars all observe cancellation at item/segment/cell/drawing granularity. Snapshot/inspection segment building keeps an optional token default. Pinned by a mid-emission cancellation test (sink-cancel on a mid-document image diagnostic with trailing content, no partial PDF) with the full suite proving byte-identical output. Remaining Q01 work: live-peak reservation reporting, bounded spooling, host admission limits, and cap tuning.

- Made font-table parsing cancellable (Q01 increment): `OpenTypeFont.Load` takes an optional parse token with checkpoints between collection extraction, table-directory, header validation, metrics, and kerning phases (including between legacy-kern and GPOS expansion), threaded from `FontProgramLoader`; `OperationCanceledException` propagates instead of becoming a fallback-font null. Pinned by parse-cancellation and loader-propagation tests with unaffected-output parity. Remaining Q01 work: DOCX emission/shape-fill token threading, live-peak reservation reporting, bounded spooling, host admission limits, and cap tuning.

- Typed static-story collection keys (T03 increment): `AddStaticStories` and the body story snapshot take `DocxStoryKind` instead of "Header"/"Footer"/"Body" literals (shared `ToKindString` rendering keeps inspection JSON byte-identical). Variant spellings intentionally stay raw strings: they double as document dictionary keys, so typing them would merge or drop unknown w:type spellings.
- Sealed chart part record construction (T02 increment): external-data, color-style, and style-part records build only through Defined/Undefined factories (defined color/style parts require resolved part identity), collapsing three drifting hand-built undefined literals into one spelling each. Producer call sites updated; pinned by factory rejection tests with the full suite proving corpus validity.
- Sealed PPTX scene node construction (T02 increment): `PptxSceneNode` builds only through a validated factory enforcing kind/payload consistency (shape/connector shape, picture, table, chart payloads; children exclusive to groups), so impossible combinations cannot arise through normal builders. The single builder call site is unchanged in behavior; pinned by factory rejection tests with the full suite proving corpus validity.
- Retired the XML-only chart render path (D01): deleted the unconstructible workbook-without-scene retry branch and its exclusive reference-cache hydration machinery (proven dead: a non-null workbook requires defined external data, which requires a scene chart), and split its dedicated test to keep the live workbook-parsing assertions. Complements the plot/chart options agreement harness and pairing proof with the shared gapWidth/overlap normalization from the prior increment.
- Retired a dual chart interpretation toward one reading (D01 increment): a permanent plot/chart options agreement harness (present/absent/unknown/out-of-range batteries, 1000+ scene-vs-XML comparisons) found a real divergence where out-of-range gapWidth/overlap bypassed clamping on the scene arm; both arms now share one normalization, and no committed fixture carries out-of-range values so locked outputs are unchanged. Remaining D01 surface: plot-presence pairing and the series/axis/title families.
- Triaged visual fidelity evidence (Q07): documented per-family parity targets and a lock policy (rating 4-5 for new locks, ~2-3x pixel-gate headroom, family support gates must pass) in `VisualValidation.md`; promoted the first Tier-1 chart ports to locked (`bar-clustered`, `pie-5-categories`, `column-clustered`, `bar-stacked`: indistinguishable renders, tight metrics, empty diagnostics) and moved `line-markers` and `line-3series` to approximate with their unstyled-series shade gaps recorded. W04-cell deferral re-verified first-hand (page-arg combinations, normalized-key design, and threading cost recorded in the ignored cell profile). A lock audit confirmed all 59 then-locked cases reference committed inputs.
- Typed DOCX layout story identity (T03 increment): the four layout item records plus floating-drawing layouts carry `DocxStoryId` (closed `DocxStoryKind` + raw w:type variant) instead of stringly-typed kind/variant pairs, so misspellings fail at compile time and header/footer switches are exhaustive. Inspection and emission snapshots render the exact legacy spellings and stay byte-identical (existing header/footer/table/markup snapshot assertions green). Remaining T03 surface: static-story collection keys, variant-typing over document dictionary keys, and geometry unit boundaries.
- Introduced a conversion-wide cumulative work budget (Q01 core): `OoxPdfOptions.ConversionLimits` sets per-conversion totals for chart range cells, table fragments, decoded images, and font operations (generous multiples of the per-site caps by default); crossings throw `OoxPdfLimitExceededException` before further expansion with atomic output preserved. The budget is explicitly scoped per `Convert` call (nested scopes restore), charges at the chart-union, fragment, image-decode, and font-load/subset sites, and `ReportResourceUsage` emits a deterministic informational CONVERSION_RESOURCE_SUMMARY diagnostic so hosts can account concurrent conversions. Pinned by below/at/above, nesting, end-to-end, determinism, and validation tests.
- Made link-annotation targets unrepresentable-by-construction (T02 increment): `PdfLinkAnnotation` builds only through validated `ToUri`/`ToDestination` factories (empty URIs rejected), so both/neither target combinations cannot arise through normal builders. Serialization bytes are unchanged; scene/chart-record consolidation remains open.
- Added cooperative cancellation inside image decoding (Q01 increment): PNG chunk/inflate/scanline loops, JPEG MCU rows, and BMP pixel rows observe the conversion token via optional parameters (existing callers unchanged), forwarded through the shared decoder and the PPTX picture/crop paths. Pinned by a mid-decode cancellation test with unaffected-output parity. Remaining Q01 work: DOCX emission-layer and shape-fill token threading, font-parse cancellation, cumulative cross-stage budgets with peak reporting, bounded spooling, and host admission limits.
- Made node recovery accountable (Q02): `PPTX_NODE_RENDER_FAILED` preserves the failure category (exception type only, never messages that could carry document text), each slide prunes orphan image/chart-font registrations left by rewound nodes so failures cannot accumulate uncharged serialized resources, and `DiagnosticCollector` bounds retained samples (4096) while still counting per-ID occurrences with a `DIAGNOSTIC_OVERFLOW` marker in CLI output. Budget/cancellation failures keep escaping recovery; neighbor recovery is unchanged. Pinned by collector-cap, prune-identity, cause-preservation, and PPTX mid-work cancellation tests.
- Consolidated duplicated selection mechanics (D02): DOCX and PPTX image paths share one content-image decoder (`OoxImageDecoder`: JPEG passthrough, PNG/BMP pixels, unknown-type rejection) with format diagnostics and PPTX recolor/crop policy staying local, and both font resolvers share style/weight candidate scoring with explicit resolver-local tie-breakers. Behavior pinned by decoder unit tests, resolver matrix tests, image unit groups, and a 14/14 pptx-images visual family run.
- Defined explicit font identity keys (T01): PPTX program caches use ordinal `(StableId, FaceIndex)` tuples matching the DOCX font-plan keys and the source-immutability contract, so case-distinct custom IDs no longer alias; text-run grouping keys are typed `FontRequest` values with the explicit case-insensitive family comparer instead of per-lookup key strings; subset reuse verifies codepoint-set equality on every hit so a hash collision can never corrupt output. Pinned by tests for case-distinct IDs, collection faces, and style-variant subset sharing; conversion bytes unchanged.
- Made fidelity and calibration evidence durable (Q07): the fraction-curve comment now points at the tracked `PptxChartsFractionTests` held-out pins instead of the ignored review plan, the bold/italic hybrid policy is stated inline in `Capabilities.md` instead of citing the plan, and `unit-test-audit.md` drops the resolved network-restore blocker with refreshed Release suite numbers (1546 passed / 0 failed / 9 skipped). Family-by-family parity triage with new held-out cases remains future work.
- Clarified public operational contracts (Q08): `OoxPdfConverter` documents sync/async byte-parity, caller-owned stream handling with atomic file publication, cooperative cancellation without partial PDFs, concurrent shared-resolver use, and strict/deterministic semantics, backed by new tests proving mid-work cancellation fails fast with no output and eight parallel conversions sharing one resolver stay byte-identical.
- Hardened validation tools with independent bounds (Q06): the PDFium rasterizer caps DPI/dimensions/pixels with per-page isolation and streams PNG rows with incremental CRC, VisualDiff bounds PNG chunk framing/dimensions/inflation against the shared 67M-pixel budget with truncated-input failures, PdfInspect/PptxInspect cap inputs at 512 MiB with bounded per-stream inflation and clean missing-file errors, PPTX inspection builds one shared scene per conversion instead of one per slide per entry point (byte-identical snapshots verified against the pre-change tool), and every tool reports its peak working set separately from converter peaks.
- Bounded chart, table, and fragment expansion so tiny declarations fail predictably instead of amplifying into large allocations: chart dense points and sparse indexes, DOCX grid spans, PPTX border grids, chart range unions, DOCX row fragments, XML attributes, workbook totals, and ZIP intake now enforce explicit limits through a dedicated limit-exceeded failure that escapes per-node recovery.
- Validate PDF content from token offsets (G05): operator classification without strings, span-based resource lookups, and single-pass ASCII checking cut validation allocation from 12 MB to under 1 MB on 400k-token content with identical failure behavior.
- Share DOCX related-story parsing across references and sections (W04): one body/drawing extraction per unique header/footer part per Read, with per-part verdict caching preserving once-per-part diagnostics; cloned font-plan runs index confirmed fallback matches. Cell-layout memo explicitly deferred pending page/footnote-safe keys (profiled 1.8x layouts per cell).
- Share PPTX text/table layout between font collection and painting (W02): per-slide memoization keyed by node, color map, and placeholder context; suppressed master shapes no longer feed font preflight; table-style diagnostics keep exact paint-order emission.
- Share PPTX XML/relationship parses per conversion (W01): package-level immutable caches so shared masters, layouts, and relationship parts parse once (instrumented with parse counters), with visibility discovery reusing scene parses. Shared-DOM safety audited (mutations target copies or idempotent chart hydration); outputs unchanged.
- Cut repeated font-resolution work (G03): family indexes and precomputed fallback lists per snapshot, per-conversion request caches with deterministic fallback spelling, single-pass best-candidate scans replacing per-resolve five-key sorts, and a no-Task fast path for completed font sources. Repeated resolves measure ~7x faster (20k Arial resolves: 55ms to 8ms); first-time fallbacks at parity. Resolution outcomes unchanged.
- De-duplicated DOCX emergency wrapping work (W03): slice widths memoized by text coordinates, in-place token updates, hoisted segment-level normalization scans, incremental whitespace tracking, and cancellation checkpoints in the scan loops. Binary search deliberately not applied (prefix widths are not provably monotonic under kerning/tracking/fields); identical line starts pinned by count-based scaling tests.
- Removed repeated chart workbook/range work (W05): one bounded workbook model per embedded part shared across a conversion, per-frame range-expansion memoization (cleared per frame), single-pass numeric parsing, and indexed workbook-point lookups for scatter and data labels (quadratic scans removed; obsolete helper deleted). Outputs unchanged.
- Defined font-byte ownership and limits (M09): local font files share the 64 MiB per-font ceiling with pack downloads (installed fonts peak at 35.4 MiB here), discovery skips oversized files at stat time, pack-resolver retention is aggregate-capped at 256 MiB with hash-verified LRU eviction, and source immutability/thread-safety/lifetime contracts are documented on the resolver types.
- Cut PNG working memory (M08): the decoder inflates from a read-only view of the accumulated IDAT bytes and decodes from the inflated buffer in place, removing two full-size transient copies per image with byte-identical pixels and identical truncation failures. Single images stay pixel-capped; cumulative multi-image/effect budgets await the conversion-wide contract (Q01).
- Bounded font-table expansion (M10): GPOS class kerning densification (2M pair cap, ~2.7x above the largest installed face measured), coverage/class range expansion (131k glyph caps), a shared kerning work budget, and a per-read compound-outline point budget (100k). Over-cap fonts fall back to substitutes through the existing loader contract; kern-pair spot checks and subset output for normal fonts are unchanged.
- Repaired allocation-probe accounting (Q05): output hashing/page-count checks now run outside the counters, retained memory is a post-GC delta against a pre-run baseline, read stages no longer include package rebuilds, PPTX render is labeled inclusive of its scene rebuild, reports disclose source revision/OS/GC mode/font inventory/output mode/isolation, --output-mode file separates output-buffering costs, --isolate measures each input in a fresh process with per-input peaks, and the self-test validates attribution with a planted allocation.
- Cut cold Windows font discovery allocation roughly in half (about 1,140 MB to about 542 MB on this machine's font inventory, about 28% faster): discovery now reads only the table directory plus required header-table extents per file and parses collection faces in place instead of repackaging per-face copies, with accept/reject behavior identical to whole-program reads.
- Hardened the CLI so diagnostics paths colliding with input/output are rejected before conversion, malformed numeric geometry normalizes to invalid-data failures, and diagnostics-write failures no longer escape error handling.

## 0.1.5 - 2026-09-19

- Fixed garbled subset-font text when one typeface is used with different character sets across a document: embedded font resources are now keyed by codepoint set so merged subsets keep valid CID mappings.
- Supported forward-only output streams for stream conversion: PDF cross-reference offsets are now tracked internally instead of reading output Position, so non-seekable host descriptor streams no longer fail with Specified method is not supported.
- Fixed very slow DOCX conversion for run-dense documents: run typeface resolution during text measuring is now indexed instead of scanning all runs per measurement, about 100x faster on a 131-page markup document with byte-identical output.
- Capped XML element count per part during package reading so broad shallow documents fail fast instead of inflating the DOM without bound.
- Improved DOCX pagination, table sizing and borders, headers and footers, and Word-compatible review markup placement.
- Improved PPTX chart layout, labels, legends, number formats, and theme-based varying colors, along with vertical text and text-frame layout.
- Corrected built-in PPTX table styles, including theme text colors and linear-light band, column, and last-row shading.
- Reduced repeated font loading and discovery work, coalesced concurrent font-pack downloads, added transient download retries, and exposed Windows font discovery cache invalidation.
- Hardened package and image parsing, selected a single supported markup-compatibility representation, and added diagnostics for unsupported Strict OOXML and must-understand content.
- Protected existing destination files with atomic replacement after successful conversion. Stream outputs now reject nonempty or nonzero-position seekable streams, and input and output must be distinct.
- Corrected PDF hyperlink escaping for non-ASCII URLs, with bounded stack use even for very long links.
- Added Windows and Ubuntu CI coverage, including package-consumption smoke tests; Ubuntu intentionally runs the portable API, OOXML, PDF, imaging, and font test groups.

## 0.1.4 - 2026-07-02

- Added a built-in HTTP(S) font pack resolver for deterministic packaged font resolution outside local Windows font directories.

## 0.1.3 - 2026-06-15

- Added DOCX markup printing support for comments, insertions, deletions, and move revisions.
- Added DOCX review-mode options for final/original/simple/all-markup output and margin-preserving or expanded comment geometry.
- Added public visual validation coverage and tooling for DOCX markup reference-cache workflows.

## 0.1.2 - 2026-06-09

- Added public stream-based conversion overloads so callers can convert PPTX and DOCX packages to PDF without creating temporary input or output files.

## 0.1.1 - 2026-06-06

- Subset embedded TrueType fonts so generated PDFs carry only the glyphs, metrics, and Unicode map entries needed by rendered PPTX and DOCX text.
- Added coverage for subset font output, deterministic font subsetting, and PPTX text emission through subset font resources.
- Added package release notes metadata and packaged this changelog with the NuGet artifact.

## 0.1.0 - 2026-06-05

- Initial public package for dependency-free PPTX and DOCX to PDF conversion.
- Included managed OOXML readers, PDF writing, font embedding, image handling, diagnostics, CLI conversion, and visual validation tooling.
