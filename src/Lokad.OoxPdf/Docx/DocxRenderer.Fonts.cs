using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static DocxFontResources PrepareFontResourcesCore(DocxDocument document, IFontResolver fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink, CancellationToken cancellationToken, bool includeReviewAutofitMarks, int? pageCount,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxFontPlan plan = DocxFontPlan.Create(document, fontResolver, cancellationToken, includeReviewAutofitMarks);
        var resources = new List<PdfFontResource>();
        var runResources = new Dictionary<DocxTextRun, DocxRunFontResource>();
        var reportedUnembeddableFaces = new HashSet<(string StableId, int FaceIndex)>();
        plan = SubstituteUnembeddableFonts(plan, fontCache, diagnosticSink, reportedUnembeddableFaces, cancellationToken);
        // Five-paragraph mixed terminals previously retained no marks. Give their
        // newly retained marks separate subsets and names, so a rejected fitting
        // path cannot change body subsets or renumber surviving fallback fonts.
        var separateMarks = new HashSet<DocxTextRun>(ReferenceEqualityComparer.Instance);
        if (includeReviewAutofitMarks)
        {
            foreach (DocxRelatedStory story in document.RelatedStories)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (story.Kind != DocxRelatedStoryKind.Comment ||
                    !DocxBalloonParagraphPolicy.HasMixedFiveParagraphTerminal(story.BodyElements)) { continue; }
                foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(story))
                {
                    if (paragraph.ParagraphMarkRun is { } mark) { separateMarks.Add(mark); }
                }
            }
        }
        DocxFontPlan bodyPlan = separateMarks.Count == 0 ? plan : new(plan.Runs.Where(run => !separateMarks.Contains(run.Run)).ToArray());
        PrepareResolvedRunFontResources(bodyPlan, resources, runResources, fontCache, cancellationToken);
        DocxRunFontResource? fallback = PrepareFallbackFontResource(bodyPlan, fontResolver, resources, runResources, fontCache, diagnosticSink, reportedUnembeddableFaces, cancellationToken);
        IReadOnlyDictionary<DocxTextRun, IReadOnlyList<DocxFallbackFontEntry>> fallbackChains = PreparePerCharacterFallbackResources(bodyPlan, fontResolver, fallback?.Resolution, resources, runResources, fontCache, diagnosticSink, reportedUnembeddableFaces, cancellationToken);
        if (separateMarks.Count != 0)
        {
            var markResources = new List<PdfFontResource>();
            PrepareResolvedRunFontResources(new(plan.Runs.Where(run => separateMarks.Contains(run.Run)).ToArray()),
                markResources, runResources, fontCache, cancellationToken, resourcePrefix: "M");
            resources.AddRange(markResources);
        }
        IDocxTextMeasurer? innerMeasurer = plan.Runs.Any(run => LoadFont(run.Resolution, fontCache, cancellationToken) is not null) || fallback is not null
            ? new DocxFontPlanTextMeasurer(plan, fallback?.Resolution, cancellationToken, fontResolver, fontCache, resolveSourceAdvances: true)
            : null;
        ReportMissingGlyphs(plan, runResources, fallbackChains, fontCache, diagnosticSink, cancellationToken);
        Dictionary<DocxTextRun, PdfFallbackFontResource> fallbackFaces = PrepareMissingFontFallback(bodyPlan, runResources, diagnosticSink, cancellationToken);
        List<PdfFallbackFontResource> fallbackFontResources = fallbackFaces.Values
            .DistinctBy(face => face.Font.ResourceKey, StringComparer.Ordinal)
            .OrderBy(face => face.ResourceName, StringComparer.Ordinal)
            .ToList();
        IDocxTextMeasurer? measurer = innerMeasurer;
        if (fallbackFaces.Count != 0)
        {
            measurer = new MissingFontRoutingMeasurer(innerMeasurer, new DocxFallbackTextMeasurer(), fallbackFaces);
        }

        ReportComplexScriptApproximation(plan, diagnosticSink, cancellationToken);

        var deferredPageCountEquations = new List<DocxEquation>();
        var mathEquations = PrepareMathEquations(document, fontResolver, fontCache, resources, cancellationToken, pageCount, deferredPageCountEquations);
        return new DocxFontResources(plan, measurer, resources, runResources, fallback, fallbackChains, fallbackFaces, fallbackFontResources)
        {
            Charts = new DocxChartResources(new PresentationFontResolver(fontResolver)),
            MathEquations = mathEquations,
            DeferredPageCountEquations = deferredPageCountEquations
        };
    }

    // CFF/OpenType-CFF fonts have valid metrics but no TrueType outlines, so the
    // PDF subsetter cannot embed them. Remap those runs to unresolved before any
    // measurement or emission so both paths share the document fallback face (F03:
    // measured and emitted glyphs use the same face) and report each substituted
    // typeface once through the diagnostic sink.
    private static DocxFontPlan SubstituteUnembeddableFonts(
        DocxFontPlan plan,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<(string StableId, int FaceIndex)> reported,
        CancellationToken cancellationToken)
    {
        List<DocxResolvedRunTypeface>? remapped = null;
        for (int i = 0; i < plan.Runs.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxResolvedRunTypeface run = plan.Runs[i];
            if (run.Resolution is not { } resolution ||
                LoadFont(resolution, fontCache, cancellationToken) is not { } font ||
                font.HasTrueTypeOutlines)
            {
                if (remapped is not null)
                {
                    remapped.Add(run);
                }

                continue;
            }

            if (remapped is null)
            {
                remapped = new List<DocxResolvedRunTypeface>(plan.Runs.Count);
                for (int j = 0; j < i; j++)
                {
                    remapped.Add(plan.Runs[j]);
                }
            }

            ReportUnembeddableFace(resolution, "substituted with the document fallback typeface", "Document fallback typeface", diagnosticSink, reported);

            remapped.Add(run with { Resolution = null });
        }

        return remapped is null ? plan : new DocxFontPlan(remapped);
    }

    // Reports an unembeddable (CFF/OpenType-CFF) face once per conversion.
    // Callers describe what took its place.
    private static void ReportUnembeddableFace(
        FontFaceResolution resolution,
        string action,
        string fallback,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<(string StableId, int FaceIndex)> reported)
    {
        if (!reported.Add((resolution.Source.StableId, resolution.FontFaceIndex)))
        {
            return;
        }

        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            "FONT_UNSUPPORTED_OUTLINES",
            OoxPdfSeverity.Warning,
            "Font '" + resolution.FamilyName + "' has no embeddable TrueType outlines (CFF/OpenType-CFF) and was " + action + ".",
            PartName: null,
            SlideIndex: null,
            PageIndex: null,
            Feature: resolution.FamilyName,
            Fallback: fallback));
    }

    private sealed class ScaledDocxTextMeasurer(IDocxTextMeasurer inner, double textScale, double lineMetricScale) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxTypographicMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            return inner.MeasureText(run, text, fontSize) * textScale;
        }

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxLineMetricsProvider lineMetrics
                ? lineMetrics.MeasureSingleLineHeight(run, fontSize) * lineMetricScale
                : fontSize * 1.2d * lineMetricScale;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxStaticTextMetricsProvider staticMetrics
                ? staticMetrics.MeasureWindowsAscender(run, fontSize) * lineMetricScale
                : fontSize * lineMetricScale;
        }

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxStaticTextMetricsProvider staticMetrics
                ? staticMetrics.MeasureWindowsDescender(run, fontSize) * lineMetricScale
                : fontSize * 0.2d * lineMetricScale;
        }

        public bool UseTypographicMetrics(DocxTextRun? run)
        {
            return inner is IDocxTypographicMetricsProvider typographic && typographic.UseTypographicMetrics(run);
        }

        public double MeasureTypographicBaselineInset(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxTypographicMetricsProvider typographic
                ? typographic.MeasureTypographicBaselineInset(run, fontSize) : 0d;
        }

        public double MeasureTypographicLineHeight(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxTypographicMetricsProvider typographic
                ? typographic.MeasureTypographicLineHeight(run, fontSize) * lineMetricScale
                : 0d;
        }
    }

    // RV01: codepoints no usable face covers are substituted with question mark at
    // emission; report them once per family (union across runs, sorted, capped) so the
    // substitution is diagnosed, never silent.
    private static void ReportMissingGlyphs(
        DocxFontPlan plan,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        IReadOnlyDictionary<DocxTextRun, IReadOnlyList<DocxFallbackFontEntry>> fallbackChains,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        var missingByFamily = new Dictionary<string, SortedSet<int>>(StringComparer.Ordinal);
        foreach (DocxResolvedRunTypeface resolved in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (resolved.Run.EffectiveProperties.Hidden || string.IsNullOrEmpty(resolved.Run.Text) ||
                !runResources.TryGetValue(resolved.Run, out DocxRunFontResource? resource))
            {
                continue;
            }

            // Runs with a per-character chain are covered face by face; only runes
            // missed by every candidate count. Layout constructs never reach glyphs.
            IReadOnlyList<OpenTypeFont> coverageFonts = fallbackChains.TryGetValue(resolved.Run, out IReadOnlyList<DocxFallbackFontEntry>? chain)
                ? chain.Select(entry => entry.Font).ToArray()
                : (LoadFont(resource.Resolution, fontCache, cancellationToken) is { } primaryFont ? new[] { primaryFont } : Array.Empty<OpenTypeFont>());
            if (coverageFonts.Count == 0)
            {
                continue;
            }

            foreach (Rune rune in resolved.Run.Text.EnumerateRunes())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!PdfFallbackFont.IsNonRenderedControl(rune) && coverageFonts.All(candidate => candidate.MapCodePoint(rune.Value) == 0))
                {
                    if (!missingByFamily.TryGetValue(resource.Resolution.FamilyName, out SortedSet<int>? codepoints))
                    {
                        codepoints = new SortedSet<int>();
                        missingByFamily[resource.Resolution.FamilyName] = codepoints;
                    }

                    codepoints.Add(rune.Value);
                }
            }
        }

        const int maxListedCodepoints = 12;
        foreach ((string family, SortedSet<int> codepoints) in missingByFamily)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string listed = string.Join(", ", codepoints.Take(maxListedCodepoints).Select(code => "U+" + code.ToString("X4")));
            if (codepoints.Count > maxListedCodepoints)
            {
                listed += ", and " + (codepoints.Count - maxListedCodepoints).ToString(CultureInfo.InvariantCulture) + " more";
            }

            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "FONT_MISSING_GLYPHS",
                OoxPdfSeverity.Warning,
                codepoints.Count.ToString(CultureInfo.InvariantCulture) + " characters of " + family + " have no glyph in any usable face (" + listed + ") and show as question mark.",
                PartName: null,
                SlideIndex: null,
                PageIndex: null,
                Feature: family,
                Fallback: "Question-mark substitution"));
        }
    }

    // RV02: complex-script runs render without shaping, so report each behavior
    // family once per conversion instead of misrendering silently. The font plan
    // already covers every rendered story, so scanning it needs no extra traversal.
    private static void ReportComplexScriptApproximation(
        DocxFontPlan plan,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        OoxComplexScriptKind seen = OoxComplexScriptKind.None;
        foreach (DocxResolvedRunTypeface resolved in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            seen |= OoxComplexScript.DetectNeeds(resolved.Run.Text);
        }

        foreach (OoxComplexScriptKind kind in new[]
        {
            OoxComplexScriptKind.Joining,
            OoxComplexScriptKind.Reordering,
            OoxComplexScriptKind.Bidirectional,
            OoxComplexScriptKind.CombiningMark
        })
        {
            if (seen.HasFlag(kind))
            {
                diagnosticSink(OoxComplexScript.CreateApproximationDiagnostic(kind));
            }
        }
    }

    // RV01: runs with text but no usable embedded resource render with the diagnosed
    // built-in fallback faces instead of vanishing. Each affected family is reported once;
    // resource names follow fixed face order (FF1 regular, FF2 bold, FF3 oblique,
    // FF4 bold-oblique) so output stays deterministic.
    private static Dictionary<DocxTextRun, PdfFallbackFontResource> PrepareMissingFontFallback(
        DocxFontPlan plan,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        var faces = new Dictionary<DocxTextRun, PdfFallbackFontResource>();
        var reportedFamilies = new HashSet<string>(StringComparer.Ordinal);
        foreach (DocxResolvedRunTypeface resolved in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(resolved.Run.Text) || runResources.ContainsKey(resolved.Run))
            {
                continue;
            }

            DocxEffectiveRunProperties effective = resolved.Run.EffectiveProperties;
            PdfFallbackFont face = PdfFallbackFont.ForStyle(effective.Bold, effective.Italic);
            string name = PdfFallbackFont.ResourceNameFor(effective.Bold, effective.Italic);
            faces[resolved.Run] = new PdfFallbackFontResource(name, face);

            string family = resolved.RequestedFamily ?? resolved.ResolvedFamily ?? "unknown";
            if (reportedFamilies.Add(family))
            {
                diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                    "FONT_NO_USABLE_FACE",
                    OoxPdfSeverity.Warning,
                    "No usable embeddable font for " + family + "; text renders with the built-in fallback typeface. Characters outside WinAnsi appear as question mark.",
                    PartName: null,
                    SlideIndex: null,
                    PageIndex: null,
                    Feature: family,
                    Fallback: "Built-in fallback typeface"));
            }
        }

        return faces;
    }

    private static void PrepareResolvedRunFontResources(
        DocxFontPlan plan,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken,
        string resourcePrefix = "F")
    {
        var resolvedRuns = new List<(DocxResolvedRunTypeface Run, FontFaceResolution Resolution)>();
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (run.Resolution is not { } resolution || LoadFont(resolution, fontCache, cancellationToken) is null)
            {
                continue;
            }
            resolvedRuns.Add((run, resolution));
        }

        foreach (IGrouping<(string StableId, int FaceIndex), (DocxResolvedRunTypeface Run, FontFaceResolution Resolution)> group in resolvedRuns.GroupBy(item => (item.Resolution.Source.StableId, item.Resolution.FontFaceIndex)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            FontFaceResolution resolution = group.First().Resolution;
            OpenTypeFont? font = LoadFont(resolution, fontCache, cancellationToken);
            if (font is null)
            {
                continue;
            }

            IReadOnlyList<int> glyphs = CollectRunGlyphs(group.Select(item => item.Run), font, cancellationToken);
            if (glyphs.Count == 0)
            {
                continue;
            }

            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs, cancellationToken);
            string name = resourcePrefix + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            var runResource = new DocxRunFontResource(name, embedded, resolution);
            resources.Add(new PdfFontResource(name, embedded));
            foreach (DocxResolvedRunTypeface run in group.Select(item => item.Run))
            {
                cancellationToken.ThrowIfCancellationRequested();
                runResources[run.Run] = runResource;
            }
        }
    }

    private static DocxRunFontResource? PrepareFallbackFontResource(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<(string StableId, int FaceIndex)> reported,
        CancellationToken cancellationToken)
    {
        FontFaceResolution resolution = ResolveDocumentBaseFont(plan, fontResolver, fontCache, cancellationToken);
        OpenTypeFont? font = LoadFont(resolution, fontCache, cancellationToken);
        if (font is null)
        {
            return null;
        }

        if (!font.HasTrueTypeOutlines)
        {
            ReportUnembeddableFace(resolution, "skipped as the document fallback typeface", "Missing-font path", diagnosticSink, reported);
            return null;
        }

        DocxResolvedRunTypeface[] fallbackRuns = plan.Runs
            .Where(run => !runResources.ContainsKey(run.Run))
            .ToArray();
        if (fallbackRuns.Length == 0)
        {
            return null;
        }

        IReadOnlyList<int> glyphs = CollectRunGlyphs(fallbackRuns, font, cancellationToken);
        if (glyphs.Count == 0)
        {
            return null;
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs, cancellationToken);
        string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
        var runResource = new DocxRunFontResource(name, embedded, resolution);
        resources.Add(new PdfFontResource(name, embedded));
        foreach (DocxResolvedRunTypeface run in fallbackRuns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            runResources[run.Run] = runResource;
        }

        return runResource;
    }

    private static IReadOnlyDictionary<DocxTextRun, IReadOnlyList<DocxFallbackFontEntry>> PreparePerCharacterFallbackResources(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        FontFaceResolution? documentFallback,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<(string StableId, int FaceIndex)> reported,
        CancellationToken cancellationToken)
    {
        var chains = new Dictionary<DocxTextRun, IReadOnlyList<DocxFallbackFontEntry>>();
        var fallbackCodepoints = new Dictionary<(string StableId, int FaceIndex), HashSet<int>>();
        var fallbackFonts = new Dictionary<(string StableId, int FaceIndex), OpenTypeFont>();
        var fallbackResolutions = new Dictionary<(string StableId, int FaceIndex), FontFaceResolution>();
        var runFallbacks = new Dictionary<DocxTextRun, (OpenTypeFont PrimaryFont, IReadOnlyList<OpenTypeFont> Fonts, IReadOnlyList<FontFaceResolution> Resolutions, IReadOnlyList<FontCoverageSpan> Spans)>();
        foreach (DocxResolvedRunTypeface resolved in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(resolved.Run.Text))
            {
                continue;
            }

            if (!runResources.TryGetValue(resolved.Run, out DocxRunFontResource? primaryResource))
            {
                continue;
            }

            OpenTypeFont? primaryFont = LoadFont(primaryResource.Resolution, fontCache, cancellationToken);
            if (primaryFont is null)
            {
                continue;
            }

            // Primary-covered runs need no fallback candidates; skipping their
            // program loads leaves split outcomes unchanged (G04).
            if (FontCoverageFallback.IsFullyCovered(resolved.Run.Text, primaryFont, cancellationToken))
            {
                continue;
            }

            var fonts = new List<OpenTypeFont>();
            var resolutions = new List<FontFaceResolution>();
            foreach (FontFaceResolution candidate in DocxFontFallbackRules.ResolveCandidateResolutions(fontResolver, resolved.Run, resolved.ResolvedFamily, primaryResource.Resolution, documentFallback))
            {
                cancellationToken.ThrowIfCancellationRequested();
                OpenTypeFont? candidateFont = fonts.Count == 0 ? primaryFont : LoadFont(candidate, fontCache, cancellationToken);
                if (candidateFont is null || fonts.Contains(candidateFont))
                {
                    continue;
                }

                if (fonts.Count != 0 && !candidateFont.HasTrueTypeOutlines)
                {
                    ReportUnembeddableFace(candidate, "skipped for per-character fallback", "Next covering typeface", diagnosticSink, reported);
                    continue;
                }

                fonts.Add(candidateFont);
                resolutions.Add(candidate);
            }

            if (fonts.Count <= 1)
            {
                continue;
            }

            IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage(resolved.Run.Text, fonts, cancellationToken);
            bool needsFallback = false;
            foreach (FontCoverageSpan span in spans)
            {
                if (span.FontIndex > 0)
                {
                    needsFallback = true;
                    break;
                }
            }

            if (!needsFallback)
            {
                continue;
            }

            runFallbacks[resolved.Run] = (primaryFont, fonts, resolutions, spans);
            for (int i = 0; i < spans.Count; i++)
            {
                FontCoverageSpan span = spans[i];
                if (span.FontIndex <= 0)
                {
                    continue;
                }

                FontFaceResolution spanResolution = resolutions[span.FontIndex];
                var key = (spanResolution.Source.StableId, spanResolution.FontFaceIndex);
                if (!fallbackCodepoints.TryGetValue(key, out HashSet<int>? codepoints))
                {
                    codepoints = new HashSet<int>();
                    fallbackCodepoints[key] = codepoints;
                    fallbackFonts[key] = fonts[span.FontIndex];
                    fallbackResolutions[key] = spanResolution;
                }

                foreach (Rune rune in resolved.Run.Text.Substring(span.Start, span.Length).EnumerateRunes())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    codepoints.Add(rune.Value);
                }
            }
        }

        if (runFallbacks.Count == 0)
        {
            return chains;
        }

        var fallbackResources = new Dictionary<(string StableId, int FaceIndex), DocxRunFontResource>();
        foreach (var key in fallbackCodepoints.Keys)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var glyphs = new HashSet<int>(fallbackCodepoints[key]);
            foreach (Rune rune in " 0123456789".EnumerateRunes())
            {
                glyphs.Add(rune.Value);
            }

            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(fallbackFonts[key], glyphs, cancellationToken);
            string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            var resource = new DocxRunFontResource(name, embedded, fallbackResolutions[key]);
            resources.Add(new PdfFontResource(name, embedded));
            fallbackResources[key] = resource;
        }

        foreach (DocxTextRun run in runFallbacks.Keys)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var entry = runFallbacks[run];
            var chain = new List<DocxFallbackFontEntry> { new DocxFallbackFontEntry(entry.PrimaryFont, runResources[run]) };
            for (int i = 1; i < entry.Fonts.Count; i++)
            {
                bool used = false;
                foreach (FontCoverageSpan span in entry.Spans)
                {
                    if (span.FontIndex == i)
                    {
                        used = true;
                        break;
                    }
                }

                if (!used)
                {
                    continue;
                }

                FontFaceResolution spanResolution = entry.Resolutions[i];
                var key = (spanResolution.Source.StableId, spanResolution.FontFaceIndex);
                chain.Add(new DocxFallbackFontEntry(entry.Fonts[i], fallbackResources[key]));
            }

            chains[run] = chain;
        }

        return chains;
    }
    // RV01: question mark joins the subset when a covered run carries codepoints
    // the face cannot map, so emission-time substitution stays extractable. A face
    // mapping question mark itself to .notdef is pathological and left as-is.
    private static IReadOnlyList<int> CollectRunGlyphs(IEnumerable<DocxResolvedRunTypeface> runs, OpenTypeFont? coverageFont, CancellationToken cancellationToken)
    {
        var glyphs = new HashSet<int>();
        bool needsQuestionMark = false;
        foreach (DocxResolvedRunTypeface run in runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (Rune rune in run.Run.Text.EnumerateRunes())
            {
                cancellationToken.ThrowIfCancellationRequested();
                glyphs.Add(rune.Value);
                if (!needsQuestionMark && coverageFont is not null && !PdfFallbackFont.IsNonRenderedControl(rune) && coverageFont.MapCodePoint(rune.Value) == 0)
                {
                    needsQuestionMark = true;
                }
            }
        }

        if (needsQuestionMark && coverageFont is not null && coverageFont.MapCodePoint(0x3F) != 0)
        {
            glyphs.Add(0x3F);
        }

        foreach (Rune rune in " 0123456789".EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            glyphs.Add(rune.Value);
        }

        return glyphs.ToArray();
    }

    private static FontFaceResolution ResolveDocumentBaseFont(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken)
    {
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (run.Resolution is { } fontResolution && LoadFont(fontResolution, fontCache, cancellationToken) is not null)
            {
                return fontResolution;
            }
        }

        return DocxFontFallbackRules.ResolveDefaultDocumentTypeface(fontResolver, false, false);
    }

    private static OpenTypeFont? LoadFont(
        FontFaceResolution? resolution,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (resolution is null)
        {
            return null;
        }

        var key = (resolution.Source.StableId, resolution.FontFaceIndex);
        if (fontCache.TryGetValue(key, out OpenTypeFont? cached))
        {
            return cached;
        }

        cached = FontProgramLoader.Load(resolution, cancellationToken);
        fontCache[key] = cached;
        return cached;
    }
}
