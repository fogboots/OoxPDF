using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Chart text parts render through separate calls whose resources merge into

    // one page-level list, so per-call numbering would collide (duplicate /Font
    // keys, last wins). Callers pass that shared list for page-unique sequencing;
    // the appended slice is already in the list, so callers must not re-add it.
    private static RenderedFonts RenderChartTextRuns(
        List<TextRun> runs,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> chartFonts,
        string resourcePrefix,
        PresentationFontResolver? fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        if (runs.Count == 0)
        {
            return new RenderedFonts(new Dictionary<FontRequest, RenderedFont>(FontRequestKeyComparer.OrdinalIgnoreCaseFamily), []);
        }
        runs = SplitRunsByResolvedTypeface(runs, fontResolver ?? new PresentationFontResolver(null), diagnosticSink, CancellationToken.None);

        RenderedFonts renderedFonts = CreateRenderedFonts(runs, fontResolver ?? new PresentationFontResolver(null), resourcePrefix, CancellationToken.None, diagnosticSink, chartFonts);
        DrawTextRunsWithFonts(runs, graphics, renderedFonts.Fonts);
        chartFonts.AddRange(renderedFonts.Resources);
        return renderedFonts;
    }

    // Legacy runs address embedded fonts by requested family, but CFF substitution
    // embeds fallback faces under fallback keys: without rewriting, substituted runs
    // would miss the lookup and vanish. Rewrite CFF-family runs through the same
    // per-glyph fallback the estimator measured with (F03). Families resolving to
    // embeddable-or-missing fonts keep legacy behavior exactly (missing fonts stay
    // invisible, as before).
    private static List<TextRun> SplitRunsByResolvedTypeface(
        IReadOnlyList<TextRun> runs,
        PresentationFontResolver fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        var rewritten = new List<TextRun>(runs.Count);
        // T01: family-name keys are typed FontRequests with explicit case-insensitive
        // family comparison (opaque source IDs are never normalized; see
        // FontRequestKeyComparer and PresentationFontResolver).
        var verdicts = new Dictionary<FontRequest, OpenTypeFont?>(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);
        var reported = new HashSet<FontRequest>(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);
        TextAdvanceEstimator? estimator = null;
        foreach (TextRun run in runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string requestedFamily = PptxFontFallbackRules.ResolveDefaultLatinTypeface(run.FontFamily);
            FontRequest familyKey = new(requestedFamily, run.Bold, run.Italic);
            if (!verdicts.TryGetValue(familyKey, out OpenTypeFont? primary))
            {
                primary = fontResolver.ResolvePresentationOpenTypeFont(new FontRequest(requestedFamily, run.Bold, run.Italic), cancellationToken)?.Font;
                verdicts[familyKey] = primary;
            }

            if (primary is null || primary.HasTrueTypeOutlines)
            {
                rewritten.Add(run);
                continue;
            }

            estimator ??= new TextAdvanceEstimator(fontResolver, cancellationToken);
            List<(string? Typeface, string Text)> segments = SplitRunByFallbackTypeface(run, estimator, cancellationToken);
            if (segments.Count == 1)
            {
                rewritten.Add(run with { FontFamily = segments[0].Typeface });
            }
            else
            {
                double cursorX = run.X;
                foreach ((string? typeface, string text) in segments)
                {
                    double width = estimator.Measure(text, run.FontSize, typeface, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled);
                    rewritten.Add(run with
                    {
                        Text = text,
                        X = cursorX,
                        Width = width,
                        FontFamily = typeface,
                        Alignment = TextAlignment.Left,
                        PreventCoalesce = true
                    });
                    cursorX += width;
                }
            }

            if (reported.Add(familyKey))
            {
                diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                    "FONT_UNSUPPORTED_OUTLINES",
                    OoxPdfSeverity.Warning,
                    "Font '" + requestedFamily + "' has no embeddable TrueType outlines (CFF/OpenType-CFF); affected runs were substituted with fallback typefaces.",
                    PartName: null,
                    SlideIndex: null,
                    PageIndex: null,
                    Feature: requestedFamily,
                    Fallback: "Per-glyph fallback typeface"));
            }
        }

        return rewritten;
    }

    private static List<(string? Typeface, string Text)> SplitRunByFallbackTypeface(TextRun run, TextAdvanceEstimator estimator, CancellationToken cancellationToken)
    {
        var segments = new List<(string? Typeface, string Text)>();
        string? segmentTypeface = null;
        var segmentText = new StringBuilder();
        void FlushSegment()
        {
            if (segmentText.Length != 0)
            {
                // RV01 residual: estimator-unresolvable runes keep a null-typeface segment so
                // downstream default-family handling substitutes a diagnosed question mark.
                segments.Add((segmentTypeface, segmentText.ToString()));
                segmentText.Clear();
            }
        }

        foreach (Rune rune in run.Text.EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? typeface = estimator.ResolveGlyphFont(run.FontFamily, run.Bold, run.Italic, rune.Value)?.Typeface;
            bool sameFamily = segmentTypeface is null ? typeface is null : typeface is not null && typeface.Equals(segmentTypeface, StringComparison.OrdinalIgnoreCase);
            if (!sameFamily)
            {
                FlushSegment();
                segmentTypeface = null;
            }

            segmentTypeface = typeface;
            segmentText.Append(char.ConvertFromUtf32(rune.Value));
        }

        FlushSegment();
        return segments;
    }

    // Counts page-level names of the form prefix + digits so chart text parts
    // that render through separate calls share one numbering sequence per page.
    private static int CountPrefixedResourceNames(List<PdfFontResource>? nameScope, string prefix)
    {
        if (nameScope is null)
        {
            return 0;
        }

        string sanitizedPrefix = PdfEmbeddedFont.SanitizeName(prefix);
        int count = 0;
        foreach (PdfFontResource resource in nameScope)
        {
            string name = PdfEmbeddedFont.SanitizeName(resource.ResourceName);
            if (name.Length > sanitizedPrefix.Length &&
                name.StartsWith(sanitizedPrefix, StringComparison.Ordinal))
            {
                bool digitsOnly = true;
                for (int i = sanitizedPrefix.Length; i < name.Length; i++)
                {
                    if (!char.IsAsciiDigit(name[i]))
                    {
                        digitsOnly = false;
                        break;
                    }
                }

                if (digitsOnly)
                {
                    count++;
                }
            }
        }

        return count;
    }

    private static RenderedFonts CreateRenderedFonts(
        IReadOnlyList<PptxPositionedTextSpan> textSpans,
        IReadOnlyList<TextRun> legacyTextRuns,
        PresentationFontResolver fontResolver,
        string resourcePrefix,
        CancellationToken cancellationToken)
    {
        var uses = new List<TextFontUse>();
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (IGrouping<string, PptxTextGlyphLayout> group in span.GlyphSpan.Glyphs.GroupBy(
                        glyph => string.IsNullOrWhiteSpace(glyph.Typeface) ? PptxFontFallbackRules.ResolveDefaultLatinTypeface(span.Run.FontFamily) : glyph.Typeface,
                         StringComparer.OrdinalIgnoreCase))
            {
                uses.Add(new TextFontUse(group.Key, span.GlyphSpan.Bold, span.GlyphSpan.Italic, group.Select(glyph => glyph.CodePoint).ToArray()));
            }
        }

        foreach (TextRun run in CoalesceUnderlineRuns(CoalesceAdjacentTextRuns(legacyTextRuns, compareHighlight: false)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string familyName = PptxFontFallbackRules.ResolveDefaultLatinTypeface(run.FontFamily);
            uses.Add(new TextFontUse(familyName, run.Bold, run.Italic, run.Text.EnumerateRunes().Select(rune => rune.Value).ToArray()));
        }

        return CreateRenderedFonts(uses, fontResolver, resourcePrefix, cancellationToken);
    }


    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextRun> textRuns, PresentationFontResolver fontResolver, string resourcePrefix, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink = null, List<PdfFontResource>? nameScope = null, bool includeFallbackFaces = true)
    {
        if (textRuns.Count == 0)
        {
            return new RenderedFonts(new Dictionary<FontRequest, RenderedFont>(FontRequestKeyComparer.OrdinalIgnoreCaseFamily), []);
        }

        textRuns = CoalesceAdjacentTextRuns(textRuns, compareHighlight: false);
        textRuns = CoalesceUnderlineRuns(textRuns);
        return CreateRenderedFonts(textRuns
            .Select(run => new TextFontUse(
                PptxFontFallbackRules.ResolveDefaultLatinTypeface(run.FontFamily),
                run.Bold,
                run.Italic,
                run.Text.EnumerateRunes().Select(rune => rune.Value).ToArray()))
            .ToArray(), fontResolver, resourcePrefix, cancellationToken, diagnosticSink, nameScope, includeFallbackFaces);
    }

    // RV02: complex-script runs render without shaping; report each behavior family
    // once per conversion (the resolver set dedups across slide and chart preparation).
    private static void ReportComplexScriptApproximation(
        IReadOnlyList<TextFontUse> uses,
        PresentationFontResolver fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        OoxComplexScriptKind seen = OoxComplexScriptKind.None;
        foreach (TextFontUse use in uses)
        {
            cancellationToken.ThrowIfCancellationRequested();
            seen |= OoxComplexScript.DetectCodepoints(use.CodePoints);
        }

        foreach (OoxComplexScriptKind kind in new[]
        {
            OoxComplexScriptKind.Joining,
            OoxComplexScriptKind.Reordering,
            OoxComplexScriptKind.Bidirectional,
            OoxComplexScriptKind.CombiningMark
        })
        {
            if (seen.HasFlag(kind) && fontResolver.ReportedMissingFontDiagnostics.Add("script:" + kind))
            {
                diagnosticSink(OoxComplexScript.CreateApproximationDiagnostic(kind));
            }
        }
    }

    // RV01: slide preparation passes includeFallbackFaces: false and covers split-run
    // families separately (emission looks those up); chart preparation keeps the
    // inline behavior for its run families.
    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextFontUse> uses, PresentationFontResolver fontResolver, string resourcePrefix, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink = null, List<PdfFontResource>? nameScope = null, bool includeFallbackFaces = true)
    {
        if (uses.Count == 0)
        {
            return new RenderedFonts(new Dictionary<FontRequest, RenderedFont>(FontRequestKeyComparer.OrdinalIgnoreCaseFamily), []);
        }

        ReportComplexScriptApproximation(uses, fontResolver, diagnosticSink, cancellationToken);

        uses = SubstituteUnembeddableFontUses(uses, fontResolver, diagnosticSink, cancellationToken);

        var fonts = new Dictionary<FontRequest, RenderedFont>(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);
        var resources = new List<PdfFontResource>();
        foreach (IGrouping<FontRequest, TextFontUse> group in uses.GroupBy(use => new FontRequest(use.FamilyName, use.Bold, use.Italic), FontRequestKeyComparer.OrdinalIgnoreCaseFamily))
        {
            cancellationToken.ThrowIfCancellationRequested();
            TextFontUse first = group.First();
            (FontFaceResolution Resolution, OpenTypeFont Font)? resolved = fontResolver.ResolvePresentationOpenTypeFont(new FontRequest(first.FamilyName, first.Bold, first.Italic), cancellationToken);
            if (resolved is null || !resolved.Value.Font.HasTrueTypeOutlines)
            {
                if (!includeFallbackFaces)
                {
                    continue;
                }

                // RV01: uses without a usable embeddable face render with the diagnosed
                // built-in fallback instead of vanishing.
                int[] fallbackCodepoints = group.SelectMany(use => use.CodePoints).Where(code => !PdfFallbackFont.IsNonRenderedControl(code)).ToArray();
                if (fallbackCodepoints.Length == 0)
                {
                    continue;
                }

                PdfFallbackFont face = PdfFallbackFont.ForStyle(first.Bold, first.Italic);
                string fallbackName = PdfFallbackFont.ResourceNameFor(first.Bold, first.Italic);
                var fallbackResolution = new FontFaceResolution(first.FamilyName, face.FaceName, new FontStyleKey(first.Bold, first.Italic), new MemoryFontProgramSource("fallback:" + face.FaceName, ReadOnlyMemory<byte>.Empty), IsFallback: true);
                fonts[group.Key] = new RenderedFont(fallbackName, null, fallbackResolution, false, false, new PdfFallbackFontResource(fallbackName, face));
                fontResolver.UsedFallbackFaces.Add(face);
                ReportMissingFontFallback(first.FamilyName, fontResolver, diagnosticSink);
                continue;
            }

            FontFaceResolution resolution = resolved.Value.Resolution;
            OpenTypeFont font = resolved.Value.Font;
            int[] codepoints = AddQuestionMarkForUncoveredCodepoints(group, font, fontResolver, diagnosticSink, cancellationToken);
            PdfEmbeddedFont embedded = fontResolver.GetOrCreateSubset(font, codepoints, cancellationToken);
            string resourceName = resourcePrefix + (CountPrefixedResourceNames(nameScope, resourcePrefix) + resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            fonts[group.Key] = new RenderedFont(resourceName, embedded, resolution, first.Bold && !resolution.Bold, first.Italic && !resolution.Italic);
            resources.Add(new PdfFontResource(resourceName, embedded));
        }

        return new RenderedFonts(fonts, resources);
    }

    // RV01: reports a family whose text has no usable embeddable face once per
    // conversion (preparation runs per slide and per chart part).
    private static void ReportMissingFontFallback(string familyName, PresentationFontResolver fontResolver, Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        if (!fontResolver.ReportedMissingFontDiagnostics.Add("face:" + familyName))
        {
            return;
        }

        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            "FONT_NO_USABLE_FACE",
            OoxPdfSeverity.Warning,
            "No usable embeddable font for " + familyName + "; text renders with the built-in fallback typeface. Characters outside WinAnsi appear as question mark.",
            PartName: null,
            SlideIndex: null,
            PageIndex: null,
            Feature: familyName,
            Fallback: "Built-in fallback typeface"));
    }

    // RV01: codepoints with no glyph in the emitting face are substituted with
    // question mark at emission; the subset carries it exactly when needed, and the
    // substitution is diagnosed once per family with the affected codepoints.
    private static int[] AddQuestionMarkForUncoveredCodepoints(IGrouping<FontRequest, TextFontUse> group, OpenTypeFont font, PresentationFontResolver fontResolver, Action<OoxPdfDiagnostic>? diagnosticSink, CancellationToken cancellationToken)
    {
        int[] codepoints = group.SelectMany(use => use.CodePoints).ToArray();
        List<int>? missing = null;
        TextAdvanceEstimator? estimator = null;
        foreach (int codePoint in codepoints)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // Emission splits runes the primary face lacks onto covering fallback
            // faces, so only runes no face covers become question marks.
            if (!PdfFallbackFont.IsNonRenderedControl(codePoint) && font.MapCodePoint(codePoint) == 0 &&
                (estimator ??= new TextAdvanceEstimator(fontResolver, cancellationToken)).ResolveGlyphFont(group.Key.FamilyName, group.Key.Bold, group.Key.Italic, codePoint) is null)
            {
                missing ??= new List<int>();
                if (!missing.Contains(codePoint))
                {
                    missing.Add(codePoint);
                }
            }
        }

        if (missing is null)
        {
            return codepoints;
        }

        ReportMissingGlyphSubstitution(group.Key.FamilyName, missing, fontResolver, diagnosticSink);

        if (font.MapCodePoint(0x3F) == 0)
        {
            return codepoints;
        }

        var withQuestionMark = new int[codepoints.Length + 1];
        Array.Copy(codepoints, withQuestionMark, codepoints.Length);
        withQuestionMark[codepoints.Length] = 0x3F;
        return withQuestionMark;
    }

    // RV01: codepoints with no glyph in any usable face are named once per family;
    // substituted marks show as question mark.
    private static void ReportMissingGlyphSubstitution(string familyName, List<int> missing, PresentationFontResolver fontResolver, Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        missing.Sort();
        if (fontResolver.ReportedMissingFontDiagnostics.Add("glyph:" + familyName))
        {
            const int maxListedCodepoints = 12;
            string listed = string.Join(", ", missing.Take(maxListedCodepoints).Select(code => "U+" + code.ToString("X4")));
            if (missing.Count > maxListedCodepoints)
            {
                listed += ", and " + (missing.Count - maxListedCodepoints).ToString(CultureInfo.InvariantCulture) + " more";
            }

            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "FONT_MISSING_GLYPHS",
                OoxPdfSeverity.Warning,
                missing.Count.ToString(CultureInfo.InvariantCulture) + " characters of " + familyName + " have no glyph in any usable face (" + listed + ") and show as question mark.",
                PartName: null,
                SlideIndex: null,
                PageIndex: null,
                Feature: familyName,
                Fallback: "Question-mark substitution"));
        }
    }

    // RV01: emission splits spans by glyph typeface, so families that only appear on
    // split runs need coverage when preparation skipped them: unresolvable families
    // take the diagnosed built-in fallback, while resolvable families gain embedded
    // subsets so their glyphs do not vanish at the emission lookup.
    private static void AddSplitFallbackFaces(
        Dictionary<FontRequest, RenderedFont> fonts,
        IEnumerable<PptxPositionedTextSpan> spans,
        PresentationFontResolver fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        List<PdfFontResource> resources,
        string resourcePrefix)
    {
        var embeddableUses = new List<TextFontUse>();
        foreach (PptxPositionedTextSpan emissionSpan in spans.SelectMany(SplitSpanByGlyphTypeface))
        {
            cancellationToken.ThrowIfCancellationRequested();
            TextRun run = emissionSpan.Run;
            var key = FontRequestForRun(run);
            if (fonts.ContainsKey(key))
            {
                continue;
            }

            bool hasRenderable = false;
            foreach (Rune rune in run.Text.EnumerateRunes())
            {
                if (!PdfFallbackFont.IsNonRenderedControl(rune))
                {
                    hasRenderable = true;
                    break;
                }
            }

            if (!hasRenderable)
            {
                continue;
            }

            (FontFaceResolution Resolution, OpenTypeFont Font)? resolved = fontResolver.ResolvePresentationOpenTypeFont(key, cancellationToken);
            if (resolved is not null && resolved.Value.Font.HasTrueTypeOutlines)
            {
                // RV01 residual follow-up: resolvable split families that were never prepared
                // get embedded subsets instead of vanishing at emission.
                embeddableUses.Add(new TextFontUse(key.FamilyName, key.Bold, key.Italic, emissionSpan.GlyphSpan.Glyphs.Select(glyph => glyph.CodePoint).ToArray()));
                continue;
            }

            PdfFallbackFont face = PdfFallbackFont.ForStyle(key.Bold, key.Italic);
            string fallbackName = PdfFallbackFont.ResourceNameFor(key.Bold, key.Italic);
            var fallbackResolution = new FontFaceResolution(key.FamilyName, face.FaceName, new FontStyleKey(key.Bold, key.Italic), new MemoryFontProgramSource("fallback:" + face.FaceName, ReadOnlyMemory<byte>.Empty), IsFallback: true);
            fonts[key] = new RenderedFont(fallbackName, null, fallbackResolution, false, false, new PdfFallbackFontResource(fallbackName, face));
            fontResolver.UsedFallbackFaces.Add(face);
            ReportMissingFontFallback(key.FamilyName, fontResolver, diagnosticSink);
        }

        if (embeddableUses.Count != 0)
        {
            RenderedFonts embeddedSplitFonts = CreateRenderedFonts(embeddableUses, fontResolver, resourcePrefix, cancellationToken, diagnosticSink, resources, includeFallbackFaces: true);
            foreach (KeyValuePair<FontRequest, RenderedFont> entry in embeddedSplitFonts.Fonts)
            {
                fonts[entry.Key] = entry.Value;
            }

            resources.AddRange(embeddedSplitFonts.Resources);
        }
    }

    // CFF/OpenType-CFF fonts have valid metrics but no TrueType outlines, so the
    // PDF subsetter cannot embed them. Rewrite affected uses to the fallback faces
    // the measurement estimator resolves for the same code points, so measured and
    // emitted glyphs use the same face (F03); report each substituted group once.
    private static IReadOnlyList<TextFontUse> SubstituteUnembeddableFontUses(
        IReadOnlyList<TextFontUse> uses,
        PresentationFontResolver fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        var groupFonts = new Dictionary<FontRequest, OpenTypeFont?>(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);
        bool needsSubstitution = false;
        foreach (IGrouping<FontRequest, TextFontUse> group in uses.GroupBy(use => new FontRequest(use.FamilyName, use.Bold, use.Italic), FontRequestKeyComparer.OrdinalIgnoreCaseFamily))
        {
            cancellationToken.ThrowIfCancellationRequested();
            TextFontUse first = group.First();
            OpenTypeFont? font = fontResolver.ResolvePresentationOpenTypeFont(new FontRequest(first.FamilyName, first.Bold, first.Italic), cancellationToken)?.Font;
            groupFonts[group.Key] = font;
            if (font is not null && !font.HasTrueTypeOutlines)
            {
                needsSubstitution = true;
            }
        }

        if (!needsSubstitution)
        {
            return uses;
        }

        var estimator = new TextAdvanceEstimator(fontResolver, cancellationToken);
        var reported = new HashSet<FontRequest>(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);
        var substituted = new List<TextFontUse>(uses.Count);
        foreach (TextFontUse use in uses)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (groupFonts[new FontRequest(use.FamilyName, use.Bold, use.Italic)] is not { } font || font.HasTrueTypeOutlines)
            {
                substituted.Add(use);
                continue;
            }
            (int CodePoint, string? Typeface)[] resolvedCodepoints = use.CodePoints
                .Select(codePoint => (CodePoint: codePoint, Typeface: estimator.ResolveGlyphFont(use.FamilyName, use.Bold, use.Italic, codePoint)?.Typeface))
                .ToArray();
            List<int>? droppedCodepoints = null;
            foreach ((int CodePoint, string? Typeface) resolved in resolvedCodepoints)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrEmpty(resolved.Typeface) || PdfFallbackFont.IsNonRenderedControl(resolved.CodePoint))
                {
                    continue;
                }

                droppedCodepoints ??= new List<int>();
                if (!droppedCodepoints.Contains(resolved.CodePoint))
                {
                    droppedCodepoints.Add(resolved.CodePoint);
                }
            }

            foreach (IGrouping<string, (int CodePoint, string? Typeface)> split in resolvedCodepoints
                         .Where(resolved => !string.IsNullOrEmpty(resolved.Typeface))
                         .GroupBy(resolved => resolved.Typeface!, StringComparer.OrdinalIgnoreCase))
            {
                cancellationToken.ThrowIfCancellationRequested();
                substituted.Add(new TextFontUse(split.Key, use.Bold, use.Italic, split.Select(resolved => resolved.CodePoint).ToArray()));
            }

            if (droppedCodepoints is not null)
            {
                // RV01 residual: runes no usable face covers are named with the
                // substitution diagnostic instead of vanishing from preparation.
                ReportMissingGlyphSubstitution(use.FamilyName, droppedCodepoints, fontResolver, diagnosticSink);
            }

            if (reported.Add(new FontRequest(use.FamilyName, use.Bold, use.Italic)))
            {
                diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                    "FONT_UNSUPPORTED_OUTLINES",
                    OoxPdfSeverity.Warning,
                    "Font '" + use.FamilyName + "' has no embeddable TrueType outlines (CFF/OpenType-CFF); affected runs were substituted with fallback typefaces.",
                    PartName: null,
                    SlideIndex: null,
                    PageIndex: null,
                    Feature: use.FamilyName,
                    Fallback: "Per-glyph fallback typeface"));
            }
        }

        return substituted;
    }

    private static void DrawTextRunsWithFonts(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics, IReadOnlyDictionary<FontRequest, RenderedFont> fonts)
    {
        DrawHighlightRunsWithFonts();
        textRuns = CoalesceAdjacentTextRuns(textRuns, compareHighlight: false);
        textRuns = CoalesceUnderlineRuns(textRuns);
        foreach (TextRun run in textRuns)
        {
            if (fonts.TryGetValue(FontRequestForRun(run), out RenderedFont rendered))
            {
                if (rendered.FallbackFace is { } fallback)
                {
                    DrawFallbackWrappedRun(fallback, run);
                }
                else if (rendered.Font is { } embedded)
                {
                    DrawWrappedRun(rendered.ResourceName, embedded, run, rendered.SyntheticBold, rendered.SyntheticItalic);
                }
            }
        }

        void DrawWrappedRun(string resourceName, PdfEmbeddedFont embedded, TextRun run, bool syntheticBold, bool syntheticItalic)
        {
            graphics.SaveState();
            if (HasTextTransform(run))
            {
                ApplyTextTransform(graphics, run);
            }

            (double clipX, double clipY, double clipWidth, double clipHeight) = InverseTransformClip(run, run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
            graphics.ClipRectangleEvenOdd(clipX, clipY, clipWidth, clipHeight);
            TextGlyphRun? glyphRun = BuildTextGlyphRun(resourceName, embedded, run, syntheticBold, syntheticItalic);
            if (glyphRun is not null)
            {
                bool needsTextAlpha = run.Alpha < 1d - PptxTextMetricRules.TextStateTolerance ||
                    (run.Outline is { } runOutline && runOutline.Alpha < 1d - PptxTextMetricRules.TextStateTolerance);
                if (needsTextAlpha)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(run.Alpha, run.Outline?.Alpha ?? 1d);
                }

                DrawGlyphText(graphics, glyphRun);

                if (run.Underline)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    if (TryGetUnderlineRectangle(embedded, glyphRun, out TextDecorationRectangle underline))
                    {
                        FillTextDecorationRectangleEvenOdd(graphics, underline);
                    }
                }

                if (run.Strike)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    if (TryGetStrikeRectangle(embedded, glyphRun, out TextDecorationRectangle strike))
                    {
                        FillTextDecorationRectangleEvenOdd(graphics, strike);
                    }
                }

                if (needsTextAlpha)
                {
                    graphics.RestoreState();
                }
            }

            graphics.RestoreState();
        }

        // RV01: fallback legacy runs distribute runes evenly across the label box
        // (these runs carry box geometry, not measured advances), encoding WinAnsi
        // with question-mark substitution for unencodable runes.
        void DrawFallbackWrappedRun(PdfFallbackFontResource fallback, TextRun run)
        {
            graphics.SaveState();
            if (HasTextTransform(run))
            {
                ApplyTextTransform(graphics, run);
            }

            (double clipX, double clipY, double clipWidth, double clipHeight) = InverseTransformClip(run, run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
            graphics.ClipRectangleEvenOdd(clipX, clipY, clipWidth, clipHeight);
            Rune[] runes = [.. run.Text.EnumerateRunes()];
            double baselineY = run.Y + run.BaselineOffset;
            if (runes.Length != 0 && BaselineIntersectsClip(run, baselineY))
            {
                bool needsTextAlpha = run.Alpha < 1d - PptxTextMetricRules.TextStateTolerance ||
                    (run.Outline is { } runOutline && runOutline.Alpha < 1d - PptxTextMetricRules.TextStateTolerance);
                if (needsTextAlpha)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(run.Alpha, run.Outline?.Alpha ?? 1d);
                }

                double pdfFontSize = OfficePdfTextEmissionProfile.FontSize(run.FontSize);
                var glyphs = new List<PdfFallbackGlyph>(runes.Length);
                double slot = run.Width / runes.Length;
                for (int i = 0; i < runes.Length; i++)
                {
                    if (PdfFallbackFont.IsNonRenderedControl(runes[i]))
                    {
                        continue;
                    }

                    PdfFallbackFont.TryEncodeWinAnsi(runes[i].Value, out byte code);
                    glyphs.Add(new PdfFallbackGlyph(run.X + i * slot, baselineY, code));
                }

                graphics.DrawFallbackText(fallback.ResourceName, pdfFontSize, run.Color.Red, run.Color.Green, run.Color.Blue, glyphs);
                if (run.Underline || run.Strike)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    FillFallbackTextDecoration(graphics, run.X, baselineY, run.Width, pdfFontSize, run.Underline, run.Strike);
                }

                if (needsTextAlpha)
                {
                    graphics.RestoreState();
                }
            }

            graphics.RestoreState();
        }

        void DrawHighlightRunsWithFonts()
        {
            foreach (TextRun run in CoalesceHighlightRuns())
            {
                if (run.HighlightColor is not { } highlight || !fonts.TryGetValue(FontRequestForRun(run), out RenderedFont rendered))
                {
                    continue;
                }

                // RV01: fallback highlights use the measured run width with diagnosed
                // constants instead of vanishing.
                if (rendered.FallbackFace is not null)
                {
                    double baselineY = run.Y + run.BaselineOffset;
                    FillFallbackHighlight(graphics, run, highlight, baselineY, run.Width);
                }
                else if (rendered.Font is { } highlightEmbedded)
                {
                    DrawHighlightRun(highlightEmbedded, run);
                }
            }

            IReadOnlyList<TextRun> CoalesceHighlightRuns()
            {
                var coalesced = new List<TextRun>(textRuns.Count);
                foreach (TextRun run in textRuns)
                {
                    if (run.Text.Length == 0 || run.HighlightColor is null)
                    {
                        continue;
                    }

                    if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1], run, true, true))
                    {
                        TextRun previous = coalesced[^1];
                        coalesced[^1] = previous with
                        {
                            Text = previous.Text + run.Text,
                            Width = run.X + run.Width - previous.X
                        };
                    }
                    else
                    {
                        coalesced.Add(run);
                    }
                }

                return coalesced;
            }

            void DrawHighlightRun(PdfEmbeddedFont embedded, TextRun run)
            {
                if (run.HighlightColor is not { } highlight)
                {
                    return;
                }

                double baselineY = run.Y + run.BaselineOffset;
                double lineWidth = MeasureRenderedText(embedded, run.Text, run.FontSize, run.CharacterSpacing, run.KerningEnabled);
                DrawHighlightRectangle(graphics, embedded, run, highlight, baselineY, lineWidth);
            }
        }
    }

    private static void DrawTextSpansWithFonts(IReadOnlyList<PptxPositionedTextSpan> textSpans, PdfGraphicsBuilder graphics, IReadOnlyDictionary<FontRequest, RenderedFont> fonts, PptxTextHyperlinkScope? hyperlinkScope = null)
    {
        DrawHighlightSpansWithFonts();
        textSpans = SplitLeadingSpacesAtHighlightBoundaries(textSpans);
        textSpans = CoalesceAdjacentTextSpans(textSpans, compareHighlight: true);
        textSpans = CoalesceUnderlineSpans(textSpans);
        textSpans = ApplyOfficePdfCharacterSpacing(textSpans);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            foreach (PptxPositionedTextSpan emissionSpan in SplitSpanByGlyphTypeface(span))
            {
                TextRun run = emissionSpan.Run;
                if (fonts.TryGetValue(FontRequestForRun(run), out RenderedFont rendered))
                {
                    if (rendered.FallbackFace is { } fallback)
                    {
                        DrawFallbackWrappedSpan(fallback, emissionSpan);
                        // Fallback spans stay link-free like other unresolved spans.
                    }
                    else if (rendered.Font is { } embedded)
                    {
                        DrawWrappedSpan(rendered.ResourceName, embedded, emissionSpan, rendered.SyntheticBold, rendered.SyntheticItalic);
                        hyperlinkScope?.CollectEmissionSpan(emissionSpan, embedded);
                    }
                }
            }
        }

        // RV01: fallback spans distribute runes across the layout span width in
        // UTF-16 proportion (layout measured these spans without per-rune advances),
        // encoding WinAnsi with question-mark substitution for unencodable runes.
        void DrawFallbackWrappedSpan(PdfFallbackFontResource fallback, PptxPositionedTextSpan span)
        {
            TextRun run = span.Run;
            graphics.SaveState();
            if (HasTextTransform(run))
            {
                ApplyTextTransform(graphics, run);
            }

            (double clipX, double clipY, double clipWidth, double clipHeight) = InverseTransformClip(run, run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
            graphics.ClipRectangleEvenOdd(clipX, clipY, clipWidth, clipHeight);
            Rune[] runes = [.. run.Text.EnumerateRunes()];
            double baselineY = run.Y + run.BaselineOffset;
            if (runes.Length != 0 && BaselineIntersectsClip(run, baselineY))
            {
                bool needsTextAlpha = run.Alpha < 1d - PptxTextMetricRules.TextStateTolerance ||
                    (run.Outline is { } runOutline && runOutline.Alpha < 1d - PptxTextMetricRules.TextStateTolerance);
                if (needsTextAlpha)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(run.Alpha, run.Outline?.Alpha ?? 1d);
                }

                double pdfFontSize = OfficePdfTextEmissionProfile.FontSize(run.FontSize);
                int totalUnits = 0;
                foreach (Rune rune in runes)
                {
                    totalUnits += rune.Utf16SequenceLength;
                }

                double gaps = Math.Max(0, runes.Length - 1) * run.CharacterSpacing;
                double unit = totalUnits == 0 ? 0d : Math.Max(0d, run.Width - gaps) / totalUnits;
                var glyphs = new List<PdfFallbackGlyph>(runes.Length);
                int usedUnits = 0;
                int gapCount = 0;
                foreach (Rune rune in runes)
                {
                    if (!PdfFallbackFont.IsNonRenderedControl(rune))
                    {
                        PdfFallbackFont.TryEncodeWinAnsi(rune.Value, out byte code);
                        glyphs.Add(new PdfFallbackGlyph(run.X + usedUnits * unit + gapCount * run.CharacterSpacing, baselineY, code));
                    }

                    usedUnits += rune.Utf16SequenceLength;
                    gapCount++;
                }

                graphics.DrawFallbackText(fallback.ResourceName, pdfFontSize, run.Color.Red, run.Color.Green, run.Color.Blue, glyphs);
                if (run.Underline || run.Strike)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    FillFallbackTextDecoration(graphics, run.X, baselineY, run.Width, pdfFontSize, run.Underline, run.Strike);
                }

                if (needsTextAlpha)
                {
                    graphics.RestoreState();
                }
            }

            graphics.RestoreState();
        }

        void DrawWrappedSpan(string resourceName, PdfEmbeddedFont embedded, PptxPositionedTextSpan span, bool syntheticBold, bool syntheticItalic)
        {
            TextRun run = span.Run;
            graphics.SaveState();
            if (HasTextTransform(run))
            {
                ApplyTextTransform(graphics, run);
            }

            (double clipX, double clipY, double clipWidth, double clipHeight) = InverseTransformClip(run, run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
            graphics.ClipRectangleEvenOdd(clipX, clipY, clipWidth, clipHeight);
            TextGlyphRun? glyphRun = BuildTextGlyphRun(resourceName, embedded, span, syntheticBold, syntheticItalic);
            if (glyphRun is not null)
            {
                bool needsTextAlpha = run.Alpha < 1d - PptxTextMetricRules.TextStateTolerance ||
                    (run.Outline is { } runOutline && runOutline.Alpha < 1d - PptxTextMetricRules.TextStateTolerance);
                if (needsTextAlpha)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(run.Alpha, run.Outline?.Alpha ?? 1d);
                }

                DrawGlyphText(graphics, glyphRun);

                if (run.Underline)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    if (TryGetUnderlineRectangle(embedded, glyphRun, out TextDecorationRectangle underline))
                    {
                        FillTextDecorationRectangleEvenOdd(graphics, underline);
                    }
                }

                if (run.Strike)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    if (TryGetStrikeRectangle(embedded, glyphRun, out TextDecorationRectangle strike))
                    {
                        FillTextDecorationRectangleEvenOdd(graphics, strike);
                    }
                }

                if (needsTextAlpha)
                {
                    graphics.RestoreState();
                }
            }

            graphics.RestoreState();
        }

        void DrawHighlightSpansWithFonts()
        {
            foreach (PptxPositionedTextSpan span in CoalesceHighlightSpans())
            {
                TextRun run = span.Run;
                if (run.HighlightColor is not { } highlight || !fonts.TryGetValue(FontRequestForRun(run), out RenderedFont rendered))
                {
                    continue;
                }

                // RV01: fallback highlights use the layout span width with diagnosed
                // constants instead of vanishing.
                if (rendered.FallbackFace is not null)
                {
                    double baselineY = span.LineBox?.BaselineY ?? run.Y + run.BaselineOffset;
                    FillFallbackHighlight(graphics, run, highlight, baselineY, span.GlyphSpan.NaturalWidth);
                }
                else if (rendered.Font is { } highlightEmbedded)
                {
                    DrawHighlightSpan(highlightEmbedded, span);
                }
            }

            IReadOnlyList<PptxPositionedTextSpan> CoalesceHighlightSpans()
            {
                var coalesced = new List<PptxPositionedTextSpan>(textSpans.Count);
                foreach (PptxPositionedTextSpan span in textSpans)
                {
                    TextRun run = span.Run;
                    if (run.Text.Length == 0 || run.HighlightColor is null)
                    {
                        continue;
                    }

                    if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1].Run, run, true, true))
                    {
                        PptxPositionedTextSpan previous = coalesced[^1];
                        TextRun mergedRun = previous.Run with
                        {
                            Text = previous.Run.Text + run.Text,
                            Width = run.X + run.Width - previous.Run.X
                        };
                        coalesced[^1] = previous with
                        {
                            Run = mergedRun,
                            EndX = span.EndX,
                            Atoms = previous.Atoms.Concat(span.Atoms).ToArray(),
                            GlyphSpan = MergeGlyphSpans(mergedRun, previous.Run, previous.GlyphSpan, run, span.GlyphSpan)
                        };
                    }
                    else
                    {
                        coalesced.Add(span);
                    }
                }

                return coalesced;
            }

            void DrawHighlightSpan(PdfEmbeddedFont embedded, PptxPositionedTextSpan span)
            {
                TextRun run = span.Run;
                if (run.HighlightColor is not { } highlight)
                {
                    return;
                }

                double baselineY = span.LineBox?.BaselineY ?? run.Y + run.BaselineOffset;
                DrawHighlightRectangle(graphics, embedded, run, highlight, baselineY, span.GlyphSpan.NaturalWidth);
            }
        }
    }

    private static FontRequest FontRequestForRun(TextRun run)
    {
        return new FontRequest(PptxFontFallbackRules.ResolveDefaultLatinTypeface(run.FontFamily), run.Bold, run.Italic);
    }

    // RV01: mirrors the atom substitution so hex encoding sees the same runes.
    private static string SubstituteUncoveredRunes(PdfEmbeddedFont embedded, string text)
    {
        bool clean = true;
        foreach (Rune rune in text.EnumerateRunes())
        {
            if (!PdfFallbackFont.IsNonRenderedControl(rune) && embedded.Font.MapCodePoint(rune.Value) == 0 && embedded.Font.MapCodePoint(0x3F) != 0)
            {
                clean = false;
                break;
            }
        }

        if (clean)
        {
            return text;
        }

        var builder = new StringBuilder(text.Length);
        foreach (Rune rune in text.EnumerateRunes())
        {
            builder.Append(!PdfFallbackFont.IsNonRenderedControl(rune) && embedded.Font.MapCodePoint(rune.Value) == 0 ? "?" : rune.ToString());
        }

        return builder.ToString();
    }

    private static TextGlyphRun? BuildTextGlyphRun(string resourceName, PdfEmbeddedFont embedded, TextRun run, bool syntheticBold, bool syntheticItalic)
    {
        string emissionText = SubstituteUncoveredRunes(embedded, run.Text);
        string glyphHex = embedded.EncodeGlyphHex(emissionText);
        double baselineY = run.Y + run.BaselineOffset;
        if (glyphHex.Length == 0 || !BaselineIntersectsClip(run, baselineY))
        {
            return null;
        }

        IReadOnlyList<TextGlyphAtom> glyphs = BuildTextGlyphAtoms();
        // RV01: substituted markers take zero advance at emission, so alignment
        // measures the emitted atoms instead of the unmapped advances.
        double lineWidth = string.Equals(emissionText, run.Text, StringComparison.Ordinal)
            ? MeasureRenderedText(embedded, run.Text, run.FontSize, run.CharacterSpacing, run.KerningEnabled)
            : glyphs.Sum(atom => atom.Advance + atom.AdjustmentBefore);
        double x = run.Alignment switch
        {
            TextAlignment.Center => run.X + Math.Max(0, run.Width - lineWidth) / 2d,
            TextAlignment.Right => run.X + Math.Max(0, run.Width - lineWidth),
            _ => run.X
        };
        double pdfFontSize = OfficePdfTextEmissionProfile.FontSize(run.FontSize);
        double pdfCharacterSpacing = run.CharacterSpacing;
        string? positioningArray = EncodeGlyphPositioningArray(embedded, glyphs, run.FontSize, pdfFontSize, pdfCharacterSpacing, forcePositioningArray: true);
        return new TextGlyphRun(run, resourceName, embedded, glyphHex, positioningArray, glyphs, x, baselineY, lineWidth, pdfFontSize, pdfCharacterSpacing, syntheticBold, syntheticItalic);

        IReadOnlyList<TextGlyphAtom> BuildTextGlyphAtoms()
        {
            var atoms = new List<TextGlyphAtom>();
            ushort previousGlyph = 0;
            foreach (Rune rune in run.Text.EnumerateRunes())
            {
                // RV01: runes with no glyph in the emitting face render as question
                // mark (diagnosed at preparation) with zero advance and no kerning,
                // matching layout, which skips them; layout constructs keep their
                // pass-through exactly as before.
                ushort glyph = embedded.Font.MapCodePoint(rune.Value);
                bool substituted = false;
                if (glyph == 0 && !PdfFallbackFont.IsNonRenderedControl(rune))
                {
                    glyph = embedded.Font.MapCodePoint(0x3F);
                    substituted = glyph != 0;
                }

                if (glyph == 0)
                {
                    continue;
                }

                double adjustmentBefore = 0d;
                if (atoms.Count > 0)
                {
                    adjustmentBefore += run.CharacterSpacing;
                    if (!substituted && run.KerningEnabled && previousGlyph != 0)
                    {
                        adjustmentBefore += embedded.Font.GetKerning(previousGlyph, glyph) * run.FontSize / embedded.Font.UnitsPerEm;
                    }
                }

                double advance = substituted ? 0d : embedded.Font.GetAdvanceWidth(glyph) * run.FontSize / embedded.Font.UnitsPerEm;
                atoms.Add(new TextGlyphAtom(substituted ? 0x3F : rune.Value, run.FontFamily, PptxGlyphTypefaceResolutionSource.Primary, glyph, advance, adjustmentBefore));
                previousGlyph = substituted ? (ushort)0 : glyph;
            }

            return atoms;
        }
    }

    private static TextGlyphRun? BuildTextGlyphRun(string resourceName, PdfEmbeddedFont embedded, PptxPositionedTextSpan span, bool syntheticBold, bool syntheticItalic)
    {
        TextRun run = span.Run;
        string glyphHex = EncodeGlyphHex(embedded, span.GlyphSpan);
        double baselineY = run.Y + run.BaselineOffset;
        if (glyphHex.Length == 0 || !BaselineIntersectsClip(run, baselineY))
        {
            return null;
        }

        double lineWidth = span.GlyphSpan.NaturalWidth;
        double x = run.Alignment switch
        {
            TextAlignment.Center => run.X + Math.Max(0, run.Width - lineWidth) / 2d,
            TextAlignment.Right => run.X + Math.Max(0, run.Width - lineWidth),
            _ => run.X
        };
        double pdfFontSize = OfficePdfTextEmissionProfile.FontSize(run.FontSize);
        double pdfCharacterSpacing = span.PdfCharacterSpacingOverride
            ?? span.GlyphSpan.CharacterSpacing;
        string? positioningArray = EncodeGlyphPositioningArray(embedded, span.GlyphSpan, pdfFontSize, pdfCharacterSpacing, forcePositioningArray: true);
        IReadOnlyList<TextGlyphAtom> glyphs = span.GlyphSpan.Glyphs
            .Select(glyph => new TextGlyphAtom(glyph.CodePoint, glyph.Typeface, glyph.TypefaceResolutionSource, glyph.GlyphId, glyph.Advance, glyph.AdjustmentBefore))
            .ToArray();
        return new TextGlyphRun(run, resourceName, embedded, glyphHex, positioningArray, glyphs, x, baselineY, lineWidth, pdfFontSize, pdfCharacterSpacing, syntheticBold, syntheticItalic);


    }

    // RV01 residual: estimator-unresolvable runes render the prepared question mark
    // instead of vanishing (diagnosed at preparation). Resolved through the same
    // ToUnicode-backed map the extractor reads.
    private static bool TryGetMarkerQuestionCid(PdfEmbeddedFont embedded, PptxTextGlyphLayout glyph, out ushort cid)
    {
        cid = 0;
        if (glyph.GlyphId != 0 || glyph.TypefaceResolutionSource != PptxGlyphTypefaceResolutionSource.Unresolved)
        {
            return false;
        }

        foreach (KeyValuePair<ushort, int> unicodeEntry in embedded.UnicodeByCid)
        {
            if (unicodeEntry.Value == 0x3F && unicodeEntry.Key != 0)
            {
                cid = unicodeEntry.Key;
                return true;
            }
        }

        return false;
    }

    private static string EncodeGlyphHex(PdfEmbeddedFont embedded, PptxTextGlyphSpanLayout span)
    {
        var builder = new StringBuilder(span.Glyphs.Count * 4);
        foreach (PptxTextGlyphLayout glyph in span.Glyphs)
        {
            if (TryGetMarkerQuestionCid(embedded, glyph, out ushort markerCid))
            {
                builder.Append(markerCid.ToString("X4", CultureInfo.InvariantCulture));
                continue;
            }

            if (embedded.TryGetEncodedCid(glyph.GlyphId, out ushort cid))
            {
                builder.Append(cid.ToString("X4", CultureInfo.InvariantCulture));
            }
        }

        return builder.ToString();
    }

    private static string? EncodeGlyphPositioningArray(PdfEmbeddedFont embedded, PptxTextGlyphSpanLayout span, double pdfFontSize, double pdfCharacterSpacing, bool forcePositioningArray)
    {
        if (span.Glyphs.Count == 0)
        {
            return null;
        }

        bool hasPositioning = false;
        var builder = new StringBuilder("[");
        for (int i = 0; i < span.Glyphs.Count; i++)
        {
            PptxTextGlyphLayout glyph = span.Glyphs[i];
            if (TryGetMarkerQuestionCid(embedded, glyph, out ushort markerCid))
            {
                builder.Append('<').Append(markerCid.ToString("X4", CultureInfo.InvariantCulture)).Append('>');
                continue;
            }

            if (i > 0)
            {
                PptxTextGlyphLayout previousGlyph = span.Glyphs[i - 1];
                double adjustmentBefore = PdfTextAdjustmentBefore(
                    glyph.AdjustmentBefore - pdfCharacterSpacing,
                    previousGlyph.Advance,
                    span.FontSize,
                    pdfFontSize);
                double adjustment = pdfFontSize <= 0d ? 0d : -adjustmentBefore * 1000d / pdfFontSize;
                if (Math.Abs(adjustment) > PptxTextMetricRules.TextStateTolerance)
                {
                    builder.Append(' ').Append(adjustment.ToString("0.###", CultureInfo.InvariantCulture)).Append(' ');
                    hasPositioning = true;
                }
            }

            if (embedded.TryGetEncodedCid(glyph.GlyphId, out ushort cid))
            {
                builder.Append('<').Append(cid.ToString("X4", CultureInfo.InvariantCulture)).Append('>');
            }
        }

        builder.Append(']');
        return hasPositioning || forcePositioningArray ? builder.ToString() : null;
    }

    private static string? EncodeGlyphPositioningArray(PdfEmbeddedFont embedded, IReadOnlyList<TextGlyphAtom> glyphs, double layoutFontSize, double pdfFontSize, double characterSpacing, bool forcePositioningArray)
    {
        if (glyphs.Count == 0)
        {
            return null;
        }

        bool hasPositioning = false;
        var builder = new StringBuilder("[");
        for (int i = 0; i < glyphs.Count; i++)
        {
            TextGlyphAtom glyph = glyphs[i];
            if (i > 0)
            {
                TextGlyphAtom previousGlyph = glyphs[i - 1];
                double adjustmentBefore = PdfTextAdjustmentBefore(
                    glyph.AdjustmentBefore - characterSpacing,
                    previousGlyph.Advance,
                    layoutFontSize,
                    pdfFontSize);
                double adjustment = pdfFontSize <= 0d ? 0d : -adjustmentBefore * 1000d / pdfFontSize;
                if (Math.Abs(adjustment) > PptxTextMetricRules.TextStateTolerance)
                {
                    builder.Append(' ').Append(adjustment.ToString("0.###", CultureInfo.InvariantCulture)).Append(' ');
                    hasPositioning = true;
                }
            }

            if (embedded.TryGetEncodedCid(glyph.GlyphId, out ushort cid))
            {
                builder.Append('<').Append(cid.ToString("X4", CultureInfo.InvariantCulture)).Append('>');
            }
        }

        builder.Append(']');
        return hasPositioning || forcePositioningArray ? builder.ToString() : null;
    }

    private static double PdfTextAdjustmentBefore(double layoutAdjustmentBefore, double previousLayoutAdvance, double layoutFontSize, double pdfFontSize)
    {
        if (layoutFontSize <= 0d || pdfFontSize <= 0d)
        {
            return layoutAdjustmentBefore;
        }

        double previousPdfAdvance = previousLayoutAdvance * pdfFontSize / layoutFontSize;
        return layoutAdjustmentBefore + previousLayoutAdvance - previousPdfAdvance;
    }

    private static void DrawGlyphText(PdfGraphicsBuilder graphics, TextGlyphRun glyphRun)
    {
        // Translucent text paints as one glyph path so overlapping contours do not
        // double the alpha; invisible text over it keeps the run extractable.
        bool outlinePath = ShouldDrawGlyphOutlinePath(glyphRun);
        if (outlinePath)
        {
            DrawGlyphOutlinePath();
        }

        TextRun run = glyphRun.Source;
        int renderingMode = outlinePath ? 3 : TextRenderingMode(glyphRun);
        double pdfFontSize = glyphRun.PdfFontSize;
        if (glyphRun.PositioningArray is null)
        {
            graphics.DrawGlyphText(
                glyphRun.ResourceName,
                pdfFontSize,
                glyphRun.X,
                glyphRun.BaselineY,
                run.Color.Red,
                run.Color.Green,
                run.Color.Blue,
                glyphRun.GlyphHex,
                glyphRun.SyntheticItalic,
                characterSpacing: glyphRun.PdfCharacterSpacing,
                textRenderingMode: renderingMode,
                strokeRed: TextStrokeColor(glyphRun).Red,
                strokeGreen: TextStrokeColor(glyphRun).Green,
                strokeBlue: TextStrokeColor(glyphRun).Blue,
                strokeWidth: TextStrokeWidth(glyphRun),
                textRotationQuarterTurns: glyphRun.Source.GlyphRotationQuarterTurns);
        }
        else
        {
            graphics.DrawGlyphPositionedText(
                glyphRun.ResourceName,
                pdfFontSize,
                glyphRun.X,
                glyphRun.BaselineY,
                run.Color.Red,
                run.Color.Green,
                run.Color.Blue,
                glyphRun.PositioningArray,
                glyphRun.SyntheticItalic,
                characterSpacing: glyphRun.PdfCharacterSpacing,
                textRenderingMode: renderingMode,
                strokeRed: TextStrokeColor(glyphRun).Red,
                strokeGreen: TextStrokeColor(glyphRun).Green,
                strokeBlue: TextStrokeColor(glyphRun).Blue,
                strokeWidth: TextStrokeWidth(glyphRun),
                textRotationQuarterTurns: glyphRun.Source.GlyphRotationQuarterTurns);
        }

        void DrawGlyphOutlinePath()
        {
            TextRun outlineRun = glyphRun.Source;
            graphics.SetFillRgb(outlineRun.Color.Red, outlineRun.Color.Green, outlineRun.Color.Blue);
            double shear = glyphRun.SyntheticItalic ? PdfGraphicsBuilder.SyntheticItalicShear : 0d;

            double cursorX = glyphRun.X;
            bool hasPath = false;
            for (int i = 0; i < glyphRun.Glyphs.Count; i++)
            {
                TextGlyphAtom glyph = glyphRun.Glyphs[i];
                if (i > 0)
                {
                    cursorX += glyph.AdjustmentBefore;
                }

                if (PdfGlyphOutlinePath.TryAppendGlyphPath(
                    graphics,
                    glyphRun.Font.Font,
                    glyph.GlyphId,
                    cursorX,
                    glyphRun.BaselineY,
                    glyphRun.PdfFontSize,
                    shear))
                {
                    hasPath = true;
                }

                cursorX += glyph.Advance;
            }

            if (hasPath)
            {
                graphics.FillCurrentPath();
            }
        }
    }

    private static bool ShouldDrawGlyphOutlinePath(TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        if (run.Alpha >= 1d - PptxTextMetricRules.TextStateTolerance ||
            run.Outline is not null ||
            glyphRun.Glyphs.Count == 0)
        {
            return false;
        }

        bool hasDrawableGlyph = false;
        foreach (TextGlyphAtom glyph in glyphRun.Glyphs)
        {
            if (!glyphRun.Font.Font.TryReadGlyphOutline(glyph.GlyphId, out var outline) ||
                outline.Contours.Count == 0)
            {
                if (IsAdvanceOnlyGlyph(glyph))
                {
                    continue;
                }

                return false;
            }

            hasDrawableGlyph = true;
        }

        return hasDrawableGlyph;

        bool IsAdvanceOnlyGlyph(TextGlyphAtom glyph)
        {
            return Rune.IsWhiteSpace(new Rune(glyph.CodePoint));
        }
    }

    private static int TextRenderingMode(TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        if (run.Outline is null)
        {
            return glyphRun.SyntheticBold ? 2 : 0;
        }

        return run.Alpha <= PptxTextMetricRules.TextStateTolerance ? 1 : 2;
    }

    private static RgbColor TextStrokeColor(TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        return run.Outline?.Color ?? run.Color;
    }

    private static double TextStrokeWidth(TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        return run.Outline?.Width ?? PptxTextMetricRules.SyntheticBoldStrokeWidth(run.FontSize);
    }

    private static double MeasureRenderedText(PdfEmbeddedFont embedded, string text, double fontSize, double characterSpacing, bool kerningEnabled)
    {
        double width = embedded.MeasureTextPoints(text, fontSize, kerningEnabled);
        int runeCount = text.EnumerateRunes().Count();
        return Math.Max(0d, width + Math.Max(0, runeCount - 1) * characterSpacing);
    }
}
