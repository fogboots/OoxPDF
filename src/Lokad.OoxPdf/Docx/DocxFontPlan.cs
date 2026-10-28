using Lokad.OoxPdf.Fonts;
using System.Text;

namespace Lokad.OoxPdf.Docx;

internal enum DocxTypefaceResolutionSource
{
    Primary,
    FontTableAlternate,
    Theme,
    ResolverFallback,
    Missing
}

internal sealed record DocxResolvedRunTypeface(
    DocxTextRun Run,
    IReadOnlyList<string> CandidateFamilies,
    string? RequestedFamily,
    string? ResolvedFamily,
    DocxTypefaceResolutionSource Source,
    FontFaceResolution? Resolution);

internal sealed record DocxFontPlan(IReadOnlyList<DocxResolvedRunTypeface> Runs)
{
    public static DocxFontPlan Create(DocxDocument document, IFontResolver fontResolver, CancellationToken cancellationToken, bool includeReviewSources = false)
    {
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<DocxTextRun> runs = DocxBlockTraversal.EnumerateSourceParagraphs(document)
            .SelectMany(paragraph => GetParagraphFontRuns(paragraph, includeReviewSources))
            // Body column-break marks consume a font-dependent pitch after the
            // frame turn, even though their empty text never emits glyphs.
            .Concat(document.BodyElements.OfType<DocxManualBreakElement>()
                .Where(element => element.Value?.Equals("column", StringComparison.OrdinalIgnoreCase) == true)
                .SelectMany(element => element.BreakParagraph is { } paragraph ? GetParagraphFontRuns(paragraph) : []))
            .Concat(document.BodyElements
                .OfType<DocxImplicitParagraphElement>()
                .Select(element => DocxImplicitParagraphElement.CreateParagraphMarkRun(element.MarkFontSizePoints)))
            .ToArray();

        var resolvedRuns = new DocxResolvedRunTypeface[runs.Count];
        DocxResolvedRunTypeface ResolveRunTypeface(DocxTextRun run, DocxFontCatalog fontCatalog)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxEffectiveRunProperties effective = run.EffectiveProperties;
            DocxTypefaceCandidates candidates = DocxFontResolver.ResolveTypeface(run, fontCatalog);
            IReadOnlyList<string> DistinctFamilies(params string?[] families)
            {
                return families
                    .OfType<string>()
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            }
    
            IReadOnlyList<string> families = DistinctFamilies(candidates.Primary, candidates.Alternate, candidates.Theme);
            if (families.Count == 0)
            {
                FontFaceResolution resolution = DocxFontFallbackRules.ResolveDefaultDocumentTypeface(fontResolver, effective.Bold, effective.Italic);
                return new DocxResolvedRunTypeface(
                    run,
                    families,
                    DocxFontFallbackRules.DefaultDocumentTypefaceRequest,
                    resolution.FamilyName,
                    DocxTypefaceResolutionSource.ResolverFallback,
                    resolution);
            }
    
            for (int i = 0; i < families.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string family = families[i];
                FontFaceResolution resolution = fontResolver.Resolve(new FontRequest(family, effective.Bold, effective.Italic));
                if (!resolution.IsFallback)
                {
                    DocxTypefaceResolutionSource SourceForCandidate(DocxTypefaceCandidates candidates, string family)
                    {
                        if (EqualsCandidate(candidates.Primary, family))
                        {
                            return DocxTypefaceResolutionSource.Primary;
                        }
                
                        if (EqualsCandidate(candidates.Alternate, family))
                        {
                            return DocxTypefaceResolutionSource.FontTableAlternate;
                        }
                
                        return DocxTypefaceResolutionSource.Theme;
                    }
    
                    return new DocxResolvedRunTypeface(run, families, family, resolution.FamilyName, SourceForCandidate(candidates, family), resolution);
                }
            }
    
            string requested = families[0];
            FontFaceResolution fallback = fontResolver.Resolve(new FontRequest(requested, effective.Bold, effective.Italic));
            return new DocxResolvedRunTypeface(run, families, requested, fallback.FamilyName, DocxTypefaceResolutionSource.ResolverFallback, fallback);
        }

        for (int i = 0; i < runs.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            resolvedRuns[i] = ResolveRunTypeface(runs[i], document.FontCatalog);
        }

        return new DocxFontPlan(resolvedRuns);
    }

    private static bool EqualsCandidate(string? candidate, string family)
    {
        return candidate is not null && candidate.Equals(family, StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<DocxTextRun> GetParagraphFontRuns(DocxParagraph paragraph, bool includeReviewSources = false)
    {
        HashSet<DocxTextRun>? pictureOwners = includeReviewSources && paragraph.Images.Any(image => image.SourceRunStyle is not null)
            ? new(ReferenceEqualityComparer.Instance) : null;
        foreach (DocxTextRun run in paragraph.Runs)
        {
            pictureOwners?.Add(run);
            yield return run;
        }

        // Glyphless picture rows still measure their paragraph-mark strut.
        // Its font participates in layout even when it paints no glyphs.
        if ((includeReviewSources || DocxPlainBodyFrame.IsPlainPictureOnlyParagraph(paragraph)) &&
            paragraph.ParagraphMarkRun is { } markRun)
        {
            yield return markRun;
        }

        if (pictureOwners is not null)
            foreach (var image in paragraph.Images)
                if (image.SourceRunStyle is { } imageStyle && pictureOwners.Add(imageStyle)) yield return imageStyle;

        if (paragraph.ListLabel is not null)
        {
            DocxTextRun? firstRun = paragraph.Runs.FirstOrDefault();
            var labelRun = DocxLayoutEngine.CreateListLabelRun(
                paragraph.ListLabel, firstRun,
                firstRun?.EffectiveProperties.FontSize ?? DocxDefaults.FontSizePoints);
            yield return labelRun;
            if (paragraph.ListLabel.NumberingChange is { } change)
                yield return DocxLayoutEngine.CreatePreviousListLabelRun(labelRun, change.PreviousText);
        }
    }

}

internal sealed class DocxFontPlanTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxTypographicMetricsProvider, IDocxHheaLineGapProvider, IDocxHheaDescenderProvider
{
    private readonly IReadOnlyList<DocxResolvedRunTypeface> runs;
    private readonly FontFaceResolution? fallbackResolution;
    private readonly CancellationToken cancellationToken;
    private readonly Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fonts;
    private readonly IFontResolver? fontResolver;
    private readonly bool resolveSourceAdvances;
    private readonly Dictionary<(string StableId, int FaceIndex, string? PrimaryFamily, bool Bold, bool Italic), IReadOnlyList<OpenTypeFont?>> candidateChains = new();
    private readonly Dictionary<DocxTextRun, DocxResolvedRunTypeface> runsByReference = new(ReferenceEqualityComparer.Instance);
    private Dictionary<DocxTextRun, DocxResolvedRunTypeface>? runsByValue;
    // Layout asks for the same resolution's font many times per line; a reference lookup
    // skips hashing the source id on every call.
    private readonly Dictionary<FontFaceResolution, OpenTypeFont?> fontsByResolution = new(ReferenceEqualityComparer.Instance);

    public DocxFontPlanTextMeasurer(DocxFontPlan plan, FontFaceResolution? fallbackResolution, CancellationToken cancellationToken, IFontResolver? fontResolver = null, Dictionary<(string StableId, int FaceIndex), OpenTypeFont?>? sharedFonts = null, bool resolveSourceAdvances = false)
    {
        runs = plan.Runs;
        foreach (DocxResolvedRunTypeface resolved in plan.Runs)
        {
            runsByReference.TryAdd(resolved.Run, resolved);
        }
        this.fallbackResolution = fallbackResolution;
        this.cancellationToken = cancellationToken;
        this.fontResolver = fontResolver;
        this.resolveSourceAdvances = resolveSourceAdvances;
        fonts = sharedFonts ?? new();
    }

    public bool TryGetUnderlineRuleMetrics(DocxTextRun? run, out double positionEm, out double thicknessEm)
    {
        positionEm = thicknessEm = 0d;
        OpenTypeFont? font = ResolveFont(run);
        if (font is null || font.UnitsPerEm == 0) return false;
        positionEm = (double)font.Post.UnderlinePosition / font.UnitsPerEm;
        thicknessEm = Math.Abs((double)font.Post.UnderlineThickness) / font.UnitsPerEm;
        return true;
    }

    public bool TryGetStrikeoutRuleMetrics(DocxTextRun? run, out double positionEm, out double thicknessEm)
    {
        positionEm = 0d;
        thicknessEm = 0d;
        OpenTypeFont? font = ResolveFont(run);
        if (font is null || font.UnitsPerEm == 0)
        {
            return false;
        }

        positionEm = (double)font.Os2.StrikeoutPosition / font.UnitsPerEm;
        thicknessEm = (double)font.Os2.StrikeoutSize / font.UnitsPerEm;
        return positionEm > 0d && thicknessEm > 0d;
    }

    public bool TryGetSingleLineEm(DocxTextRun? run, out double singleLineEm)
    {
        singleLineEm = 0d;
        OpenTypeFont? font = ResolveFont(run);
        if (font is null || font.UnitsPerEm == 0)
        {
            return false;
        }

        singleLineEm = DocxLineMetrics.MeasureOpenTypeSingleLineHeight(font, 1d);
        return singleLineEm > 0d;
    }

    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if ((resolved?.Resolution ?? fallbackResolution) is not FontFaceResolution resolution)
        {
            return 0d;
        }

        OpenTypeFont? font = LoadFont(resolution);
        if (font is null || font.UnitsPerEm == 0)
        {
            return 0d;
        }

        // Primary-covered text measures identically without the split; skipping
        // candidate loads leaves measurement unchanged (G04).
        if (FontCoverageFallback.IsFullyCovered(text, font, cancellationToken))
        {
            return MeasureSingleFontText(run, text, fontSize, font);
        }

        IReadOnlyList<OpenTypeFont?> candidates = ResolveCandidateFonts(run, resolved, resolution, font);
        if (candidates.Count <= 1)
        {
            return MeasureSingleFontText(run, text, fontSize, font);
        }

        double widthPoints = 0d;
        foreach (FontCoverageSpan span in FontCoverageFallback.SplitByCoverage(text, candidates, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OpenTypeFont? spanFont = span.FontIndex < 0 ? font : candidates[span.FontIndex];
            if (spanFont is null || spanFont.UnitsPerEm == 0)
            {
                continue;
            }

            widthPoints += MeasureSingleFontRawText(text.Substring(span.Start, span.Length), fontSize, spanFont, UsesPairKerning(run));
        }

        return DocxTextSpacing.AddCharacterSpacing(widthPoints, run, text);
    }

    private double MeasureSingleFontText(DocxTextRun? run, string text, double fontSize, OpenTypeFont font)
    {
        double units = 0d;
        ushort previousGlyph = 0;
        bool pairKerningEnabled = UsesPairKerning(run);
        foreach (Rune rune in text.EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            ushort glyph = font.MapCodePoint(rune.Value);
            if (pairKerningEnabled && previousGlyph != 0 && glyph != 0)
            {
                units += font.GetKerning(previousGlyph, glyph);
            }

            units += font.GetAdvanceWidth(glyph);
            previousGlyph = glyph;
        }

        return DocxTextSpacing.AddCharacterSpacing(units * fontSize / font.UnitsPerEm, run, text);
    }

    private bool UsesPairKerning(DocxTextRun? run) => run is null ||
        (resolveSourceAdvances ? DocxTextAdvances.Resolve(run).PairKerningEnabled : run.PairKerningEnabled);

    private double MeasureSingleFontRawText(string text, double fontSize, OpenTypeFont font, bool pairKerningEnabled)
    {
        double units = 0d;
        ushort previousGlyph = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            ushort glyph = font.MapCodePoint(rune.Value);
            if (pairKerningEnabled && previousGlyph != 0 && glyph != 0)
            {
                units += font.GetKerning(previousGlyph, glyph);
            }

            units += font.GetAdvanceWidth(glyph);
            previousGlyph = glyph;
        }

        return units * fontSize / font.UnitsPerEm;
    }

    public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if ((resolved?.Resolution ?? fallbackResolution) is not FontFaceResolution resolution)
        {
            return fontSize;
        }

        OpenTypeFont? font = LoadFont(resolution);
        return font is null
            ? fontSize
            : DocxLineMetrics.MeasureOpenTypeSingleLineHeight(font, fontSize);
    }

    public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if ((resolved?.Resolution ?? fallbackResolution) is not FontFaceResolution resolution)
        {
            return fontSize;
        }

        OpenTypeFont? font = LoadFont(resolution);
        return font is null
            ? fontSize
            : DocxLineMetrics.MeasureHheaLineHeight(font, fontSize);
    }
    public double MeasureTypographicLineHeight(DocxTextRun? run, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if ((resolved?.Resolution ?? fallbackResolution) is not FontFaceResolution resolution)
        {
            return 0d;
        }

        OpenTypeFont? font = LoadFont(resolution);
        return font is null
            ? 0d
            : DocxLineMetrics.MeasureTypographicLineHeight(font, fontSize);
    }

    public double MeasureHheaAscender(DocxTextRun? run, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if ((resolved?.Resolution ?? fallbackResolution) is not FontFaceResolution resolution)
        {
            return 0d;
        }

        OpenTypeFont? font = LoadFont(resolution);
        return font is null
            ? 0d
            : DocxLineMetrics.MeasureHheaAscender(font, fontSize);
    }

    public double MeasureHheaLineGap(DocxTextRun? run, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if ((resolved?.Resolution ?? fallbackResolution) is not FontFaceResolution resolution)
        {
            return 0d;
        }

        OpenTypeFont? font = LoadFont(resolution);
        return font is null
            ? 0d
            : DocxLineMetrics.MeasureHheaLineGap(font, fontSize);
    }

    public double MeasureHheaDescender(DocxTextRun? run, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if ((resolved?.Resolution ?? fallbackResolution) is not FontFaceResolution resolution)
        {
            return 0d;
        }

        OpenTypeFont? font = LoadFont(resolution);
        return font is null
            ? 0d
            : DocxLineMetrics.MeasureHheaDescender(font, fontSize);
    }

    public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OpenTypeFont? font = ResolveFont(run);
        return font is null ? fontSize : DocxLineMetrics.MeasureWindowsAscender(font, fontSize);
    }

    public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OpenTypeFont? font = ResolveFont(run);
        return font is null ? 0d : DocxLineMetrics.MeasureWindowsDescender(font, fontSize);
    }

    public bool UseTypographicMetrics(DocxTextRun? run)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OpenTypeFont? font = ResolveFont(run);
        return font is not null && (font.Os2.SelectionFlags & 0x80) != 0;
    }

    public double MeasureTypographicBaselineInset(DocxTextRun? run, double fontSize)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OpenTypeFont? font = ResolveFont(run);
        return font is null ? 0d : DocxLineMetrics.MeasureTypographicBaselineInset(font, fontSize);
    }

    private IReadOnlyList<OpenTypeFont?> ResolveCandidateFonts(DocxTextRun? run, DocxResolvedRunTypeface? resolved, FontFaceResolution resolution, OpenTypeFont primary)
    {
        bool bold = run is not null && run.Bold;
        bool italic = run is not null && run.Italic;
        string? primaryFamily = resolved?.ResolvedFamily;
        var key = (resolution.Source.StableId, resolution.FontFaceIndex, primaryFamily, bold, italic);
        if (candidateChains.TryGetValue(key, out IReadOnlyList<OpenTypeFont?>? cached))
        {
            return cached;
        }

        var candidates = new List<OpenTypeFont?>();
        if (fontResolver is null)
        {
            candidates.Add(primary);
        }
        else
        {
            foreach (FontFaceResolution candidate in DocxFontFallbackRules.ResolveCandidateResolutions(fontResolver, run, primaryFamily, resolution, fallbackResolution))
            {
                cancellationToken.ThrowIfCancellationRequested();
                OpenTypeFont? candidateFont = candidates.Count == 0 ? primary : LoadFont(candidate);
                if (candidateFont is not null && !candidates.Contains(candidateFont))
                {
                    candidates.Add(candidateFont);
                }
            }
        }

        IReadOnlyList<OpenTypeFont?> chain = candidates;
        candidateChains[key] = chain;
        return chain;
    }

    private DocxResolvedRunTypeface? ResolveRun(DocxTextRun? run)
    {
        if (run is null)
        {
            return null;
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (runsByReference.TryGetValue(run, out DocxResolvedRunTypeface? fast))
        {
            return fast;
        }

        // Cloned runs miss the reference-identity index and match the first plan run with
        // equal values. Memoize confirmed matches so each distinct clone looks up once.
        if (runsByValue is null)
        {
            runsByValue = new Dictionary<DocxTextRun, DocxResolvedRunTypeface>();
            foreach (DocxResolvedRunTypeface candidate in runs)
            {
                runsByValue.TryAdd(candidate.Run, candidate);
            }
        }

        if (runsByValue.TryGetValue(run, out DocxResolvedRunTypeface? resolved))
        {
            runsByReference[run] = resolved;
            return resolved;
        }

        return null;
    }

    private OpenTypeFont? ResolveFont(DocxTextRun? run)
    {
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        return (resolved?.Resolution ?? fallbackResolution) is FontFaceResolution resolution
            ? LoadFont(resolution)
            : null;
    }

    private OpenTypeFont? LoadFont(FontFaceResolution resolution)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (fontsByResolution.TryGetValue(resolution, out OpenTypeFont? known))
        {
            return known;
        }

        var key = (resolution.Source.StableId, resolution.FontFaceIndex);
        if (!fonts.TryGetValue(key, out OpenTypeFont? loaded))
        {
            loaded = FontProgramLoader.Load(resolution, cancellationToken);
            fonts[key] = loaded;
        }

        fontsByResolution[resolution] = loaded;
        return loaded;
    }
}

internal sealed record DocxFontPlanSnapshot(
    int RunCount,
    int PrimaryCount,
    int FontTableAlternateCount,
    int ThemeCount,
    int ResolverFallbackCount,
    int MissingCount,
    int DistinctCandidateFamilyCount,
    int DistinctResolvedFamilyCount,
    IReadOnlyList<DocxFontMetricBucketSnapshot> MetricBuckets)
{
    public static DocxFontPlanSnapshot FromPlan(DocxFontPlan plan)
    {
        IReadOnlyList<DocxFontMetricBucketSnapshot> ToMetricBuckets(DocxFontPlan plan)
        {
            string? HashFamily(string? family)
            {
                if (string.IsNullOrWhiteSpace(family))
                {
                    return null;
                }
        
                byte[] bytes = Encoding.UTF8.GetBytes(family.ToUpperInvariant());
                byte[] hash = System.Security.Cryptography.SHA256.HashData(bytes);
                return Convert.ToHexString(hash).Substring(0, 12).ToLowerInvariant();
            }
    
            return plan.Runs
                .GroupBy(run => new
                {
                    run.Source,
                    FontSize = Math.Round(run.Run.EffectiveProperties.FontSize, 3),
                    ResolvedFamilyHash = HashFamily(run.ResolvedFamily)
                })
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key.Source.ToString(), StringComparer.Ordinal)
                .ThenBy(group => group.Key.FontSize)
                .Select(group => ToMetricBucketSnapshot(
                    group.Key.Source.ToString(),
                    group.Key.FontSize,
                    group.Key.ResolvedFamilyHash,
                    group.Count(),
                    group.FirstOrDefault()?.Resolution))
                .ToArray();
        }

        return new DocxFontPlanSnapshot(
            plan.Runs.Count,
            Count(plan, DocxTypefaceResolutionSource.Primary),
            Count(plan, DocxTypefaceResolutionSource.FontTableAlternate),
            Count(plan, DocxTypefaceResolutionSource.Theme),
            Count(plan, DocxTypefaceResolutionSource.ResolverFallback),
            Count(plan, DocxTypefaceResolutionSource.Missing),
            plan.Runs
                .SelectMany(run => run.CandidateFamilies)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count(),
            plan.Runs
                .Select(run => run.ResolvedFamily)
                .Where(family => !string.IsNullOrWhiteSpace(family))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count(),
            ToMetricBuckets(plan));
    }

    private static int Count(DocxFontPlan plan, DocxTypefaceResolutionSource source)
    {
        return plan.Runs.Count(run => run.Source == source);
    }

    private static DocxFontMetricBucketSnapshot ToMetricBucketSnapshot(
        string source,
        double fontSize,
        string? resolvedFamilyHash,
        int runCount,
        FontFaceResolution? resolution)
    {
        OpenTypeFont? font = FontProgramLoader.Load(resolution, CancellationToken.None);
        if (font is null)
        {
            return new DocxFontMetricBucketSnapshot(
                source,
                fontSize,
                resolvedFamilyHash,
                runCount,
                UnitsPerEm: null,
                TypographicAscender: null,
                TypographicDescender: null,
                TypographicLineGap: null,
                WindowsAscender: null,
                WindowsDescender: null,
                SingleLineHeightPoints: null,
                WindowsAscenderPoints: null,
                WindowsDescenderPoints: null);
        }

        return new DocxFontMetricBucketSnapshot(
            source,
            fontSize,
            resolvedFamilyHash,
            runCount,
            font.UnitsPerEm,
            font.Os2.TypographicAscender,
            font.Os2.TypographicDescender,
            font.Os2.TypographicLineGap,
            font.Os2.WindowsAscender,
            font.Os2.WindowsDescender,
            DocxLineMetrics.MeasureOpenTypeSingleLineHeight(font, fontSize),
            DocxLineMetrics.MeasureWindowsAscender(font, fontSize),
            DocxLineMetrics.MeasureWindowsDescender(font, fontSize));
    }

}

internal sealed record DocxFontMetricBucketSnapshot(
    string Source,
    double FontSize,
    string? ResolvedFamilyHash,
    int RunCount,
    int? UnitsPerEm,
    int? TypographicAscender,
    int? TypographicDescender,
    int? TypographicLineGap,
    int? WindowsAscender,
    int? WindowsDescender,
    double? SingleLineHeightPoints,
    double? WindowsAscenderPoints,
    double? WindowsDescenderPoints);
