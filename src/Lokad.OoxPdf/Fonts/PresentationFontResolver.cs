using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Fonts;

internal sealed class PresentationFontResolver
{
    private readonly IFontResolver primary;
    private readonly IFontCatalog fontCatalog;
    // The catalog is a conversion snapshot. There are exactly four style
    // orderings; cache them once instead of sorting again for each missing glyph.
    private readonly Dictionary<(bool Bold, bool Italic), FontFaceResolution[]> fallbackOrders = new();
    // IDs are ordinal and scoped to their resolver. A separate discovery catalog
    // must not alias the primary resolver; a primary catalog shares its cache.
    private readonly Dictionary<(bool SeparateCatalog, string StableId, int FaceIndex), OpenTypeFont?> openTypeFonts = new();
    private readonly Dictionary<(OpenTypeFont Font, ulong CodePointsHash), SubsetEntry> subsets = new();

    // RV01: per-conversion record of diagnosed missing-font families and glyph sets,
    // so per-slide and per-chart preparation reports each family once.
    public HashSet<string> ReportedMissingFontDiagnostics { get; } = new(StringComparer.Ordinal);

    // RV01: fallback faces used anywhere in the conversion (slides and chart
    // parts share this resolver), so pages register them without threading lists.
    public HashSet<PdfFallbackFont> UsedFallbackFaces { get; } = new();

    public PresentationFontResolver(IFontResolver? primary)
    {
        this.primary = primary ?? new WindowsFontResolver();
        fontCatalog = this.primary as IFontCatalog ?? new WindowsFontResolver();
    }

    public FontFaceResolution Resolve(FontRequest request)
    {
        return primary.Resolve(request);
    }

    public (FontFaceResolution Resolution, OpenTypeFont Font)? ResolvePresentationOpenTypeFont(FontRequest request, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        FontFaceResolution resolution = Resolve(request);
        OpenTypeFont? font = GetOrLoadOpenTypeFont(resolution, cancellationToken);
        return font is null ? null : (resolution, font);
    }

    public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        return fontCatalog.GetDiscoveredFonts();
    }

    internal IReadOnlyList<FontFaceResolution> GetFallbackFonts(bool bold, bool italic)
    {
        var key = (bold, italic);
        if (!fallbackOrders.TryGetValue(key, out FontFaceResolution[]? ordered))
        {
            ordered = fontCatalog.GetDiscoveredFonts()
                .Where(f => !f.HasMathTable)
                .OrderBy(f => f.Bold == bold ? 0 : 1000)
                .ThenBy(f => f.Italic == italic ? 0 : 1000)
                .ThenBy(f => Math.Abs(f.WeightClass - (bold ? 700 : 400)))
                .ThenBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(f => f.Source.StableId, StringComparer.Ordinal)
                .ThenBy(f => f.FontFaceIndex)
                .ToArray();
            fallbackOrders.Add(key, ordered);
        }
        return ordered;
    }

    internal bool IsPrimaryProgram(FontFaceResolution primaryResolution, FontFaceResolution catalogResolution) =>
        ReferenceEquals(primary, fontCatalog) &&
        primaryResolution.FontFaceIndex == catalogResolution.FontFaceIndex &&
        string.Equals(primaryResolution.Source.StableId, catalogResolution.Source.StableId, StringComparison.Ordinal);

    internal OpenTypeFont? GetOrLoadOpenTypeFont(FontFaceResolution resolution, CancellationToken cancellationToken, bool fromCatalog = false)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var key = (fromCatalog && !ReferenceEquals(primary, fontCatalog), resolution.Source.StableId, resolution.FontFaceIndex);
        if (openTypeFonts.TryGetValue(key, out OpenTypeFont? cached))
        {
            return cached;
        }

        cached = FontProgramLoader.Load(resolution, cancellationToken);
        openTypeFonts[key] = cached;
        return cached;
    }

    // Conversion-scoped identical subsets: repeated slides embed the same glyph
    // sets, so subsetting runs once per (program, codepoint set) instead of once
    // per slide. The subsetter is a pure function of those inputs, so shared
    // instances keep identical bytes, CID maps, and resource keys (G04).
    internal PdfEmbeddedFont GetOrCreateSubset(OpenTypeFont font, IReadOnlyList<int> codePoints, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // T01: the 64-bit set hash is a lookup shortcut, not an identity proof. Stored
        // sets are compared on every hit so a hash collision can never corrupt output;
        // a mismatch falls back to a fresh uncached subset instead of poisoning the cache.
        int[] sorted = SortDeduplicate(codePoints);
        var key = (font, HashSortedCodePoints(sorted));
        if (subsets.TryGetValue(key, out SubsetEntry? cached) && cached.CodePoints.AsSpan().SequenceEqual(sorted))
        {
            return cached.Subset;
        }

        PdfEmbeddedFont created = PdfEmbeddedFont.Create(font, codePoints, cancellationToken);
        if (cached is null)
        {
            subsets[key] = new SubsetEntry(created, sorted);
        }

        return created;
    }

    private sealed record SubsetEntry(PdfEmbeddedFont Subset, int[] CodePoints);

    private static int[] SortDeduplicate(IReadOnlyList<int> codePoints)
    {
        int[] sorted = codePoints.ToArray();
        Array.Sort(sorted);
        int count = 0;
        int previous = 0;
        bool first = true;
        foreach (int codePoint in sorted)
        {
            if (!first && codePoint == previous)
            {
                continue;
            }

            first = false;
            previous = codePoint;
            sorted[count++] = codePoint;
        }

        Array.Resize(ref sorted, count);
        return sorted;
    }
    // R15: one normalized codepoint set feeds both the lookup hash and the hit
    // equality check, so repeated subset requests sort/deduplicate once instead of
    // once for hashing and again for comparison.
    private static ulong HashSortedCodePoints(int[] sorted)
    {
        ulong hash = 14695981039346656037ul;
        foreach (int codePoint in sorted)
        {
            hash ^= (uint)codePoint;
            hash *= 1099511628211ul;
        }

        return hash;
    }
}
