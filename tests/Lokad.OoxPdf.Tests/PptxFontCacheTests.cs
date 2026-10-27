using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxFontCacheTests
{
    public static void PptxFallbackOrderReusesStyleSnapshotWithoutChangingPriority()
    {
        FontFaceResolution regular = FontFace("Regular", "regular", TestFontBuilder.CreateTestFont());
        FontFaceResolution bold = FontFace("Bold", "bold", TestFontBuilder.CreateTestFont()) with
        { Style = new FontStyleKey(Bold: true, WeightClass: 700) };
        var resolver = new PresentationFontResolver(new CatalogResolver(bold, regular));
        var regularOrder = resolver.GetFallbackFonts(false, false);
        var boldOrder = resolver.GetFallbackFonts(true, false);
        TestAssert.True(ReferenceEquals(regular, regularOrder[0]), "Regular style must retain its priority.");
        TestAssert.True(ReferenceEquals(bold, boldOrder[0]), "Bold style must retain its priority.");
        TestAssert.True(ReferenceEquals(regularOrder, resolver.GetFallbackFonts(false, false)), "Repeated glyph lookups reuse the same bounded ordering.");
        TestAssert.True(ReferenceEquals(boldOrder, resolver.GetFallbackFonts(true, false)), "Styles have separate cached orderings.");
        TestAssert.Equal(2, regularOrder.Count);
        TestAssert.Equal(2, boldOrder.Count);
    }

    public static void PptxEstimatorsSharingResolverReuseLoadedPrograms()
    {
        var program = new MemoryFontProgramSource("cache-test", TestFontBuilder.CreateTestFont());
        var resolver = new PresentationFontResolver(new SingleProgramResolver(program));
        Type estimatorType = typeof(PptxRenderer).GetNestedType(
            "TextAdvanceEstimator",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected text advance estimator.");
        object first = System.Activator.CreateInstance(estimatorType, resolver, CancellationToken.None) ?? throw new InvalidOperationException("Expected estimator instance.");
        object second = System.Activator.CreateInstance(estimatorType, resolver, CancellationToken.None) ?? throw new InvalidOperationException("Expected estimator instance.");
        System.Reflection.MethodInfo resolve = estimatorType.GetMethod("ResolveOpenTypeFont") ?? throw new InvalidOperationException("Expected program resolution bridge.");

        object? firstFont = resolve.Invoke(first, ["TestFont", false, false]);
        object? secondFont = resolve.Invoke(second, ["TestFont", false, false]);

        TestAssert.True(firstFont is not null, "The test program should resolve.");
        TestAssert.True(ReferenceEquals(firstFont, secondFont), "Estimators sharing a resolver must reuse the loaded program instead of reloading per estimator.");
    }

    public static void PptxSlideRenderLoadsEachFontProgramOnce()
    {
        byte[] fontBytes = TestFontBuilder.CreateTestFont();
        var counters = new System.Collections.Concurrent.ConcurrentDictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
        var resolver = new PerFamilyCountingResolver(fontBytes, counters);
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="2" name="First"/><p:nvPr/></p:nvSpPr>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="TestFont"/></a:rPr><a:t>Alpha</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="3" name="Second"/><p:nvPr/></p:nvSpPr>
                      <p:spPr><a:xfrm><a:off x="914400" y="2743200"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="TestFont"/></a:rPr><a:t>Beta</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    <p:graphicFrame>
                      <p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:nvPr/></p:nvGraphicFramePr>
                      <p:xfrm><a:off x="914400" y="4572000"/><a:ext cx="3657600" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="TestFont"/></a:rPr><a:t>Gamma</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="TestFont"/></a:rPr><a:t>Delta</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        int pageCount;
        using (System.IO.FileStream stream = System.IO.File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, System.Threading.CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, System.Threading.CancellationToken.None);
            pageCount = new PptxRenderer(resolver).RenderPages(document, package, diagnosticSink: null, System.Threading.CancellationToken.None).Count();
        }

        TestAssert.Equal(1, pageCount);
        TestAssert.Equal(1, counters.Count);
        TestAssert.Equal("counting:TestFont", counters.Keys.Single());
        TestAssert.Equal(1, counters.Values.Single());
    }

    private sealed class PerFamilyCountingResolver(byte[] fontBytes, System.Collections.Concurrent.ConcurrentDictionary<string, int> counters) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            string family = string.IsNullOrWhiteSpace(request.FamilyName) ? "(default)" : request.FamilyName;
            var source = new CountingProgramSource("counting:" + family, fontBytes, counters);
            return new FontFaceResolution(
                request.FamilyName,
                family,
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                source,
                IsFallback: false);
        }

        private sealed class CountingProgramSource(string stableId, byte[] bytes, System.Collections.Concurrent.ConcurrentDictionary<string, int> counters) : IFontProgramSource
        {
            public string StableId => stableId;

            public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct)
            {
                counters.AddOrUpdate(stableId, 1, (_, count) => count + 1);
                ct.ThrowIfCancellationRequested();
                return ValueTask.FromResult((ReadOnlyMemory<byte>)bytes);
            }
        }
    }
    public static void PptxSharedResolverDeduplicatesIdenticalSubsets()
    {
        var program = new MemoryFontProgramSource("cache-test", TestFontBuilder.CreateTestFont());
        var resolver = new PresentationFontResolver(new SingleProgramResolver(program));
        var resolved = resolver.ResolvePresentationOpenTypeFont(new FontRequest("TestFont", false, false), CancellationToken.None) ?? throw new InvalidOperationException("Expected test program resolution.");

        PdfEmbeddedFont first = resolver.GetOrCreateSubset(resolved.Font, [65, 66, 65], CancellationToken.None);
        PdfEmbeddedFont same = resolver.GetOrCreateSubset(resolved.Font, [66, 65], CancellationToken.None);
        PdfEmbeddedFont other = resolver.GetOrCreateSubset(resolved.Font, [67], CancellationToken.None);

        TestAssert.True(ReferenceEquals(first, same), "Identical codepoint sets must share one subset regardless of order or duplicates.");
        TestAssert.True(!ReferenceEquals(first, other), "Different codepoint sets must keep distinct subsets.");
    }
    public static void CaseDistinctStableIdsDoNotAliasPrograms()
    {
        // T01: StableIds are opaque resolver-scoped identities compared ordinally.
        // Two custom IDs differing only by case must load separately (the PPTX program
        // cache previously used OrdinalIgnoreCase, aliasing them like family names).
        byte[] fontBytes = TestFontBuilder.CreateTestFont();
        var upper = new MemoryFontProgramSource("Case:Font", fontBytes);
        var lower = new MemoryFontProgramSource("case:font", fontBytes);
        var resolver = new PresentationFontResolver(null);
        OpenTypeFont? first = resolver.GetOrLoadOpenTypeFont(
            new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(false, false, 400, 0, false), upper, IsFallback: false),
            CancellationToken.None);
        OpenTypeFont? second = resolver.GetOrLoadOpenTypeFont(
            new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(false, false, 400, 0, false), lower, IsFallback: false),
            CancellationToken.None);

        TestAssert.NotNull(first);
        TestAssert.NotNull(second);
        TestAssert.True(!ReferenceEquals(first, second), "Case-distinct program IDs must load separately, not alias.");
    }

    public static void CollectionFaceIndexesCacheSeparately()
    {
        // T01: multiple TTC faces of one program resolve consistently through the
        // ordinal (StableId, FaceIndex) program key.
        byte[] collection = TestFontBuilder.CreateCollection("TestFontA", "TestFontB");
        TestAssert.Equal(2, OpenTypeFont.GetCollectionFontCount(collection));
        var source = new MemoryFontProgramSource("collection:test", collection);
        var resolver = new PresentationFontResolver(null);
        OpenTypeFont? first = resolver.GetOrLoadOpenTypeFont(
            new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(false, false, 400, 0, false), source, IsFallback: false),
            CancellationToken.None);
        OpenTypeFont? second = resolver.GetOrLoadOpenTypeFont(
            new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(false, false, 400, 1, false), source, IsFallback: false),
            CancellationToken.None);

        TestAssert.NotNull(first);
        TestAssert.NotNull(second);
        TestAssert.True(!ReferenceEquals(first, second), "Distinct collection faces must keep distinct cache entries.");
        TestAssert.True(ReferenceEquals(first, resolver.GetOrLoadOpenTypeFont(
            new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(false, false, 400, 0, false), source, IsFallback: false),
            CancellationToken.None)), "Repeated face lookups must reuse the cached program.");
    }

    public static void StyleVariantsShareProgramAndSubsetConsistently()
    {
        // T01: request style differences keep distinct resolutions but deterministically
        // share one program and one subset per (program, codepoint set).
        var program = new MemoryFontProgramSource("style-test", TestFontBuilder.CreateTestFont());
        var resolver = new PresentationFontResolver(new SingleProgramResolver(program));
        var plain = resolver.ResolvePresentationOpenTypeFont(new FontRequest("TestFont", false, false), CancellationToken.None)
            ?? throw new InvalidOperationException("Expected test program resolution.");
        var bold = resolver.ResolvePresentationOpenTypeFont(new FontRequest("TestFont", true, false), CancellationToken.None)
            ?? throw new InvalidOperationException("Expected test program resolution.");

        TestAssert.True(ReferenceEquals(plain.Font, bold.Font), "Style variants of one program must load once.");
        PdfEmbeddedFont first = resolver.GetOrCreateSubset(plain.Font, [65, 66], CancellationToken.None);
        PdfEmbeddedFont second = resolver.GetOrCreateSubset(bold.Font, [66, 65, 66], CancellationToken.None);

        TestAssert.True(ReferenceEquals(first, second), "Same program and codepoint set must share one subset across style variants.");
    }

    public static void PptxFallbackPreservesCaseDistinctProgramIdentities()
    {
        FontFaceResolution primary = FontFace("Primary", "Case:Font", TestFontBuilder.CreateTestFont());
        FontFaceResolution fallback = FontFace("Fallback", "case:font", CreateDifferentCoverageFont());
        var resolver = new PresentationFontResolver(new CatalogResolver(primary, fallback));
        Type estimatorType = typeof(PptxRenderer).GetNestedType("TextAdvanceEstimator", System.Reflection.BindingFlags.NonPublic)!;
        object estimator = Activator.CreateInstance(estimatorType, resolver, CancellationToken.None)!;
        object? glyph = estimatorType.GetMethod("ResolveGlyphFont")!.Invoke(estimator, ["Primary", false, false, 0x100]);
        TestAssert.NotNull(glyph);
        TestAssert.Equal("Fallback", (string)glyph!.GetType().GetProperty("Typeface")!.GetValue(glyph)!);
        TestAssert.True(!resolver.IsPrimaryProgram(primary, fallback), "Case-distinct opaque IDs must remain eligible.");
    }

    public static void PptxIndependentFontOwnersDoNotAliasProgramsOrSubsets()
    {
        FontFaceResolution primary = FontFace("Primary", "same-id", TestFontBuilder.CreateTestFont());
        FontFaceResolution fallback = FontFace("Fallback", "same-id", CreateDifferentCoverageFont());
        var resolver = new PresentationFontResolver(new SingleProgramResolver((MemoryFontProgramSource)primary.Source));
        OpenTypeFont first = resolver.GetOrLoadOpenTypeFont(primary, CancellationToken.None)!;
        OpenTypeFont second = resolver.GetOrLoadOpenTypeFont(fallback, CancellationToken.None, fromCatalog: true)!;
        TestAssert.Equal((ushort)0, first.MapCodePoint(0x100));
        TestAssert.True(second.MapCodePoint(0x100) != 0, "An independent catalog must load its own program despite an equal ID.");
        TestAssert.True(!resolver.IsPrimaryProgram(primary, fallback), "Independent owners cannot exclude each other's IDs.");
        TestAssert.True(ReferenceEquals(second, resolver.GetOrLoadOpenTypeFont(fallback, CancellationToken.None, fromCatalog: true)), "Catalog repeats must reuse their own parsed program.");
        PdfEmbeddedFont firstSubset = resolver.GetOrCreateSubset(first, [0x100], CancellationToken.None);
        PdfEmbeddedFont secondSubset = resolver.GetOrCreateSubset(second, [0x100], CancellationToken.None);
        TestAssert.True(!ReferenceEquals(firstSubset, secondSubset), "Different programs must not share a subset through an equal source ID.");
        TestAssert.True(ReferenceEquals(secondSubset, resolver.GetOrCreateSubset(second, [0x100, 0x100], CancellationToken.None)), "Repeated sets on one program must reuse the subset.");
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => resolver.GetOrLoadOpenTypeFont(fallback, canceled.Token, fromCatalog: true));
        TestAssert.Throws<OperationCanceledException>(() => resolver.GetOrCreateSubset(second, [0x100], canceled.Token));
    }

    public static void PptxSharedCatalogReusesPrimaryProgram()
    {
        FontFaceResolution primary = FontFace("Primary", "shared", TestFontBuilder.CreateTestFont());
        var resolver = new PresentationFontResolver(new CatalogResolver(primary, primary));
        OpenTypeFont? first = resolver.GetOrLoadOpenTypeFont(primary, CancellationToken.None);
        TestAssert.NotNull(first);
        TestAssert.True(ReferenceEquals(first, resolver.GetOrLoadOpenTypeFont(primary, CancellationToken.None, fromCatalog: true)), "A primary catalog has one ownership namespace.");
        TestAssert.True(resolver.IsPrimaryProgram(primary, primary), "The same primary/catalog program is excluded from redundant fallback work.");
    }

    private static FontFaceResolution FontFace(string family, string id, byte[] bytes) =>
        new(family, family, new FontStyleKey(), new MemoryFontProgramSource(id, bytes), false);

    private static byte[] CreateDifferentCoverageFont()
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        (int cmap, _) = TestFontBuilder.GetTableRange(bytes, "cmap");
        int format12 = cmap + (int)System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(cmap + 16, 4));
        // Shift the first format-12 range from ASCII to U+0100..U+017F;
        // retain valid glyph numbers, outlines and the other ranges.
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(format12 + 16, 4), 0x100);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(format12 + 20, 4), 0x17F);
        return bytes;
    }

    private sealed class CatalogResolver(FontFaceResolution primary, FontFaceResolution fallback) : IFontResolver, IFontCatalog
    {
        public FontFaceResolution Resolve(FontRequest request) =>
            string.Equals(request.FamilyName, fallback.ResolvedFamily, StringComparison.OrdinalIgnoreCase) ? fallback : primary;

        public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts() => [primary, fallback];
    }

    private sealed class SingleProgramResolver(MemoryFontProgramSource program) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return new FontFaceResolution(
                request.FamilyName,
                "TestFont",
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                program,
                IsFallback: false);
        }
    }
}
