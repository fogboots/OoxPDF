using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class FontFamilyAliasTests
{
    private static readonly (ushort Language, string Name)[] Names = [(0x0404, "字體"), (0x0409, "Review Family"), (0x0411, "字體"), (0x0409, "REVIEW FAMILY")];

    public static void FontFamilyAliasesPreserveCanonicalIdentityAndSharedSource() => InDirectory(directory =>
    {
        byte[] bytes = TestFontBuilder.CreateFontWithFamilyNames(Names);
        File.WriteAllBytes(Path.Combine(directory, "face.ttf"), bytes);
        OpenTypeFont.FontDiscoveryHeaders headers = OpenTypeFont.ReadDiscoveryHeaders(bytes, 0);
        TestAssert.Equal("字體", headers.FamilyName);
        TestAssert.Equal("字體|Review Family", string.Join("|", headers.FamilyNames));
        TestAssert.Equal("字體", OpenTypeFont.Load(bytes).FamilyName);
        var resolver = new WindowsFontResolver(directory);
        FontFaceResolution english = resolver.Resolve(new("Review Family"));
        FontFaceResolution localized = resolver.Resolve(new("字體"));
        TestAssert.True(!english.IsFallback && !localized.IsFallback, "Both declared aliases must resolve exactly.");
        TestAssert.Equal("Review Family", english.RequestedFamily);
        TestAssert.Equal("REVIEW FAMILY", resolver.Resolve(new("REVIEW FAMILY")).RequestedFamily);
        TestAssert.True(ReferenceEquals(english.Source, localized.Source), "Aliases must borrow the same owned program.");
        TestAssert.Equal(1, resolver.GetDiscoveredFonts().Count);
        TestAssert.True(FontProgramLoader.Load(english, default)!.MapCodePoint(0x1F600) != 0, "Alias lookup must retain supplementary glyph coverage.");
    });

    public static void FontFamilyAliasesKeepCollectionFaceOwnership() => InDirectory(directory =>
    {
        byte[] bytes = TestFontBuilder.CreateCollection("字體", "Other Family", Names);
        File.WriteAllBytes(Path.Combine(directory, "faces.ttc"), bytes);
        var resolver = new WindowsFontResolver(directory);
        FontFaceResolution alias = resolver.Resolve(new("Review Family"));
        FontFaceResolution other = resolver.Resolve(new("Other Family"));
        TestAssert.True(!alias.IsFallback && !other.IsFallback, "Collection names must resolve their own faces.");
        TestAssert.Equal(0, alias.FontFaceIndex);
        TestAssert.Equal(1, other.FontFaceIndex);
        TestAssert.True(ReferenceEquals(alias.Source, other.Source), "Collection faces share the file source, not the face index.");
        TestAssert.Equal(2, resolver.GetDiscoveredFonts().Count);
        TestAssert.Equal("字體", FontProgramLoader.Load(alias, default)!.FamilyName);
        TestAssert.Equal("Other Family", FontProgramLoader.Load(other, default)!.FamilyName);
    });

    public static void FontFamilyAliasesRetainStyleRanking() => InDirectory(directory =>
    {
        byte[] regular = TestFontBuilder.CreateFontWithFamilyNames(Names);
        byte[] bold = regular.ToArray();
        int os2 = TestFontBuilder.GetTableRange(bold, "OS/2").Offset;
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(bold.AsSpan(os2+4, 2), 700);
        File.WriteAllBytes(Path.Combine(directory, "regular.ttf"), regular);
        File.WriteAllBytes(Path.Combine(directory, "bold.ttf"), bold);
        var resolver = new WindowsFontResolver(directory);
        foreach (string family in new[] { "Review Family", "字體" }) {
            FontFaceResolution normal = resolver.Resolve(new(family));
            FontFaceResolution heavy = resolver.Resolve(new(family, Bold: true));
            TestAssert.Equal(400, normal.WeightClass);
            TestAssert.Equal(700, heavy.WeightClass);
            TestAssert.True(!normal.Bold && heavy.Bold, "Aliases keep ordinary style selection.");
        }
        TestAssert.Equal(2, resolver.GetDiscoveredFonts().Count);
    });

    public static void FontCompactDiscoveryMatchesFullFontAndCollectionCoverage() => InDirectory(directory =>
    {
        foreach (var (name, bytes) in new[] {
            ("single.ttf", TestFontBuilder.CreateFontWithFamilyNames(Names)),
            ("faces.ttc", TestFontBuilder.CreateCollection("First", "Other Family", Names)) })
        {
            string path = Path.Combine(directory, name);
            File.WriteAllBytes(path, bytes);
            TestAssert.True(WindowsFontResolver.TryReadCompactDiscoveryHeaders(path, out var compact), "A valid font must use compact discovery.");
            TestAssert.Equal(OpenTypeFont.GetCollectionFontCount(bytes), compact.Count);
            for (int face = 0; face < compact.Count; face++)
            {
                var original = OpenTypeFont.ReadDiscoveryHeaders(bytes, face);
                TestAssert.Equal(string.Join("|", original.FamilyNames), string.Join("|", compact[face].FamilyNames));
                TestAssert.Equal(original.WeightClass, compact[face].WeightClass);
                var font = OpenTypeFont.Load(bytes, face);
                foreach (int codePoint in new[] { 0, 32, 65, 0x100, 0x03b1, 0x1f600, 0x10ffff })
                    TestAssert.Equal(font.MapCodePoint(codePoint) != 0, compact[face].HasCodePoint!(codePoint));
            }
        }
    });

    public static void FontCompactDiscoveryDoesNotAllocateSparseOutlineGaps() => InDirectory(directory =>
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        var original = OpenTypeFont.ReadDiscoveryHeaders(bytes, 0);
        var name = TestFontBuilder.GetTableRange(bytes, "name");
        byte[] names = bytes.AsSpan(name.Offset, name.Length).ToArray();
        const int farOffset = 32 * 1024 * 1024;
        int count = OpenTypeFont.U16(bytes, 4);
        for (int i = 0; i < count; i++)
        {
            int record = 12 + i * 16;
            if (System.Text.Encoding.ASCII.GetString(bytes, record, 4) == "name")
                System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(record + 8, 4), farOffset);
        }
        string path = Path.Combine(directory, "sparse.ttf");
        using (var stream = File.Create(path))
        { stream.Write(bytes); stream.Position = farOffset; stream.Write(names); }
        long before = GC.GetAllocatedBytesForCurrentThread();
        TestAssert.True(WindowsFontResolver.TryReadCompactDiscoveryHeaders(path, out var compact), "Sparse tables must stay readable.");
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        TestAssert.True(allocated < 1024 * 1024, "Discovery must not allocate the 32 MiB gap before the name table.");
        TestAssert.Equal(original.FamilyName, compact[0].FamilyName);
        // An invalid table extent must not be hidden by repacking.
        using (var stream = File.OpenWrite(path)) { stream.SetLength(bytes.Length); }
        TestAssert.True(!WindowsFontResolver.TryReadCompactDiscoveryHeaders(path, out _), "Truncated extents must use the validated fallback path.");
    });

    private static void InDirectory(Action<string> action)
    {
        string directory = Path.Combine(Path.GetTempPath(), "font-alias-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try { action(directory); }
        finally {
            WindowsFontResolver.InvalidateDiscoveryCaches();
            Directory.Delete(directory, recursive: true);
        }
    }
}
