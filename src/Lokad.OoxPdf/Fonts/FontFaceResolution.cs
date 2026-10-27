namespace Lokad.OoxPdf.Fonts;

public sealed record FontFaceResolution(
    string RequestedFamily,
    string ResolvedFamily,
    FontStyleKey Style,
    IFontProgramSource Source,
    bool IsFallback)
{
    // Optional compact coverage hint from trusted discovery. False means this
    // font cannot supply a glyph, so fallback need not load its whole program.
    internal Func<int, bool>? HasCodePoint { get; init; }
    internal string FamilyName => ResolvedFamily;
    internal bool Bold => Style.Bold;
    internal bool Italic => Style.Italic;
    internal int WeightClass => Style.WeightClass;
    internal int FontFaceIndex => Style.FaceIndex;
    internal bool HasMathTable => Style.HasMathTable;
}

public sealed record FontStyleKey(
    bool Bold = false,
    bool Italic = false,
    int WeightClass = 400,
    int FaceIndex = 0,
    bool HasMathTable = false);
