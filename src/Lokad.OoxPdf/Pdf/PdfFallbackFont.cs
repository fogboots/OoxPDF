using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

// RV01: deterministic diagnosed fallback face for document text with no usable
// embeddable font. The standard-14 Helvetica family needs no embedding, so text
// stays visible and extractable; layout measures with the same deterministic
// average advances the fallback measurers use, and emission positions every rune
// absolutely, so measured and emitted selection agree by construction.
// WinAnsi covers U+0020-U+00FF plus the cp1252 0x80-0x9F set; anything else
// renders as question mark (diagnosed) rather than vanishing silently.
internal sealed class PdfFallbackFont
{
    public static PdfFallbackFont Helvetica { get; } = new("Helvetica", bold: false, italic: false);
    public static PdfFallbackFont HelveticaBold { get; } = new("Helvetica-Bold", bold: true, italic: false);
    public static PdfFallbackFont HelveticaOblique { get; } = new("Helvetica-Oblique", bold: false, italic: true);
    public static PdfFallbackFont HelveticaBoldOblique { get; } = new("Helvetica-BoldOblique", bold: true, italic: true);

    private PdfFallbackFont(string faceName, bool bold, bool italic)
    {
        FaceName = faceName;
        Bold = bold;
        Italic = italic;
    }

    public string FaceName { get; }

    public bool Bold { get; }

    public bool Italic { get; }

    public string ResourceKey => "Fallback-" + FaceName;

    // Fixed resource names in face order so output stays deterministic across documents.
    public static string ResourceNameFor(bool bold, bool italic)
    {
        return bold ? (italic ? "FF4" : "FF2") : (italic ? "FF3" : "FF1");
    }

    // Page assembly for conversion-wide fallback use (slides and chart parts share
    // one resolver): fixed face order keeps output deterministic.
    public static IReadOnlyList<PdfFallbackFontResource> ToResources(IEnumerable<PdfFallbackFont> faces)
    {
        return faces
            .Distinct()
            .OrderBy(face => ResourceNameFor(face.Bold, face.Italic), StringComparer.Ordinal)
            .Select(face => new PdfFallbackFontResource(ResourceNameFor(face.Bold, face.Italic), face))
            .ToArray();
    }

    public static PdfFallbackFont ForStyle(bool bold, bool italic)
    {
        return bold ? (italic ? HelveticaBoldOblique : HelveticaBold) : (italic ? HelveticaOblique : Helvetica);
    }

    // Deterministic fallback metrics (1000-unit em): 500 for most runes, 1000
    // for wide East-Asian/fullwidth ranges, 0 for combining marks. These are
    // diagnosed-degradation constants, not a claim about Helvetica advances.
    public const int UnitsPerEm = 1000;
    public const double AscentEm = 0.718d;
    public const double DescentEm = 0.207d;
    public const double SingleLineHeightEm = 1.2d;
    public const double UnderlinePositionEm = -0.1d;
    public const double UnderlineThicknessEm = 0.05d;
    public const double StrikeoutPositionEm = 0.25d;
    public const double StrikeoutThicknessEm = 0.05d;

    public static double MeasureAdvanceEm(Rune rune)
    {
        return MeasureAdvanceEm(rune.Value);
    }

    public static double MeasureAdvanceEm(int codePoint)
    {
        if (codePoint < 0x20 || codePoint == 0x7F)
        {
            return 0d;
        }

        if (IsNonRenderedControl(codePoint))
        {
            return 0d;
        }

        // Combining marks render onto their base glyph without advancing.
        UnicodeCategory markCategory = CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(codePoint), 0);
        if (markCategory is UnicodeCategory.NonSpacingMark or UnicodeCategory.EnclosingMark)
        {
            return 0d;
        }

        if (IsWide(codePoint))
        {
            return 1000d;
        }

        return 500d;
    }

    // RV01: layout constructs and format controls never reach glyph emission
    // (tabs expand to spacing, breaks split lines, marks combine). Excluding them
    // keeps missing-glyph detection, subset question-mark inclusion and emission
    // substitution aligned with what emission can actually drop.
    public static bool IsNonRenderedControl(Rune rune)
    {
        return IsNonRenderedControl(rune.Value);
    }

    public static bool IsNonRenderedControl(int codePoint)
    {
        if (codePoint < 0 || codePoint > 0x10FFFF)
        {
            return true;
        }

        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(codePoint), 0);
        return category is UnicodeCategory.Control
            or UnicodeCategory.Format
            or UnicodeCategory.LineSeparator
            or UnicodeCategory.ParagraphSeparator
            or UnicodeCategory.Surrogate
            || codePoint is >= 0xFE00 and <= 0xFE0F;
    }
    private static bool IsWide(int codePoint)
    {
        return (codePoint >= 0x1100 && codePoint <= 0x115F)
            || (codePoint >= 0x2E80 && codePoint <= 0xA4CF)
            || (codePoint >= 0xAC00 && codePoint <= 0xD7A3)
            || (codePoint >= 0xF900 && codePoint <= 0xFAFF)
            || (codePoint >= 0xFE10 && codePoint <= 0xFE1F)
            || (codePoint >= 0xFE30 && codePoint <= 0xFE4F)
            || (codePoint >= 0xFF00 && codePoint <= 0xFF60)
            || (codePoint >= 0xFFE0 && codePoint <= 0xFFE6)
            || (codePoint >= 0x20000 && codePoint <= 0x3FFFF);
    }

    // WinAnsi (cp1252): ASCII/Latin-1 are identity; the 0x80-0x9F set maps through
    // the table below; controls, DEL, the five undefined positions and everything
    // outside WinAnsi render as question mark rather than vanishing.
    public const byte SubstitutionByte = 0x3F;

    public static bool TryEncodeWinAnsi(int codePoint, out byte value)
    {
        if (codePoint >= 0x20 && codePoint <= 0x7E)
        {
            value = (byte)codePoint;
            return true;
        }

        if (codePoint >= 0xA0 && codePoint <= 0xFF)
        {
            value = (byte)codePoint;
            return true;
        }

        for (int i = 0; i < SpecialWinAnsiBytes.Length; i++)
        {
            if (SpecialWinAnsiCodePoints[i] == codePoint)
            {
                value = SpecialWinAnsiBytes[i];
                return true;
            }
        }

        value = SubstitutionByte;
        return false;
    }

    private static readonly byte[] SpecialWinAnsiBytes = new byte[]
    {
        0x80, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8A, 0x8B, 0x8C, 0x8E,
        0x91, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0x9B, 0x9C, 0x9E, 0x9F,
    };

    private static readonly int[] SpecialWinAnsiCodePoints = new int[]
    {
        0x20AC, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x017D,
        0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x017E, 0x0178,
    };

    // Static WinAnsi ToUnicode map (every encodable byte): extraction stays exact
    // for preserved text without per-conversion collection. Batched at 100 entries.
    public static string WinAnsiToUnicodeCMap { get; } = BuildWinAnsiToUnicodeCMap();

    private static string BuildWinAnsiToUnicodeCMap()
    {
        var entries = new List<(byte Code, int Unicode)>();
        for (int code = 0x20; code <= 0xFF; code++)
        {
            if (TryEncodeWinAnsi(code, out byte value) && value == code)
            {
                entries.Add(((byte)code, code));
            }
        }

        for (int i = 0; i < SpecialWinAnsiBytes.Length; i++)
        {
            entries.Add((SpecialWinAnsiBytes[i], SpecialWinAnsiCodePoints[i]));
        }

        entries.Sort((left, right) => left.Code.CompareTo(right.Code));
        var builder = new StringBuilder();
        builder.AppendLine("/CIDInit /ProcSet findresource begin");
        builder.AppendLine("12 dict begin");
        builder.AppendLine("begincmap");
        builder.AppendLine("/CIDSystemInfo << /Registry (Adobe) /Ordering (WinAnsi) /Supplement 0 >> def");
        builder.AppendLine("/CMapName /Adobe-WinAnsi-UCS def");
        builder.AppendLine("/CMapType 2 def");
        builder.AppendLine("1 begincodespacerange");
        builder.AppendLine("<20> <FF>");
        builder.AppendLine("endcodespacerange");
        for (int start = 0; start < entries.Count; start += 100)
        {
            int count = Math.Min(100, entries.Count - start);
            builder.AppendLine(count.ToString(CultureInfo.InvariantCulture) + " beginbfchar");
            for (int i = start; i < start + count; i++)
            {
                builder.Append((char)60).Append(entries[i].Code.ToString("X2", CultureInfo.InvariantCulture)).Append((char)62).Append((char)32).Append((char)60).Append(entries[i].Unicode.ToString("X4", CultureInfo.InvariantCulture)).Append((char)62);
                builder.AppendLine();
            }
            builder.AppendLine("endbfchar");
        }
        builder.AppendLine("endcmap");
        builder.AppendLine("CMapName currentdict /CMap defineresource pop");
        builder.AppendLine("end");
        builder.AppendLine("end");
        return builder.ToString();
    }
}

internal sealed record PdfFallbackFontResource(string ResourceName, PdfFallbackFont Font);

internal readonly record struct PdfFallbackGlyph(double X, double Y, byte Code);
