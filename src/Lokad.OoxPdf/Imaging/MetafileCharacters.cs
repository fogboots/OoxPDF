namespace Lokad.OoxPdf.Imaging;

// Turns metafile text into Unicode. GDI text is encoded in the selected font's
// charset; the Symbol and MT Extra fonts (equation editors draw with them) use
// their own encodings, which map here to the Unicode characters they show so that
// installed fonts can draw them. A character without a known equivalent throws
// NotSupportedException rather than drawing the wrong glyph.
internal static class MetafileCharacters
{
    // Adobe Symbol encoding from 0x20 (Unicode SYMBOL.TXT, with the bracket and
    // integral pieces mapped to their Miscellaneous Technical counterparts).
    private static readonly int[] Symbol =
    [
        0x0020, 0x0021, 0x2200, 0x0023, 0x2203, 0x0025, 0x0026, 0x220B, 0x0028, 0x0029, 0x2217, 0x002B, 0x002C, 0x2212, 0x002E, 0x002F,
        0x0030, 0x0031, 0x0032, 0x0033, 0x0034, 0x0035, 0x0036, 0x0037, 0x0038, 0x0039, 0x003A, 0x003B, 0x003C, 0x003D, 0x003E, 0x003F,
        0x2245, 0x0391, 0x0392, 0x03A7, 0x0394, 0x0395, 0x03A6, 0x0393, 0x0397, 0x0399, 0x03D1, 0x039A, 0x039B, 0x039C, 0x039D, 0x039F,
        0x03A0, 0x0398, 0x03A1, 0x03A3, 0x03A4, 0x03A5, 0x03C2, 0x03A9, 0x039E, 0x03A8, 0x0396, 0x005B, 0x2234, 0x005D, 0x22A5, 0x005F,
        0x203E, 0x03B1, 0x03B2, 0x03C7, 0x03B4, 0x03B5, 0x03C6, 0x03B3, 0x03B7, 0x03B9, 0x03D5, 0x03BA, 0x03BB, 0x03BC, 0x03BD, 0x03BF,
        0x03C0, 0x03B8, 0x03C1, 0x03C3, 0x03C4, 0x03C5, 0x03D6, 0x03C9, 0x03BE, 0x03C8, 0x03B6, 0x007B, 0x007C, 0x007D, 0x223C, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0x20AC, 0x03D2, 0x2032, 0x2264, 0x2044, 0x221E, 0x0192, 0x2663, 0x2666, 0x2665, 0x2660, 0x2194, 0x2190, 0x2191, 0x2192, 0x2193,
        0x00B0, 0x00B1, 0x2033, 0x2265, 0x00D7, 0x221D, 0x2202, 0x2022, 0x00F7, 0x2260, 0x2261, 0x2248, 0x2026, 0x23D0, 0x23AF, 0x21B5,
        0x2135, 0x2111, 0x211C, 0x2118, 0x2297, 0x2295, 0x2205, 0x2229, 0x222A, 0x2283, 0x2287, 0x2284, 0x2282, 0x2286, 0x2208, 0x2209,
        0x2220, 0x2207, 0x00AE, 0x00A9, 0x2122, 0x220F, 0x221A, 0x22C5, 0x00AC, 0x2227, 0x2228, 0x21D4, 0x21D0, 0x21D1, 0x21D2, 0x21D3,
        0x25CA, 0x27E8, 0x00AE, 0x00A9, 0x2122, 0x2211, 0x239B, 0x239C, 0x239D, 0x23A1, 0x23A2, 0x23A3, 0x23A7, 0x23A8, 0x23A9, 0x23AA,
        0, 0x27E9, 0x222B, 0x2320, 0x23AE, 0x2321, 0x239E, 0x239F, 0x23A0, 0x23A4, 0x23A5, 0x23A6, 0x23AB, 0x23AC, 0x23AD, 0
    ];

    // MT Extra (MathType) characters that have standard Unicode equivalents; the
    // table follows LibreOffice's MT Extra conversion, without its private-use entries.
    private static readonly Dictionary<int, int> MtExtra = new()
    {
        [0x20] = 0x0020, [0x23] = 0x0300, [0x24] = 0x0302, [0x25] = 0x0303, [0x26] = 0x0307, [0x28] = 0x2323, [0x29] = 0x2322,
        [0x3A] = 0x223C, [0x3B] = 0x2243, [0x3C] = 0x22B2, [0x3D] = 0x226A, [0x3E] = 0x22B3, [0x3F] = 0x226B,
        [0x40] = 0x225C, [0x41] = 0x2259, [0x42] = 0x2250, [0x43] = 0x2210, [0x44] = 0x019B, [0x49] = 0x2229, [0x4A] = 0x2127,
        [0x4B] = 0x2026, [0x4C] = 0x22EF, [0x4D] = 0x22EE, [0x4E] = 0x22F0, [0x4F] = 0x22F1,
        [0x50] = 0x2225, [0x51] = 0x2235, [0x52] = 0x2221, [0x53] = 0x2222, [0x55] = 0x222A, [0x56] = 0x25B3, [0x57] = 0x25A1,
        [0x58] = 0x25AD, [0x59] = 0x25B1, [0x5A] = 0x2197, [0x5B] = 0x2199, [0x5D] = 0x2198, [0x5E] = 0x2196,
        [0x60] = 0x2035, [0x61] = 0x21A6, [0x62] = 0x2195, [0x63] = 0x21D5, [0x64] = 0x25CB, [0x65] = 0x2299, [0x66] = 0x227B,
        [0x68] = 0x210F, [0x6C] = 0x2113, [0x6D] = 0x2213, [0x6F] = 0x2218,
        [0x70] = 0x227A, [0x72] = 0x20D7, [0x73] = 0x20D6, [0x74] = 0x20E1, [0x76] = 0x20D1, [0x77] = 0x20D0, [0x7B] = 0xFE38, [0x7D] = 0xFE37,
        [0x80] = 0x21C4, [0x83] = 0x21CC, [0x86] = 0x21C0, [0x87] = 0x21BD,
        [0xA1] = 0x211D, [0xA2] = 0x2124, [0xA3] = 0x2102, [0xA4] = 0x211A, [0xA5] = 0x2115, [0xA7] = 0x301A, [0xA8] = 0x301B,
        [0xE8] = 0x0160, [0xE9] = 0x00DD, [0xEA] = 0x00DE, [0xEB] = 0x00D0, [0xEC] = 0x0161, [0xED] = 0x00FD, [0xEE] = 0x00FE, [0xEF] = 0x00F0,
        [0xF0] = 0xFB01, [0xF1] = 0xFB02, [0xF2] = 0x0131, [0xF3] = 0x00B9, [0xF4] = 0x00B2, [0xF5] = 0x00B3, [0xF6] = 0x00BD, [0xF7] = 0x00BC,
        [0xF8] = 0x00BE, [0xF9] = 0x2044, [0xFA] = 0x00A6, [0xFB] = 0x02DD, [0xFC] = 0x02D8, [0xFD] = 0x02C7, [0xFE] = 0x02DA, [0xFF] = 0x02DB
    };

    // Windows-1252 0x80-0x9F; the rest of the code page is ISO-8859-1.
    private static readonly int[] Windows1252 =
    [
        0x20AC, 0, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0, 0x017D, 0,
        0, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0, 0x017E, 0x0178
    ];

    // Windows-1253 (Greek) 0x80-0xBF; 0xC0-0xFE are U+0390 onwards except 0xD2.
    private static readonly int[] Windows1253 =
    [
        0x20AC, 0, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0, 0x2030, 0, 0x2039, 0, 0, 0, 0,
        0, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0, 0x2122, 0, 0x203A, 0, 0, 0, 0,
        0x00A0, 0x0385, 0x0386, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7, 0x00A8, 0x00A9, 0, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x2015,
        0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x0384, 0x00B5, 0x00B6, 0x00B7, 0x0388, 0x0389, 0x038A, 0x00BB, 0x038C, 0x00BD, 0x038E, 0x038F
    ];

    private enum Encoding { Symbol, MtExtra, Windows1252, Windows1253, CodePage }

    // A face named with a leading '@' is the vertical variant of the named font: GDI
    // turns its East Asian characters a quarter turn so that they stand upright when
    // the text runs top to bottom, and draws every other character as the plain face.
    public static bool IsVertical(string faceName) => faceName.StartsWith('@');

    // Characters that stand upright in vertical text (Unicode UAX #50 orientation U and
    // Tu: Han, kana, Hangul, Yi, compatibility and full-width forms, CJK punctuation).
    // Brackets, dashes and similar (orientation Tr) are upright only through the font's
    // vertical alternate glyph and otherwise stay turned with the line.
    public static bool IsUprightInVertical(int codePoint, out bool needsAlternate)
    {
        needsAlternate = codePoint is >= 0x3008 and <= 0x3011 or >= 0x3014 and <= 0x301F or 0x3030 or 0x30A0 or 0x30FC or
            0xFE59 or 0xFE5A or 0xFE5B or 0xFE5C or 0xFE5D or 0xFE5E or
            0xFF08 or 0xFF09 or 0xFF0D or 0xFF1C or 0xFF1D or 0xFF1E or 0xFF3B or 0xFF3D or 0xFF3F or
            0xFF5B or 0xFF5C or 0xFF5D or 0xFF5E or 0xFF5F or 0xFF60 or 0xFFE3;
        return codePoint is >= 0x1100 and <= 0x11FF or >= 0x2E80 and <= 0x2FFF or >= 0x3000 and <= 0x4DBF or
            >= 0x4E00 and <= 0xA4CF or >= 0xA960 and <= 0xA97F or >= 0xAC00 and <= 0xD7FF or
            >= 0xF900 and <= 0xFAFF or >= 0xFE10 and <= 0xFE1F or >= 0xFE30 and <= 0xFE6F or
            >= 0xFF00 and <= 0xFF60 or >= 0xFFE0 and <= 0xFFE7 or >= 0x20000 and <= 0x3FFFD;
    }

    // The family that draws the mapped characters of a metafile font.
    public static string DrawingFamily(string faceName)
    {
        faceName = faceName.TrimStart('@');
        return Classify(faceName, charset: 0) switch
        {
            Encoding.Symbol or Encoding.MtExtra => "Times New Roman",
            _ => faceName.Length == 0 || faceName.Equals("System", StringComparison.OrdinalIgnoreCase) ? "Arial" : faceName
        };
    }

    // Eight-bit text (WMF records and EMR_EXTTEXTOUTA).
    public static int[] Decode(string faceName, int charset, ReadOnlySpan<byte> text)
    {
        Encoding encoding = Classify(faceName, charset);
        if (encoding == Encoding.CodePage)
        {
            return DecodeCodePage(faceName, charset, text);
        }

        var codePoints = new int[text.Length];
        for (int i = 0; i < text.Length; i++)
        {
            codePoints[i] = Map(encoding, faceName, text[i]);
        }

        return codePoints;
    }

    // UTF-16 text (EMR_EXTTEXTOUTW). Symbol fonts address their glyphs either directly
    // or through the U+F020-U+F0FF private-use range. A surrogate pair yields one code
    // point and -1 for its low half, which keeps the character advances aligned. The
    // character set does not apply to Unicode text.
    public static int[] Decode(string faceName, int charset, ReadOnlySpan<char> text)
    {
        Encoding encoding = Classify(faceName, charset: 0);
        var codePoints = new int[text.Length];
        for (int i = 0; i < text.Length; i++)
        {
            int code = text[i];
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                codePoints[i] = char.ConvertToUtf32(text[i], text[i + 1]);
                codePoints[++i] = -1;
                continue;
            }

            if (encoding is Encoding.Symbol or Encoding.MtExtra)
            {
                codePoints[i] = Map(encoding, faceName, code is >= 0xF000 and <= 0xF0FF ? code - 0xF000 : code is <= 0xFF ? code : 0);
            }
            else if (code is >= 0xF000 and <= 0xF0FF)
            {
                throw new NotSupportedException($"Metafile font '{faceName}' uses symbol-font characters.");
            }
            else
            {
                codePoints[i] = code;
            }
        }

        return codePoints;
    }

    private static Encoding Classify(string faceName, int charset)
    {
        faceName = faceName.TrimStart('@');
        if (faceName.Equals("Symbol", StringComparison.OrdinalIgnoreCase))
        {
            return Encoding.Symbol;
        }

        if (faceName.Equals("MT Extra", StringComparison.OrdinalIgnoreCase))
        {
            return Encoding.MtExtra;
        }

        return charset switch
        {
            0 or 1 => Encoding.Windows1252,
            161 => Encoding.Windows1253,
            _ when CodePageOf(charset) != 0 => Encoding.CodePage,
            _ => throw new NotSupportedException($"Metafile font '{faceName}' uses character set {charset}.")
        };
    }

    // The Windows code page of a GDI character set, or 0.
    private static int CodePageOf(int charset) => charset switch
    {
        77 => 10000, 128 => 932, 129 => 949, 130 => 1361, 134 => 936, 136 => 950, 162 => 1254, 163 => 1258,
        177 => 1255, 178 => 1256, 186 => 1257, 204 => 1251, 222 => 874, 238 => 1250, 255 => 437, _ => 0
    };

    // Eight-bit text in a Windows code page. A double-byte character yields its code
    // point at the lead byte and -1 at the trail byte, which keeps the advances aligned.
    private static int[] DecodeCodePage(string faceName, int charset, ReadOnlySpan<byte> text)
    {
        int codePage = CodePageOf(charset);
        System.Text.Encoding encoding = System.Text.CodePagesEncodingProvider.Instance.GetEncoding(codePage,
            System.Text.EncoderFallback.ExceptionFallback, System.Text.DecoderFallback.ExceptionFallback) ??
            throw new NotSupportedException($"Metafile font '{faceName}' uses unavailable code page {codePage}.");
        var codePoints = new int[text.Length];
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] < 0x20)
            {
                codePoints[i] = text[i];
                continue;
            }

            int length = i + 1 < text.Length && IsLeadByte(codePage, text[i]) ? 2 : 1;
            string decoded;
            try
            {
                decoded = encoding.GetString(text.Slice(i, length));
            }
            catch (System.Text.DecoderFallbackException)
            {
                throw new NotSupportedException($"Character 0x{text[i]:X2} of metafile font '{faceName}' has no Unicode equivalent.");
            }

            if (decoded.Length == 0 || decoded.Length > 2)
            {
                throw new NotSupportedException($"Character 0x{text[i]:X2} of metafile font '{faceName}' has no Unicode equivalent.");
            }

            codePoints[i] = char.ConvertToUtf32(decoded, 0);
            if (length == 2)
            {
                codePoints[++i] = -1;
            }
        }

        return codePoints;
    }

    private static bool IsLeadByte(int codePage, byte value) => codePage switch
    {
        932 => value is >= 0x81 and <= 0x9F or >= 0xE0 and <= 0xFC,
        936 or 949 or 950 => value is >= 0x81 and <= 0xFE,
        1361 => value is >= 0x84 and <= 0xD3 or >= 0xD8 and <= 0xDE or >= 0xE0 and <= 0xF9,
        _ => false
    };

    private static int Map(Encoding encoding, string faceName, int code)
    {
        // Control characters keep their advance and draw nothing.
        if (code < 0x20)
        {
            return code;
        }

        int mapped = encoding switch
        {
            Encoding.Symbol => Symbol[code - 0x20],
            Encoding.MtExtra => MtExtra.GetValueOrDefault(code),
            Encoding.Windows1252 => code is >= 0x80 and <= 0x9F ? Windows1252[code - 0x80] : code,
            _ => code < 0x80 ? code : code < 0xC0 ? Windows1253[code - 0x80] : code is 0xD2 or 0xFF ? 0 : 0x0390 + code - 0xC0
        };

        if (mapped == 0)
        {
            throw new NotSupportedException($"Character 0x{code:X2} of metafile font '{faceName}' has no Unicode equivalent.");
        }

        return mapped;
    }
}
