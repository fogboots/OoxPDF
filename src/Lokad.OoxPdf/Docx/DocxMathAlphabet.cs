namespace Lokad.OoxPdf.Docx;

// OMML run styles (m:sty, m:scr) to Unicode Mathematical Alphanumeric Symbols.
// Letters default to italic; digits and other characters stay upright. Holes in
// the Unicode block map to their Letterlike Symbols code points.
internal static class MathAlphabet
{
    public static int Map(int codePoint, string? style, string? script)
    {
        bool bold = style is "b" or "bi";
        bool italic = style is null or "i" or "bi";
        bool latinUpper = codePoint is >= 'A' and <= 'Z', latinLower = codePoint is >= 'a' and <= 'z';
        bool digit = codePoint is >= '0' and <= '9';
        bool greekUpper = codePoint is >= 0x0391 and <= 0x03A9 && codePoint != 0x03A2;
        bool greekLower = codePoint is >= 0x03B1 and <= 0x03C9;
        int offset = latinUpper ? codePoint - 'A' : latinLower ? codePoint - 'a' + 26 : 0;
        switch (script)
        {
            case "script" when latinUpper || latinLower:
                return Hole(bold ? 0x1D4D0 + offset : ScriptHoles(codePoint) ?? 0x1D49C + offset);
            case "fraktur" when latinUpper || latinLower:
                return bold ? 0x1D56C + offset : FrakturHoles(codePoint) ?? 0x1D504 + offset;
            case "double-struck" when latinUpper || latinLower:
                return DoubleStruckHoles(codePoint) ?? 0x1D538 + offset;
            case "double-struck" when digit:
                return 0x1D7D8 + codePoint - '0';
            case "sans-serif" when latinUpper || latinLower:
                return (bold, italic && style is not null) switch
                {
                    (true, true) => 0x1D63C + offset,
                    (true, false) => 0x1D5D4 + offset,
                    (false, true) => 0x1D608 + offset,
                    _ => 0x1D5A0 + offset,
                };
            case "sans-serif" when digit:
                return (bold ? 0x1D7EC : 0x1D7E2) + codePoint - '0';
            case "monospace" when latinUpper || latinLower:
                return 0x1D670 + offset;
            case "monospace" when digit:
                return 0x1D7F6 + codePoint - '0';
        }
        if (latinUpper || latinLower)
        {
            return (bold, italic) switch
            {
                (true, true) => 0x1D468 + offset,
                (true, false) => 0x1D400 + offset,
                (false, true) => codePoint == 'h' ? 0x210E : 0x1D434 + offset,
                _ => codePoint,
            };
        }
        if (digit) { return bold ? 0x1D7CE + codePoint - '0' : codePoint; }
        if (greekUpper || greekLower)
        {
            int greek = greekUpper ? codePoint - 0x0391 : codePoint - 0x03B1 + 26;
            return (bold, italic) switch
            {
                (true, true) => 0x1D71C + greek,
                (true, false) => 0x1D6A8 + greek,
                (false, true) => 0x1D6E2 + greek,
                _ => codePoint,
            };
        }
        if (codePoint == 0x2202 && italic && !bold) { return 0x1D715; }
        return codePoint;
    }

    private static int Hole(int codePoint) => codePoint;

    private static int? ScriptHoles(int c) => c switch
    {
        'B' => 0x212C, 'E' => 0x2130, 'F' => 0x2131, 'H' => 0x210B, 'I' => 0x2110, 'L' => 0x2112,
        'M' => 0x2133, 'R' => 0x211B, 'e' => 0x212F, 'g' => 0x210A, 'o' => 0x2134, _ => null,
    };

    private static int? FrakturHoles(int c) => c switch
    {
        'C' => 0x212D, 'H' => 0x210C, 'I' => 0x2111, 'R' => 0x211C, 'Z' => 0x2128, _ => null,
    };

    private static int? DoubleStruckHoles(int c) => c switch
    {
        'C' => 0x2102, 'H' => 0x210D, 'N' => 0x2115, 'P' => 0x2119, 'Q' => 0x211A, 'R' => 0x211D, 'Z' => 0x2124, _ => null,
    };
}
