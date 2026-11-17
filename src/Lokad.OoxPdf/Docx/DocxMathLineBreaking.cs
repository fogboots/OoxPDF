namespace Lokad.OoxPdf.Docx;

// Word breaks a display equation that is wider than its column after a
// top-level binary or relation operator; continuation lines align right.
internal static class DocxMathLineBreaking
{
    public static IReadOnlyList<DocxPreparedMathParagraph> Split(DocxPreparedMathParagraph math, double width, CancellationToken token)
    {
        DocxMathBox box = math.Box;
        var glyphs = box.Glyphs;
        // Break positions in x: the end of each glyph after which a break is allowed.
        var stops = box.InlineBreakAfterGlyphs
            .Where(index => index > 0 && index <= glyphs.Count)
            .Select(index => glyphs[index - 1].X + glyphs[index - 1].Advance)
            .Distinct()
            .Order()
            .ToList();
        stops.Add(box.Width);
        var lines = new List<DocxPreparedMathParagraph>();
        double start = 0d, previous = 0d;
        foreach (double stop in stops)
        {
            token.ThrowIfCancellationRequested();
            if (stop - start > width && previous > start)
            {
                lines.Add(Slice(math, start, previous));
                start = previous;
            }
            previous = stop;
        }
        lines.Add(Slice(math, start, box.Width));
        return lines;
    }

    private static DocxPreparedMathParagraph Slice(DocxPreparedMathParagraph math, double start, double end)
    {
        DocxMathBox box = math.Box;
        const double epsilon = 1e-6;
        bool Inside(double x) => x >= start - epsilon && x < end - epsilon;
        var glyphs = box.Glyphs.Where(g => Inside(g.X)).Select(g => g with { X = g.X - start }).ToArray();
        var rules = box.Rules.Where(r => Inside(r.X)).Select(r => r with { X = r.X - start }).ToArray();
        var slice = box with
        {
            Width = end - start, LogicalAdvance = null, Glyphs = glyphs, Rules = rules,
            InlineBreakAfterGlyphs = [],
        };
        return math with { Box = slice };
    }
}
