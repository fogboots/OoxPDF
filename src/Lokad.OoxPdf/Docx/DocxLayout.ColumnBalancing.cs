namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Word balances the columns of a section that a continuous section break ends:
    // the region's content is redistributed so the columns end at about the same
    // height. Placed items keep their order and their vertical offsets within a
    // column; a column boundary falls only between items. Returns the new cursor,
    // or null when the region holds an item kind that cannot move between columns.
    private static double? BalanceColumnRegion(List<DocxLayoutItem> items, int start,
        IReadOnlyList<DocxLayoutColumnFrame> frames, double regionTop)
    {
        int count = frames.Count;
        // Lines keep the wrap width of the column that measured them: balance equal columns only.
        if (count <= 1 || start >= items.Count || frames.Any(frame => Math.Abs(frame.Width - frames[0].Width) > 0.01d)) { return null; }
        var placed = new List<(int Index, int Column, double Top, double Bottom)>();
        for (int index = start; index < items.Count; index++)
        {
            if (Bounds(items[index]) is not { } bounds) { return null; }
            int? column = DocxLayoutColumnOwnership.ResolveColumnIndex(frames, bounds.X, 0d);
            if (column is null) { return null; }
            placed.Add((index, column.Value, bounds.Top, bounds.Bottom));
        }
        // Flow position: columns stacked end to end, keeping each column's own gaps.
        var used = new double[count];
        foreach (var item in placed) { used[item.Column] = Math.Max(used[item.Column], regionTop - item.Bottom); }
        var offsets = new double[count];
        for (int c = 1; c < count; c++) { offsets[c] = offsets[c - 1] + used[c - 1]; }
        double target = used.Sum() / count;
        // Every column's first item takes the place of the region's first item.
        double alignTop = placed.Max(item => item.Top);
        var ordered = placed.OrderBy(item => item.Column).ThenByDescending(item => item.Top).ToList();
        int current = 0;
        double columnStart = 0d, bottom = regionTop;
        bool columnHasItems = false;
        foreach (var item in ordered)
        {
            double flowTop = offsets[item.Column] + regionTop - item.Top;
            double flowBottom = offsets[item.Column] + regionTop - item.Bottom;
            if (columnHasItems && current < count - 1 && flowBottom - columnStart > target + 0.5d)
            {
                current++;
                columnStart = flowTop;
                columnHasItems = false;
            }
            double deltaY = alignTop - (flowTop - columnStart) - item.Top;
            double deltaX = frames[current].X - frames[item.Column].X;
            items[item.Index] = Translate(items[item.Index], deltaX, deltaY)!;
            bottom = Math.Min(bottom, item.Bottom + deltaY);
            columnHasItems = true;
        }
        return bottom;
    }

    private static (double X, double Top, double Bottom)? Bounds(DocxLayoutItem item) => item switch
    {
        DocxTextLineLayout line => (line.X, line.BaselineY + line.FontSize, line.BaselineY + line.FontSize - (line.LineHeight ?? line.FontSize * 1.2d)),
        DocxInlineImageLayout image => (image.X, image.Y + image.Height, image.Y),
        DocxTableRowLayout row => (row.Table.TableX, row.Y + row.Height, row.Y),
        DocxEquationLayout equation => (equation.X, equation.BaselineY + equation.Math.BaselineInset * equation.Scale,
            equation.BaselineY + (equation.Math.BaselineInset - equation.Math.LineHeight) * equation.Scale),
        _ => null,
    };

    private static DocxLayoutItem? Translate(DocxLayoutItem item, double deltaX, double deltaY) => item switch
    {
        DocxTextLineLayout line => DocxRenderer.TranslateTextLine(line, deltaX, deltaY),
        DocxInlineImageLayout image => image.Translate(deltaX, deltaY),
        DocxTableRowLayout row => DocxRenderer.TranslateTableRow(row, deltaX, deltaY),
        DocxEquationLayout equation => equation with { X = equation.X + deltaX, BaselineY = equation.BaselineY + deltaY },
        _ => null,
    };
}
