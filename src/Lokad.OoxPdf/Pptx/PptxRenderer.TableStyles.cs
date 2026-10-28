namespace Lokad.OoxPdf.Pptx;

internal static class PptxTableStyleResolver
{
    public static PptxSceneFillStyle ReadCellFill(PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme, PptxColorMap colorMap)
    {
        if (tableStyle.Definition is not null)
        {
            return PptxTableStyleCascade.ReadCellFill(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, colorMap);
        }

        double alpha = 1d;
        if (!tableStyle.IsSupported ||
            !theme.TryResolveColor(tableStyle.Accent, colorMap, out RgbColor accent))
        {
            return default;
        }

        int bodyColumnIndex = columnIndex - (tableStyle.FirstColumn ? 1 : 0);
        if (tableStyle.Kind == PptxBuiltInTableStyleKind.MediumStyle2 &&
            ((tableStyle.FirstRow && rowIndex == 0) ||
                (tableStyle.LastRow && rowIndex == rowCount - 1) ||
                (tableStyle.FirstColumn && columnIndex == 0) ||
                (tableStyle.LastColumn && columnIndex == columnCount - 1)))
        {
            return new PptxSceneFillStyle(true, accent, alpha);
        }

        if (tableStyle.Kind == PptxBuiltInTableStyleKind.DarkStyle1)
        {
            if (tableStyle.FirstRow && rowIndex == 0)
            {
                return new PptxSceneFillStyle(true, new RgbColor(0, 0, 0), alpha);
            }

            if ((tableStyle.FirstColumn && columnIndex == 0) ||
                (tableStyle.LastColumn && columnIndex == columnCount - 1))
            {
                return new PptxSceneFillStyle(true, LinearLightColor.ShadeTowardBlack(accent, 0.60d), alpha);
            }

            if (tableStyle.LastRow && rowIndex == rowCount - 1)
            {
                // Office's built-in lastRow fill specifies accent shade 40000.
                return new PptxSceneFillStyle(true, LinearLightColor.ShadeTowardBlack(accent, 0.40d), alpha);
            }
        }

        // LightStyle1 firstRow fills only without bandRow: with bandRow the Office header
        // stays unfilled (dashboard probe); firstRow-only keeps the legacy accent fill.
        if (tableStyle.Kind == PptxBuiltInTableStyleKind.LightStyle1 &&
            tableStyle.FirstRow &&
            !tableStyle.BandRow &&
            rowIndex == 0)
        {
            return new PptxSceneFillStyle(true, accent, alpha);
        }

        int bodyRowIndex = rowIndex - (tableStyle.FirstRow ? 1 : 0);
        if (tableStyle.Kind == PptxBuiltInTableStyleKind.LightStyle1)
        {
            // LightStyle1 banded rows keep the raw accent (Office-calibrated: no alpha
            // wash and no firstRow fill).
            if (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0)
            {
                return new PptxSceneFillStyle(true, accent, alpha);
            }

            if (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0)
            {
                return new PptxSceneFillStyle(true, accent, 0.4d);
            }

            return default;
        }

        if (tableStyle.Kind == PptxBuiltInTableStyleKind.MediumStyle2)
        {
            bool banded = (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0) ||
                (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0);
            // MediumStyle2 band fills are linear-light tints toward white (Office-calibrated
            // exact weights 0.60 banded / 0.80 unbanded).
            RgbColor color = banded
                ? LinearLightColor.TintTowardWhite(accent, 0.60d)
                : LinearLightColor.TintTowardWhite(accent, 0.80d);
            return new PptxSceneFillStyle(true, color, alpha);
        }

        if (tableStyle.Kind == PptxBuiltInTableStyleKind.DarkStyle1)
        {
            bool banded = (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0) ||
                (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0);
            // DarkStyle1 band fills are linear-light shades toward black (Office-calibrated
            // exact weight 0.60 banded; unbanded rows keep the raw accent).
            RgbColor color = banded
                ? LinearLightColor.ShadeTowardBlack(accent, 0.60d)
                : accent;
            return new PptxSceneFillStyle(true, color, alpha);
        }

        return default;
    }

    public static PptxSceneTableCellTextStyle ReadCellTextStyle(PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme, PptxColorMap colorMap)
    {
        if (tableStyle.Definition is not null)
        {
            return PptxTableStyleCascade.ReadCellTextStyle(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, colorMap);
        }

        bool bold = false;
        RgbColor? color = null;
        bool supportedStyle = tableStyle.IsSupported &&
            tableStyle.Kind is PptxBuiltInTableStyleKind.MediumStyle2 or
                PptxBuiltInTableStyleKind.LightStyle1 or
                PptxBuiltInTableStyleKind.DarkStyle1;
        bool firstRow = tableStyle.FirstRow && rowIndex == 0;
        bool firstCol = tableStyle.FirstColumn && columnIndex == 0;
        bool lastRow = tableStyle.LastRow &&
            rowIndex == rowCount - 1;
        bool lastCol = tableStyle.LastColumn &&
            columnCount > 0 &&
            columnIndex == columnCount - 1;
        // With bandRow the LightStyle1 header stays unfilled, so firstRow text stays dark
        // like the body (Office-calibrated: black header text); filled headers (Medium/Dark,
        // and LightStyle1 firstRow-only) keep white text.
        if (supportedStyle &&
            firstRow &&
            !(tableStyle.Kind == PptxBuiltInTableStyleKind.LightStyle1 && tableStyle.BandRow) &&
            theme.TryResolveColor("lt1", colorMap, out RgbColor firstRowColor))
        {
            color = firstRowColor;
            bold = true;
        }

        // DarkStyle1 body text stays light on the dark fills (Office-calibrated: white
        // unstyled text on banded and unbanded body rows).
        if (color is null &&
            supportedStyle &&
            tableStyle.Kind == PptxBuiltInTableStyleKind.DarkStyle1 &&
            theme.TryResolveColor("lt1", colorMap, out RgbColor darkBodyColor))
        {
            color = darkBodyColor;
        }

        if (supportedStyle && (firstCol || lastRow || lastCol))
        {
            bold = true;
            if (tableStyle.Kind is PptxBuiltInTableStyleKind.MediumStyle2 or PptxBuiltInTableStyleKind.DarkStyle1 &&
                (lastRow || lastCol) &&
                theme.TryResolveColor("lt1", colorMap, out RgbColor conditionalColor))
            {
                color = conditionalColor;
            }
        }

        return new PptxSceneTableCellTextStyle(color, bold);
    }
}
