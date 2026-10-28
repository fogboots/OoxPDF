using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxSceneTable(
    IReadOnlyList<double> ColumnWidths,
    IReadOnlyList<double> RowHeights,
    IReadOnlyList<PptxSceneTableRow> Rows,
    PptxSceneTableStyle Style,
    XElement? Source);

internal readonly record struct PptxSceneTableStyle(
    string? StyleId,
    string Name,
    PptxBuiltInTableStyleKind Kind,
    string Accent,
    bool IsSupported,
    bool FirstRow,
    string FirstRowValue,
    bool LastRow,
    string LastRowValue,
    bool FirstColumn,
    string FirstColumnValue,
    bool LastColumn,
    string LastColumnValue,
    bool BandRow,
    string BandRowValue,
    bool BandColumn,
    string BandColumnValue,
    XElement? Definition = null)
{
    public bool HasStyle => StyleId is not null;
}

internal enum PptxBuiltInTableStyleKind
{
    Unknown,
    LightStyle1,
    DarkStyle1,
    MediumStyle2
}

internal readonly record struct PptxBuiltInTableStyle(string Name, PptxBuiltInTableStyleKind Kind, string Accent);

internal static class PptxBuiltInTableStyles
{
    public static bool TryGet(string? styleId, out PptxBuiltInTableStyle style)
    {
        return Styles.TryGetValue(styleId ?? string.Empty, out style);
    }

    private static IReadOnlyDictionary<string, PptxBuiltInTableStyle> Styles { get; } =
        new Dictionary<string, PptxBuiltInTableStyle>(StringComparer.OrdinalIgnoreCase)
        {
            ["{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}"] = new("Light-Style-1", PptxBuiltInTableStyleKind.LightStyle1, "tx1"),
            ["{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"] = new("Light-Style-1", PptxBuiltInTableStyleKind.LightStyle1, "accent1"),
            ["{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}"] = new("Light-Style-1", PptxBuiltInTableStyleKind.LightStyle1, "accent2"),
            ["{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}"] = new("Light-Style-1", PptxBuiltInTableStyleKind.LightStyle1, "accent3"),
            ["{D27102A9-8310-4765-A935-A1911B00CA55}"] = new("Light-Style-1", PptxBuiltInTableStyleKind.LightStyle1, "accent4"),
            ["{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}"] = new("Light-Style-1", PptxBuiltInTableStyleKind.LightStyle1, "accent5"),
            ["{68D230F3-CF80-4859-8CE7-A43EE81993B5}"] = new("Light-Style-1", PptxBuiltInTableStyleKind.LightStyle1, "accent6"),
            ["{E8034E78-7F5D-4C2E-B375-FC64B27BC917}"] = new("Dark-Style-1", PptxBuiltInTableStyleKind.DarkStyle1, "dk1"),
            ["{125E5076-3810-47DD-B79F-674D7AD40C01}"] = new("Dark-Style-1", PptxBuiltInTableStyleKind.DarkStyle1, "accent1"),
            ["{37CE84F3-28C3-443E-9E96-99CF82512B78}"] = new("Dark-Style-1", PptxBuiltInTableStyleKind.DarkStyle1, "accent2"),
            ["{D03447BB-5D67-496B-8E87-E561075AD55C}"] = new("Dark-Style-1", PptxBuiltInTableStyleKind.DarkStyle1, "accent3"),
            ["{E929F9F4-4A8F-4326-A1B4-22849713DDAB}"] = new("Dark-Style-1", PptxBuiltInTableStyleKind.DarkStyle1, "accent4"),
            ["{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}"] = new("Dark-Style-1", PptxBuiltInTableStyleKind.DarkStyle1, "accent5"),
            ["{AF606853-7671-496A-8E4F-DF71F8EC918B}"] = new("Dark-Style-1", PptxBuiltInTableStyleKind.DarkStyle1, "accent6"),
            ["{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"] = new("Medium-Style-2", PptxBuiltInTableStyleKind.MediumStyle2, "tx1"),
            ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"] = new("Medium-Style-2", PptxBuiltInTableStyleKind.MediumStyle2, "accent1"),
            ["{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}"] = new("Medium-Style-2", PptxBuiltInTableStyleKind.MediumStyle2, "accent2"),
            ["{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}"] = new("Medium-Style-2", PptxBuiltInTableStyleKind.MediumStyle2, "accent3"),
            ["{00A15C55-8517-42AA-B614-E9B94910E393}"] = new("Medium-Style-2", PptxBuiltInTableStyleKind.MediumStyle2, "accent4"),
            ["{7DF18680-E054-41AD-8BC1-D1AEF772440D}"] = new("Medium-Style-2", PptxBuiltInTableStyleKind.MediumStyle2, "accent5"),
            ["{93296810-A885-4BE3-A3E7-6D5BEEA58F35}"] = new("Medium-Style-2", PptxBuiltInTableStyleKind.MediumStyle2, "accent6")
        };
}

internal sealed record PptxSceneTableRow(IReadOnlyList<PptxSceneTableCell> Cells);

internal readonly record struct PptxSceneTableCell(
    int ColumnSpan,
    int RowSpan,
    bool IsMergedContinuation,
    PptxSceneTextInsets TextInsets,
    PptxSceneTableCellTextInsetSources TextInsetSources,
    PptxSceneTextInsetValues TextInsetValues,
    PptxSceneTableCellVerticalAnchor VerticalAnchor,
    string? VerticalAnchorValue,
    PptxSceneTableCellVerticalAnchorSource VerticalAnchorSource,
    PptxSceneFillStyle Fill,
    PptxSceneTableCellBorders Borders,
    PptxSceneFillStyle StyleFill,
    PptxSceneTableCellTextStyle StyleText,
    XElement? TextBody,
    XElement? LayoutTextBody,
    bool HasUnsupportedTextOrientation,
    bool HasUnsupportedVerticalOverflow,
    int LeadingEmptyTextParagraphCount);

internal readonly record struct PptxSceneTableCellTextStyle(RgbColor? Color, bool Bold, bool Italic = false);

internal readonly record struct PptxSceneTableCellTextInsetSources(
    PptxSceneTableCellTextInsetSource Left,
    PptxSceneTableCellTextInsetSource Right,
    PptxSceneTableCellTextInsetSource Top,
    PptxSceneTableCellTextInsetSource Bottom);

internal enum PptxSceneTableCellTextInsetSource
{
    Default,
    BodyProperties,
    CellProperties
}

internal readonly record struct PptxSceneTableCellBorders(
    PptxSceneTableCellBorder Left,
    PptxSceneTableCellBorder Right,
    PptxSceneTableCellBorder Top,
    PptxSceneTableCellBorder Bottom)
{
    public bool HasExplicitBorder => Left.IsSpecified || Right.IsSpecified || Top.IsSpecified || Bottom.IsSpecified;
}

internal readonly record struct PptxSceneTableCellBorder(bool IsSpecified, PptxSceneLineStyle Line);

internal readonly record struct PptxSceneTextInsets(double Left, double Right, double Top, double Bottom);

internal readonly record struct PptxSceneTextInsetValues(string? Left, string? Right, string? Top, string? Bottom);

internal enum PptxSceneTableCellVerticalAnchor
{
    Top,
    Middle,
    Bottom
}

internal enum PptxSceneTableCellVerticalAnchorSource
{
    Default,
    CellProperties
}
