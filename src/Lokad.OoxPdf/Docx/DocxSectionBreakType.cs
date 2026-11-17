namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Closed section-break type of a DOCX section break element (w:sectPr/w:type).
/// Closed set — the w:ST_SectionMark vocabulary (absent means Word default nextPage).
/// The reader parses the raw val at the boundary, so layout predicates compare
/// enum members instead of raw strings. Telemetry carriers (snapshots, layout
/// properties) keep the string spelling via ToValueString.
/// </summary>
internal enum DocxSectionBreakType
{
    NextPage,
    NextColumn,
    Continuous,
    EvenPage,
    OddPage
}

internal static class DocxSectionBreakTypeExtensions
{
    /// <summary>
    /// Stable spelling of a section-break type. Kept as the historical OOXML val
    /// spellings so inspection snapshots stay comparable.
    /// </summary>
    public static string ToValueString(this DocxSectionBreakType type)
    {
        return type switch
        {
            DocxSectionBreakType.NextPage => "nextPage",
            DocxSectionBreakType.NextColumn => "nextColumn",
            DocxSectionBreakType.Continuous => "continuous",
            DocxSectionBreakType.EvenPage => "evenPage",
            DocxSectionBreakType.OddPage => "oddPage",
            _ => throw new ArgumentOutOfRangeException(nameof(type), type, "Unknown section-break type.")
        };
    }

    /// <summary>
    /// Parses a w:type val into its section-break type. Null (absent element) stays
    /// null (Word default: next page). Unknown vals also read as null, as Word's repair and
    /// LibreOffice do; a paragraph section break with one still reports
    /// DOCX_UNSUPPORTED_SECTION_BREAK.
    /// </summary>
    public static DocxSectionBreakType? FromValue(string? value)
    {
        return value?.ToLowerInvariant() switch
        {
            null => null,
            "nextpage" => DocxSectionBreakType.NextPage,
            "nextcolumn" => DocxSectionBreakType.NextColumn,
            "continuous" => DocxSectionBreakType.Continuous,
            "evenpage" => DocxSectionBreakType.EvenPage,
            "oddpage" => DocxSectionBreakType.OddPage,
            _ => null
        };
    }
}
