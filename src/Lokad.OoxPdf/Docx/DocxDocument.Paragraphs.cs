namespace Lokad.OoxPdf.Docx;

internal sealed record DocxParagraph(
    IReadOnlyList<DocxTextRun> Runs,
    IReadOnlyList<DocxInlineImage> Images,
    string? StyleId,
    DocxTextAlignment Alignment,
    string? AlignmentValue,
    double SpacingBeforePoints,
    double SpacingAfterPoints,
    double LineSpacingFactor,
    double? LineSpacingPoints,
    DocxParagraphSpacing Spacing,
    DocxParagraphKeepRules KeepRules,
    DocxListLabel? ListLabel)
{
    public int? OutlineLevel { get; init; }
    public DocxParagraphIndent Indent { get; init; } = DocxParagraphIndent.Empty;
    public IReadOnlyList<DocxTabStop> TabStops { get; init; } = [];
    public bool? SnapToGrid { get; init; }
    public string? SnapToGridValue { get; init; }
    public bool? WordWrap { get; init; }
    public string? WordWrapValue { get; init; }
    public DocxParagraphStyleResolution StyleResolution { get; init; } = DocxParagraphStyleResolution.Empty;
    // Only contextual spacing uses the declared default identity. Keep the
    // authored StyleId and the existing property inheritance intact.
    public string? ContextualSpacingDefaultStyleId { get; init; }
    public IReadOnlyList<DocxInlineReference> InlineReferences { get; init; } = [];
    public IReadOnlyList<DocxCommentRange> CommentRanges { get; init; } = [];
    public IReadOnlyList<DocxRevisionRange> RevisionRanges { get; init; } = [];
    public IReadOnlyList<DocxInlineTextBox> InlineTextBoxes { get; init; } = [];
    public IReadOnlyList<DocxEquation> Equations { get; init; } = [];
    private IReadOnlyList<DocxInlineItem>? inlineContent;
    private IReadOnlyList<DocxTextRun>? inlineContentRuns;
    private IReadOnlyList<DocxInlineImage>? inlineContentImages;
    private IReadOnlyList<DocxEquation>? inlineContentEquations;
    private IReadOnlyList<DocxInlineTextBox>? inlineContentTextBoxes;
    public IReadOnlyList<DocxInlineItem>? InlineContent
    {
        // Transformed legacy paragraphs must rebuild this source view. Do not
        // read stale text/style objects after a coordinator replaces a collection.
        get => ReferenceEquals(inlineContentRuns, Runs) && ReferenceEquals(inlineContentImages, Images) &&
            ReferenceEquals(inlineContentEquations, Equations) && ReferenceEquals(inlineContentTextBoxes, InlineTextBoxes)
                ? inlineContent : null;
        init
        {
            inlineContent = value;
            inlineContentRuns = Runs;
            inlineContentImages = Images;
            inlineContentEquations = Equations;
            inlineContentTextBoxes = InlineTextBoxes;
        }
    }
    public IReadOnlyList<DocxFieldReference> FieldReferences { get; init; } = [];
    public IReadOnlyList<DocxHyperlinkSpan> Hyperlinks { get; init; } = [];
    public IReadOnlyList<DocxBookmarkAnchor> BookmarkAnchors { get; init; } = [];
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
    // Actual empty source runs own review metadata without contributing visible
    // text, font resources or layout metrics.
    public IReadOnlyList<DocxNonRenderingRunOwner> NonRenderingRunOwners { get; init; } = [];
    public bool HasDeletedParagraphMark { get; init; }

    // Hidden deletions/move-from text in Final/Simple views still size Office
    // autofit columns. Original rejects insertions/move-to text and excludes
    // those runs from both painting and this measurement text.
    public string DeletedText { get; init; } = string.Empty;

    // Pilcrow (paragraph-mark) run size through the style cascade, excluding direct
    // run formatting. Office sizes row-end and break-spill spaces at pilcrow size
    // (spill probes 2026-09-28: marker-direct 9pt carries 12pt spills). Null keeps
    // legacy run-size behavior (synthetic and unresolvable paragraphs).
    public double? ParagraphMarkFontSize { get; init; }
    // Comment autofit retains an empty mark run for anchors. Plain comment
    // stories retain a space run for their printed terminal blank's face.
    public DocxTextRun? ParagraphMarkRun { get; init; }

    public DocxEffectiveParagraphProperties EffectiveProperties => new(
        StyleId,
        Alignment,
        AlignmentValue,
        SpacingBeforePoints,
        SpacingAfterPoints,
        LineSpacingFactor,
        LineSpacingPoints,
        Spacing,
        KeepRules,
        Indent,
        TabStops,
        SnapToGrid,
        SnapToGridValue,
        WordWrap,
        WordWrapValue,
        StyleResolution);
}

internal sealed record DocxNonRenderingRunOwner(int SourceRunIndex, IReadOnlyList<DocxRevisionInfo> Revisions);

internal sealed record DocxEffectiveParagraphProperties(
    string? StyleId,
    DocxTextAlignment Alignment,
    string? AlignmentValue,
    double SpacingBeforePoints,
    double SpacingAfterPoints,
    double LineSpacingFactor,
    double? LineSpacingPoints,
    DocxParagraphSpacing Spacing,
    DocxParagraphKeepRules KeepRules,
    DocxParagraphIndent Indent,
    IReadOnlyList<DocxTabStop> TabStops,
    bool? SnapToGrid,
    string? SnapToGridValue,
    bool? WordWrap,
    string? WordWrapValue,
    DocxParagraphStyleResolution StyleResolution);

internal sealed record DocxParagraphStyleResolution(
    string? StyleId,
    bool StyleFound,
    int StyleDepth,
    bool HasDocumentDefaultParagraphProperties,
    bool HasDirectParagraphProperties,
    bool HasTableStyleParagraphProperties)
{
    public static DocxParagraphStyleResolution Empty { get; } = new(null, false, 0, false, false, false);
}

// TextOffset is relative to SourceRunIndex, like emitted text-segment offsets.
internal sealed record DocxBookmarkAnchor(
    string? Id,
    string? Name,
    int SourceRunIndex,
    int TextRunIndex,
    int TextOffset)
{
    // Position in the retained ordered source, before the next inline item.
    // Null preserves the legacy run-relative fallback for unmigrated paragraphs.
    public int? InlineItemIndex { get; init; }
}

internal sealed record DocxInlineReference(
    DocxRelatedStoryKind Kind,
    string? Id,
    string? CustomMarkFollowsValue,
    string? DisplayText,
    int SourceRunIndex,
    int RunChildIndex,
    int TextOffsetInRun)
{
    public string? CustomMarkText { get; init; }
    public DocxRevisionInfo? Revision { get; init; }
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
    public DocxTextRun? DisplayRun { get; init; }
    public string? NoteNumberRestartValue { get; init; }
}

internal sealed record DocxCommentRange(
    string? Id,
    int? StartSourceRunIndex,
    int? StartTextOffset,
    int? EndSourceRunIndex,
    int? EndTextOffset,
    int? ReferenceSourceRunIndex,
    int? ReferenceTextOffset)
{
    public int? StartInlineItemIndex { get; init; }
    public int? EndInlineItemIndex { get; init; }
}

internal sealed record DocxRevisionRange(
    DocxRevisionKind Kind,
    string? Id,
    string? Name,
    string? Author,
    string? Date,
    int? StartSourceRunIndex,
    int? StartTextOffset,
    int? EndSourceRunIndex,
    int? EndTextOffset);

internal sealed record DocxFieldReference(
    DocxFieldKind Kind,
    DocxFieldSourceKind SourceKind,
    string? Instruction,
    string? Placeholder,
    int SourceRunIndex,
    int TextRunIndex,
    int TextRunCount,
    int TextLength)
{
    public bool HasSeparate { get; init; }
    public bool HasCachedResult { get; init; }
    public bool RendersCachedResult { get; init; }
    public bool UsesPlaceholder { get; init; }
    public int NestingDepth { get; init; }
    public int InstructionRunCount { get; init; }
    public int ResultRunCount { get; init; }
}

internal sealed record DocxHyperlinkSpan(
    string? RelationshipId,
    string? Anchor,
    string? Tooltip,
    string? HistoryValue,
    string? Target,
    string? TargetMode,
    string? ResolvedTarget,
    int SourceRunStartIndex,
    int SourceRunCount,
    int TextRunStartIndex,
    int TextRunCount,
    int TextLength)
{
    public int? StartInlineItemIndex { get; init; }
    public int? EndInlineItemIndex { get; init; }
}

internal sealed record DocxTabStop(
    double? PositionPoints,
    string? PositionValue,
    string? Value,
    string? LeaderValue);

internal sealed record DocxParagraphSpacing(
    string? BeforeValue,
    string? AfterValue,
    string? BeforeLinesValue,
    string? AfterLinesValue,
    string? BeforeAutoSpacingValue,
    string? AfterAutoSpacingValue,
    string? LineValue,
    string? LineRuleValue,
    bool? ContextualSpacing)
{
    public static DocxParagraphSpacing Empty { get; } = new(null, null, null, null, null, null, null, null, null);

    public DocxParagraphSpacing Merge(DocxParagraphSpacing other)
    {
        bool hasOtherBeforeSide = HasBeforeSpacingSide(other);
        bool hasOtherAfterSide = HasAfterSpacingSide(other);
        return new DocxParagraphSpacing(
            hasOtherBeforeSide ? other.BeforeValue : BeforeValue,
            hasOtherAfterSide ? other.AfterValue : AfterValue,
            hasOtherBeforeSide ? other.BeforeLinesValue : BeforeLinesValue,
            hasOtherAfterSide ? other.AfterLinesValue : AfterLinesValue,
            hasOtherBeforeSide ? other.BeforeAutoSpacingValue : BeforeAutoSpacingValue,
            hasOtherAfterSide ? other.AfterAutoSpacingValue : AfterAutoSpacingValue,
            other.LineValue ?? LineValue,
            // A local numeric line value with no rule uses automatic spacing.
            // Standalone rule tokens retain their established metadata/geometry.
            other.LineValue is not null ? other.LineRuleValue : other.LineRuleValue ?? LineRuleValue,
            other.ContextualSpacing ?? ContextualSpacing);
    }

    internal static bool HasBeforeSpacingSide(DocxParagraphSpacing spacing)
    {
        return spacing.BeforeValue is not null ||
            spacing.BeforeLinesValue is not null ||
            spacing.BeforeAutoSpacingValue is not null;
    }

    internal static bool HasAfterSpacingSide(DocxParagraphSpacing spacing)
    {
        return spacing.AfterValue is not null ||
            spacing.AfterLinesValue is not null ||
            spacing.AfterAutoSpacingValue is not null;
    }
}

internal sealed record DocxParagraphKeepRules(
    bool? KeepNext,
    string? KeepNextValue,
    bool? KeepLines,
    string? KeepLinesValue,
    bool? WidowControl,
    string? WidowControlValue)
{
    public static DocxParagraphKeepRules Empty { get; } = new(null, null, null, null, null, null);

    public DocxParagraphKeepRules Merge(DocxParagraphKeepRules other)
    {
        return new DocxParagraphKeepRules(
            other.KeepNext ?? KeepNext,
            other.KeepNextValue ?? KeepNextValue,
            other.KeepLines ?? KeepLines,
            other.KeepLinesValue ?? KeepLinesValue,
            other.WidowControl ?? WidowControl,
            other.WidowControlValue ?? WidowControlValue);
    }
}

internal sealed record DocxParagraphIndent(
    double? LeftPoints,
    double? RightPoints,
    double? FirstLinePoints,
    double? HangingPoints,
    string? LeftValue,
    string? RightValue,
    string? FirstLineValue,
    string? HangingValue)
{
    public static DocxParagraphIndent Empty { get; } = new(null, null, null, null, null, null, null, null);

    public DocxParagraphIndent Merge(DocxParagraphIndent other)
    {
        bool hasOtherFirstLineSide = other.FirstLineValue is not null || other.HangingValue is not null;
        return new DocxParagraphIndent(
            other.LeftPoints ?? LeftPoints,
            other.RightPoints ?? RightPoints,
            hasOtherFirstLineSide ? other.FirstLinePoints : FirstLinePoints,
            hasOtherFirstLineSide ? other.HangingPoints : HangingPoints,
            other.LeftValue ?? LeftValue,
            other.RightValue ?? RightValue,
            hasOtherFirstLineSide ? other.FirstLineValue : FirstLineValue,
            hasOtherFirstLineSide ? other.HangingValue : HangingValue);
    }
}

internal sealed record DocxListLabel(
    string Text,
    string FormatValue,
    string LevelTextValue,
    string SuffixValue,
    string NumberId,
    int Level,
    DocxNumberingIndent Indent,
    DocxTextRunStyle Style)
{
    // Decoration and revision ownership are independent of the first text run.
    // Placement-time font/spacing decisions remain on the actual base run.
    public DocxTextRun? DecorationSource { get; init; }
    public DocxNumberingChange? NumberingChange { get; init; }
    public string DisplayText => (NumberingChange?.PreviousText ?? string.Empty) + Text;
}

// Derived original/current counters retain an actual source mark as their
// trigger. They are not new source revisions and must not create balloons.
internal sealed record DocxNumberingChange(string PreviousText, DocxRevisionInfo SourceMark);

internal sealed record DocxNumberingIndent(
    double? LeftPoints,
    double? RightPoints,
    double? FirstLinePoints,
    double? HangingPoints,
    double? NumberingTabPositionPoints,
    string? LeftValue,
    string? RightValue,
    string? FirstLineValue,
    string? HangingValue,
    string? NumberingTabValue,
    string? NumberingTabPositionValue)
{
    public static DocxNumberingIndent Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null);
}

// R17: parsed run vertical alignment. The raw w:val spelling stays on
// VerticalAlignmentValue as provenance; execution switches on the parsed enum so
// superscript/subscript handling cannot drift between measurement and reading.
// Unknown spellings fall back to Baseline, matching the legacy comparisons, which
// only special-cased the two script shifts (measurement compared case-insensitively).
internal enum DocxRunVerticalAlignment
{
    Baseline,
    Superscript,
    Subscript
}

internal sealed record DocxTextRun(
    string Text,
    double FontSize,
    string? ColorHex,
    bool Bold,
    bool Italic,
    bool Underline,
    string? UnderlineValue,
    string? FontFamily,
    double CharacterSpacingPoints,
    bool AllCaps,
    string? VerticalAlignmentValue,
    bool Strike,
    string? StrikeValue,
    bool DoubleStrike,
    string? DoubleStrikeValue,
    string? HighlightValue,
    string? ShadingFillHex,
    string? ShadingValue,
    string? ShadingColor,
    bool SmallCaps,
    string? SmallCapsValue,
    bool Hidden,
    string? HiddenValue,
    string? UnderlineColorHex)
{

    public DocxTextRun(
        string Text,
        double FontSize,
        string? ColorHex,
        bool Bold,
        bool Italic,
        bool Underline,
        string? UnderlineValue,
        string? FontFamily)
        : this(Text, FontSize, ColorHex, Bold, Italic, Underline, UnderlineValue, FontFamily, 0d, false, null, false, null, false, null, null, null, null, null, false, null, false, null, null)
    {
    }

    public DocxTextRun(
        string Text,
        double FontSize,
        string? ColorHex,
        bool Bold,
        bool Italic,
        bool Underline,
        string? UnderlineValue,
        string? FontFamily,
        double CharacterSpacingPoints)
        : this(Text, FontSize, ColorHex, Bold, Italic, Underline, UnderlineValue, FontFamily, CharacterSpacingPoints, false, null, false, null, false, null, null, null, null, null, false, null, false, null, null)
    {
    }
    public DocxRunFonts Fonts { get; init; } = DocxRunFonts.Empty;
    public string? KerningMinimumHalfPointsValue { get; init; }
    public bool PairKerningEnabled { get; init; } = true;
    // Retained from an explicit footnoteRef/endnoteRef in the related story;
    // ordinary text and custom marks never acquire this role from their value/style.
    public bool IsAutomaticNoteStoryMarker { get; init; }
    public bool IsAutomaticNoteReferenceMarker { get; init; }
    // Set only on a qualified Word-compatible body clone; source metadata stays unchanged.
    public bool UseLogicalSizeForPdfAdvances { get; init; }

    // Typed field identity for PAGE/NUMPAGES placeholders. Null for ordinary text, including
    // literal {PAGE}/{NUMPAGES} content, so downstream substitution touches only actual fields.
    public DocxFieldKind? FieldKind { get; init; }

    // Office A/B (w56 superscript plus w57b subscript size curves, Word-COM rendered):
    // style-less unsized super/subscript runs scale glyphs from the 11pt document
    // default, not the 12pt body fallback, while paragraph layout keeps the resolved
    // nominal (max-nominal stays 12 there). Null means nominal (explicit, styled, or
    // table-linked sizes, plus all non-script runs).
    public double? ScriptBaseFontSize { get; init; }
    public DocxRunStyleResolution StyleResolution { get; init; } = DocxRunStyleResolution.Empty;
    public int SourceRunIndex { get; init; } = -1;
    public int SourceTextOffsetInRun { get; init; } = 0;
    public DocxRevisionInfo? Revision { get; init; }
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];

    public DocxEffectiveRunProperties EffectiveProperties => new(
        FontSize,
        ColorHex,
        Bold,
        Italic,
        Underline,
        UnderlineValue,
        FontFamily,
        Fonts,
        CharacterSpacingPoints,
        AllCaps,
        VerticalAlignmentValue,
        Strike,
        StrikeValue,
        DoubleStrike,
        DoubleStrikeValue,
        HighlightValue,
        ShadingFillHex,
        ShadingValue,
        ShadingColor,
        SmallCaps,
        SmallCapsValue,
        Hidden,
        HiddenValue,
        UnderlineColorHex,
        StyleResolution)
    {
        KerningMinimumHalfPointsValue = KerningMinimumHalfPointsValue,
        PairKerningEnabled = PairKerningEnabled
    };

    public DocxRunVerticalAlignment VerticalAlignment => ParseVerticalAlignment(VerticalAlignmentValue);

    public static DocxRunVerticalAlignment ParseVerticalAlignment(string? value)
    {
        if (string.Equals(value, "superscript", StringComparison.OrdinalIgnoreCase))
        {
            return DocxRunVerticalAlignment.Superscript;
        }

        if (string.Equals(value, "subscript", StringComparison.OrdinalIgnoreCase))
        {
            return DocxRunVerticalAlignment.Subscript;
        }

        return DocxRunVerticalAlignment.Baseline;
    }
}

internal sealed record DocxEffectiveRunProperties(
    double FontSize,
    string? ColorHex,
    bool Bold,
    bool Italic,
    bool Underline,
    string? UnderlineValue,
    string? FontFamily,
    DocxRunFonts Fonts,
    double CharacterSpacingPoints,
    bool AllCaps,
    string? VerticalAlignmentValue,
    bool Strike,
    string? StrikeValue,
    bool DoubleStrike,
    string? DoubleStrikeValue,
    string? HighlightValue,
    string? ShadingFillHex,
    string? ShadingValue,
    string? ShadingColor,
    bool SmallCaps,
    string? SmallCapsValue,
    bool Hidden,
    string? HiddenValue,
    string? UnderlineColorHex,
    DocxRunStyleResolution StyleResolution)
{
    public string? KerningMinimumHalfPointsValue { get; init; }
    public bool PairKerningEnabled { get; init; } = true;
    public DocxRunVerticalAlignment VerticalAlignment => DocxTextRun.ParseVerticalAlignment(VerticalAlignmentValue);
}

internal sealed record DocxRunStyleResolution(
    string? CharacterStyleId,
    bool CharacterStyleFound,
    int CharacterStyleDepth,
    bool HasDocumentDefaultRunProperties,
    bool HasParagraphStyleRunProperties,
    bool HasCharacterStyleRunProperties,
    bool HasDirectRunProperties,
    bool HasTableStyleRunProperties)
{
    public static DocxRunStyleResolution Empty { get; } = new(null, false, 0, false, false, false, false, false);
}

internal sealed record DocxRevisionInfo
{
    public DocxRevisionInfo(
        DocxRevisionKind kind,
        string? id,
        string? author,
        string? date,
        string sourceElement,
        DocxRevisionPropertyFamily? propertyChangeFamily,
        IReadOnlyList<string> propertyElementNames)
    {
        Kind = kind;
        Id = id;
        Author = author;
        Date = date;
        SourceElement = sourceElement;
        PropertyChangeFamily = propertyChangeFamily;
        PropertyElementNames = propertyElementNames.ToArray();
    }

    public DocxRevisionKind Kind { get; init; }

    public string? Id { get; init; }

    public string? Author { get; init; }

    public string? Date { get; init; }

    public string SourceElement { get; init; }

    public DocxRevisionPropertyFamily? PropertyChangeFamily { get; init; }

    public IReadOnlyList<string> PropertyElementNames { get; init; }

    public DocxRunFormattingChange? RunFormattingChange { get; init; }
}

// Resolved before/after values, independent of source token spelling and layout.
// Null fields are unchanged; a null record retains the generic revision summary.
internal readonly record struct DocxRunFormattingChange(bool? Bold, bool? Italic);

internal sealed record DocxTextRunStyle(
    double? FontSize,
    string? ColorHex,
    bool? Bold,
    bool? Italic,
    bool? Underline,
    string? UnderlineValue,
    string? FontFamily,
    DocxRunFonts Fonts,
    double? CharacterSpacingPoints,
    bool? AllCaps,
    string? VerticalAlignmentValue,
    bool? Strike,
    string? StrikeValue,
    bool? DoubleStrike,
    string? DoubleStrikeValue,
    string? HighlightValue,
    string? ShadingFillHex,
    string? ShadingValue,
    string? ShadingColor,
    bool? SmallCaps,
    string? SmallCapsValue,
    bool? Hidden,
    string? HiddenValue,
    string? UnderlineColorHex)
{
    public string? KerningMinimumHalfPointsValue { get; init; }

    public DocxTextRunStyle(
        double? FontSize,
        string? ColorHex,
        bool? Bold,
        bool? Italic,
        bool? Underline,
        string? UnderlineValue,
        string? FontFamily,
        DocxRunFonts Fonts)
        : this(FontSize, ColorHex, Bold, Italic, Underline, UnderlineValue, FontFamily, Fonts, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null)
    {
    }
    public static DocxTextRunStyle Empty { get; } = new(null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

    public DocxTextRun ApplyTo(DocxTextRun? baseRun, string text, double fallbackFontSize)
    {
        var source = baseRun ?? new DocxTextRun(string.Empty, fallbackFontSize, null, false, false, false, null, null);
        return new DocxTextRun(
            text,
            FontSize ?? source.FontSize,
            ColorHex ?? source.ColorHex,
            Bold ?? source.Bold,
            Italic ?? source.Italic,
            Underline ?? source.Underline,
            UnderlineValue ?? source.UnderlineValue,
            FontFamily ?? source.FontFamily,
            CharacterSpacingPoints ?? source.CharacterSpacingPoints,
            AllCaps ?? source.AllCaps,
            VerticalAlignmentValue ?? source.VerticalAlignmentValue,
            Strike ?? source.Strike,
            StrikeValue ?? source.StrikeValue,
            DoubleStrike ?? source.DoubleStrike,
            DoubleStrikeValue ?? source.DoubleStrikeValue,
            HighlightValue ?? source.HighlightValue,
            ShadingFillHex ?? source.ShadingFillHex,
            ShadingValue ?? source.ShadingValue,
            ShadingColor ?? source.ShadingColor,
            SmallCaps ?? source.SmallCaps,
            SmallCapsValue ?? source.SmallCapsValue,
            Hidden ?? source.Hidden,
            HiddenValue ?? source.HiddenValue,
            UnderlineColorHex ?? source.UnderlineColorHex)
        {
            Fonts = source.Fonts.Merge(Fonts),
            KerningMinimumHalfPointsValue = KerningMinimumHalfPointsValue ?? source.KerningMinimumHalfPointsValue,
            PairKerningEnabled = source.PairKerningEnabled,
            IsAutomaticNoteReferenceMarker = source.IsAutomaticNoteReferenceMarker,
            UseLogicalSizeForPdfAdvances = source.UseLogicalSizeForPdfAdvances,
            StyleResolution = source.StyleResolution,
            SourceRunIndex = source.SourceRunIndex,
            SourceTextOffsetInRun = source.SourceTextOffsetInRun,
            Revision = source.Revision,
            Revisions = source.Revisions
        };
    }
}

internal sealed record DocxRunFonts(
    string? Ascii,
    string? HighAnsi,
    string? EastAsia,
    string? ComplexScript,
    string? AsciiTheme,
    string? HighAnsiTheme,
    string? EastAsiaTheme,
    string? ComplexScriptTheme)
{
    public static DocxRunFonts Empty { get; } = new(null, null, null, null, null, null, null, null);

    public DocxRunFonts Merge(DocxRunFonts other)
    {
        return new DocxRunFonts(
            other.Ascii ?? Ascii,
            other.HighAnsi ?? HighAnsi,
            other.EastAsia ?? EastAsia,
            other.ComplexScript ?? ComplexScript,
            other.AsciiTheme ?? AsciiTheme,
            other.HighAnsiTheme ?? HighAnsiTheme,
            other.EastAsiaTheme ?? EastAsiaTheme,
            other.ComplexScriptTheme ?? ComplexScriptTheme);
    }
}

// Missing bytes retain the authored frame; painting reports the unresolved asset.
internal sealed record DocxInlineImage(double WidthPoints, double HeightPoints, string ContentType, byte[]? Bytes, string? PartName)
{
    public string? ClickTarget { get; init; }
    // Owning w:r properties for review decoration; an image-only run has no text.
    public DocxTextRun? SourceRunStyle { get; init; }
    public DocxPictureCrop Crop { get; init; }
    public bool FlipHorizontal { get; init; }
    public bool FlipVertical { get; init; }
    public double RotationDegrees { get; init; }
    public DocxPictureFrame? PictureFrame { get; init; }
    // DrawingML SVG extensions can carry an ordinary raster preview. Retain
    // that authored fallback for a missing or unparseable vector image.
    public DocxInlineImage? RasterFallback { get; init; }
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
    // RV05: source run affinity for ordered inline atoms. Images read from a run
    // carry that run index; -1 means the position was not recorded.
    public int SourceRunIndex { get; init; } = -1;
    public int SourceTextOffsetInRun { get; init; }
    public int SourceOrderInRun { get; init; }
}

internal readonly record struct DocxPictureCrop(double Left, double Top, double Right, double Bottom)
{
    public bool IsEmpty => Left == 0d && Top == 0d && Right == 0d && Bottom == 0d;
    public bool IsUsable => double.IsFinite(Left) && double.IsFinite(Top) &&
        double.IsFinite(Right) && double.IsFinite(Bottom) &&
        double.IsFinite(1d - Left - Right) && double.IsFinite(1d - Top - Bottom) &&
        1d - Left - Right > 0d && 1d - Top - Bottom > 0d;
}

internal readonly record struct DocxPictureFrame(double Width, double Height)
{
    public bool IsUsable => double.IsFinite(Width) && double.IsFinite(Height) && Width > 0d && Height > 0d;
}

// An inline DrawingML textbox (wp:inline plus wps:txbx, no blip): unlike floating
// drawings it has no page anchor. Source order lets layout use an inline frame;
// unsupported source owners keep block placement. Extents/insets are file values.
internal sealed record DocxInlineTextBox(
    string? ExtentCxValue,
    string? ExtentCyValue,
    string? TextBoxInsetLeftValue = null,
    string? TextBoxInsetTopValue = null,
    string? TextBoxInsetRightValue = null,
    string? TextBoxInsetBottomValue = null,
    // RV05 inline-box probe (edge-inlinebox, Word COM): stroked inline textbox
    // content starts inside the border like floating boxes.
    string? TextBoxStrokeWidthEmuValue = null,
    // RV05 inline frame stroking: resolved outline color hex (RRGGBB). Null keeps
    // legacy borderless rendering.
    string? TextBoxStrokeColorHex = null)
{
    public int SourceRunIndex { get; init; } = -1;
    public int SourceTextOffsetInRun { get; init; }
    public int SourceOrderInRun { get; init; }
    public IReadOnlyList<DocxBodyElement> BodyElements { get; init; } = [];
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal enum DocxTextAlignment
{
    Left,
    Center,
    Right,
    Justified
}
