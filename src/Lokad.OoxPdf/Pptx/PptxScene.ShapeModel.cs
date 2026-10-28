using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxSceneSlide(
    int Index,
    string PartName,
    string? MasterPartName,
    string? LayoutPartName,
    XDocument? MasterXml,
    XDocument? LayoutXml,
    XDocument SlideXml,
    IReadOnlyDictionary<string, OoxRelationship> MasterRelationships,
    IReadOnlyDictionary<string, OoxRelationship> LayoutRelationships,
    IReadOnlyDictionary<string, OoxRelationship> SlideRelationships,
    PptxColorMap MasterColorMap,
    PptxColorMap LayoutColorMap,
    PptxColorMap SlideColorMap,
    PptxSceneBackground MasterBackground,
    PptxSceneBackground LayoutBackground,
    PptxSceneBackground SlideBackground,
    bool HasTransition,
    bool HasTiming,
    bool HasOleObject,
    bool ShowMasterShapes,
    IReadOnlyList<PptxSceneNode> MasterNodes,
    IReadOnlyList<PptxSceneNode> LayoutNodes,
    IReadOnlyList<PptxSceneNode> SlideNodes);

internal readonly record struct PptxSceneBackground(
    bool HasFill,
    RgbColor Color,
    double Alpha);

// T02: kind/payload combinations are validated at construction so impossible states
// cannot arise through normal builders. The scene builder is the only producer;
// TextBody stays unconstrained (read for every kind) as does Bounds (absent without
// shape properties). Unknown kinds carry no payload and never reach here (skipped).
internal sealed record PptxSceneNode
{
    public PptxSceneNodeKind Kind { get; }

    public string Id { get; }

    public string Name { get; }

    public bool IsPlaceholder { get; }

    public bool IsSmartArtGraphicFrame { get; }

    public PptxSceneHyperlinkClick HyperlinkClick { get; }

    public PptxSceneBounds? Bounds { get; }

    public PptxSceneShape? Shape { get; }

    public PptxSceneTextBody? TextBody { get; }

    public PptxScenePicture? Picture { get; }

    public PptxSceneTable? Table { get; }

    public PptxSceneChart? Chart { get; }

    public PptxSceneGroupTransform GroupTransform { get; }

    public IReadOnlyList<PptxSceneNode> Children { get; }

    public XElement Source { get; }

    private PptxSceneNode(
        PptxSceneNodeKind kind,
        string id,
        string name,
        bool isPlaceholder,
        bool isSmartArtGraphicFrame,
        PptxSceneHyperlinkClick hyperlinksClick,
        PptxSceneBounds? bounds,
        PptxSceneShape? shape,
        PptxSceneTextBody? textBody,
        PptxScenePicture? picture,
        PptxSceneTable? table,
        PptxSceneChart? chart,
        PptxSceneGroupTransform groupTransform,
        IReadOnlyList<PptxSceneNode> children,
        XElement source)
    {
        Kind = kind;
        Id = id;
        Name = name;
        IsPlaceholder = isPlaceholder;
        IsSmartArtGraphicFrame = isSmartArtGraphicFrame;
        HyperlinkClick = hyperlinksClick;
        Bounds = bounds;
        Shape = shape;
        TextBody = textBody;
        Picture = picture;
        Table = table;
        Chart = chart;
        GroupTransform = groupTransform;
        Children = children;
        Source = source;
    }

    public static PptxSceneNode Create(
        PptxSceneNodeKind kind,
        string id,
        string name,
        bool isPlaceholder,
        bool isSmartArtGraphicFrame,
        PptxSceneHyperlinkClick hyperlinksClick,
        PptxSceneBounds? bounds,
        PptxSceneShape? shape,
        PptxSceneTextBody? textBody,
        PptxScenePicture? picture,
        PptxSceneTable? table,
        PptxSceneChart? chart,
        PptxSceneGroupTransform groupTransform,
        IReadOnlyList<PptxSceneNode> children,
        XElement source)
    {
        ArgumentNullException.ThrowIfNull(id);
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(children);
        ArgumentNullException.ThrowIfNull(source);
        if ((kind is PptxSceneNodeKind.Shape or PptxSceneNodeKind.Connector) && shape is null)
        {
            throw new ArgumentException("Shape nodes require a shape payload.", nameof(shape));
        }

        if (kind == PptxSceneNodeKind.Picture && picture is null)
        {
            throw new ArgumentException("Picture nodes require a picture payload.", nameof(picture));
        }

        if (kind == PptxSceneNodeKind.Table && table is null)
        {
            throw new ArgumentException("Table nodes require a table payload.", nameof(table));
        }

        if (kind == PptxSceneNodeKind.Chart && chart is null)
        {
            throw new ArgumentException("Chart nodes require a chart payload.", nameof(chart));
        }

        if (shape is not null && kind is not (PptxSceneNodeKind.Shape or PptxSceneNodeKind.Connector))
        {
            throw new ArgumentException("Shape payloads belong to shape or connector nodes.", nameof(shape));
        }

        if (picture is not null && kind != PptxSceneNodeKind.Picture)
        {
            throw new ArgumentException("Picture payloads belong to picture nodes.", nameof(picture));
        }

        if (table is not null && kind != PptxSceneNodeKind.Table)
        {
            throw new ArgumentException("Table payloads belong to table nodes.", nameof(table));
        }

        if (chart is not null && kind != PptxSceneNodeKind.Chart)
        {
            throw new ArgumentException("Chart payloads belong to chart nodes.", nameof(chart));
        }

        if (children.Count != 0 && kind != PptxSceneNodeKind.Group)
        {
            throw new ArgumentException("Only group nodes carry children.", nameof(children));
        }

        return new PptxSceneNode(kind, id, name, isPlaceholder, isSmartArtGraphicFrame, hyperlinksClick, bounds, shape, textBody, picture, table, chart, groupTransform, children, source);
    }
}

internal readonly record struct PptxSceneHyperlinkClick(
    bool IsDefined,
    string? RelationshipId,
    string? Action);

internal sealed record PptxSceneShape(
    string Preset,
    IReadOnlyDictionary<string, double> PresetAdjustments,
    bool HasCustomGeometry,
    PptxSceneCustomGeometry CustomGeometry,
    PptxFormatSchemeReference FillReference,
    PptxFormatSchemeReference LineReference,
    bool NoFill,
    bool LineNoFill,
    PptxSceneFillStyle Fill,
    PptxSceneGradientFill GradientFill,
    PptxScenePatternFill PatternFill,
    PptxSceneShapePictureFill PictureFill,
    bool HasUnsupportedTransparency,
    PptxSceneGlow Glow,
    PptxSceneOuterShadow OuterShadow,
    PptxSceneShapeEffectFamily Effects,
    PptxSceneLineStyle Line,
    PptxSceneLineEnd HeadEnd,
    PptxSceneLineEnd TailEnd);

internal sealed record PptxSceneCustomGeometry(
    bool HasGeometry,
    bool HasUnsupportedGeometry,
    IReadOnlyList<PptxSceneCustomGuide> Guides,
    IReadOnlyList<PptxSceneCustomPath> Paths);

internal sealed record PptxSceneCustomGuide(
    string Name,
    string Formula);

internal sealed record PptxSceneCustomPath(
    double Width,
    double Height,
    bool AllowsFill,
    bool AllowsStroke,
    IReadOnlyList<PptxSceneCustomCommand> Commands);

internal sealed record PptxSceneCustomCommand(
    PptxSceneCustomCommandKind Kind,
    IReadOnlyList<PptxSceneCustomPoint> Points,
    string RadiusX,
    string RadiusY,
    string StartAngle,
    string SweepAngle);

internal enum PptxSceneCustomCommandKind
{
    MoveTo,
    LineTo,
    CubicBezierTo,
    QuadraticBezierTo,
    ArcTo,
    Close
}

internal readonly record struct PptxSceneCustomPoint(
    string X,
    string Y);

internal readonly record struct PptxSceneFillStyle(
    bool HasFill,
    RgbColor Color,
    double Alpha);

internal sealed record PptxSceneGradientFill(
    bool HasGradient,
    bool HasGradientSource,
    bool HasUnsupportedGradient,
    double AngleDegrees,
    IReadOnlyList<PptxSceneGradientStop> Stops);

internal readonly record struct PptxSceneGradientStop(
    double Offset,
    RgbColor Color,
    double Alpha);

internal readonly record struct PptxScenePatternFill(
    bool HasPattern,
    bool HasPatternSource,
    bool HasUnsupportedPattern,
    string Preset,
    RgbColor Foreground,
    RgbColor Background,
    double Alpha);

internal readonly record struct PptxSceneShapePictureFill(
    bool HasPicture,
    string RelationshipId,
    string? TargetPartName,
    PptxSceneImageResource? Resource,
    PptxSceneRect Crop,
    PptxSceneRect Fill,
    double Alpha,
    string? AlphaValue,
    PptxScenePictureTile Tile);

internal readonly record struct PptxSceneGlow(
    bool HasGlow,
    RgbColor Color,
    double Alpha,
    double Radius);

internal readonly record struct PptxSceneOuterShadow(
    bool HasShadow,
    RgbColor Color,
    double Alpha,
    double OffsetX,
    double OffsetY,
    double BlurRadius);

internal readonly record struct PptxSceneShapeEffectFamily(
    bool HasEffectList,
    bool HasEffectDag,
    IReadOnlyList<string> UnsupportedEffectNames);

internal readonly record struct PptxSceneLineStyle(
    bool HasLine,
    RgbColor Color,
    double Width,
    double Alpha,
    IReadOnlyList<double> DashPattern,
    string? DashPreset,
    PptxSceneLineCompound? Compound,
    string? CompoundValue,
    int? Cap,
    string? CapValue,
    int? Join,
    string? JoinValue,
    bool WidthSpecified,
    bool NoFill = false)
{
    public bool HasDash => DashPattern is { Count: > 0 };
}

internal enum PptxSceneLineCompound
{
    Single,
    Double,
    ThickThin,
    ThinThick,
    Triple
}

internal enum PptxSceneLineEndKind
{
    None,
    Triangle,
    Arrow,
    Stealth,
    Diamond,
    Oval
}

internal readonly record struct PptxSceneLineEnd(
    PptxSceneLineEndKind Kind,
    string? TypeValue,
    double WidthScale,
    string? WidthValue,
    double LengthScale,
    string? LengthValue)
{
    public bool IsNone => Kind == PptxSceneLineEndKind.None;
}

internal sealed record PptxScenePicture(
    string? RelationshipId,
    string? TargetPartName,
    PptxSceneImageResource? Resource,
    PptxSceneRect Crop,
    PptxSceneRect Fill,
    double Alpha,
    string? AlphaValue,
    PptxSceneImageRecolor Recolor,
    bool HasVideo,
    bool HasAudio,
    PptxScenePictureTile Tile,
    PptxSceneLineStyle Line,
    PptxSceneOuterShadow OuterShadow,
    string Preset,
    IReadOnlyDictionary<string, double> PresetAdjustments,
    bool HasCustomGeometry,
    PptxSceneCustomGeometry CustomGeometry);

internal sealed record PptxSceneImageResource(
    string PartName,
    string ContentType,
    byte[] Bytes);

internal sealed record PptxScenePackageResource(
    string PartName,
    string ContentType,
    byte[] Bytes);

internal readonly record struct PptxSceneGroupTransform(
    long OffsetX,
    long OffsetY,
    long Width,
    long Height,
    long ChildOffsetX,
    long ChildOffsetY,
    double ScaleX,
    double ScaleY,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical)
{
    public static PptxSceneGroupTransform Identity { get; } = new(0, 0, 0, 0, 0, 0, 1d, 1d, 0d, FlipHorizontal: false, FlipVertical: false);
}

internal readonly record struct PptxSceneRect(
    double Left,
    double Top,
    double Right,
    double Bottom,
    string? LeftValue,
    string? TopValue,
    string? RightValue,
    string? BottomValue)
{
    public bool IsEmpty => Left == 0d && Top == 0d && Right == 0d && Bottom == 0d;
}
