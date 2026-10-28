using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal readonly record struct PptxScenePictureTile(
    bool HasTile,
    string? TileValue,
    string? AlignmentValue,
    string? FlipValue,
    string? ScaleXValue,
    string? ScaleYValue,
    string? OffsetXValue,
    string? OffsetYValue);

internal enum PptxSceneImageRecolorKind
{
    None,
    Luminance,
    Duotone,
    Grayscale,
    BiLevel,
    ColorChange
}

internal readonly record struct PptxSceneImageRecolor(
    PptxSceneImageRecolorKind Kind,
    double Brightness,
    double Contrast,
    RgbColor Dark,
    RgbColor Light,
    double Threshold,
    string? KindValue,
    string? BrightnessValue,
    string? ContrastValue,
    string? ThresholdValue)
{
    public static PptxSceneImageRecolor None { get; } = new(PptxSceneImageRecolorKind.None, 0d, 0d, default, default, 0d, null, null, null, null);

    public static PptxSceneImageRecolor Luminance(
        double brightness,
        double contrast,
        string? kindValue,
        string? brightnessValue,
        string? contrastValue)
    {
        return new PptxSceneImageRecolor(
            PptxSceneImageRecolorKind.Luminance,
            Math.Clamp(brightness, -1d, 1d),
            Math.Clamp(contrast, -1d, 1d),
            default,
            default,
            0d,
            kindValue,
            brightnessValue,
            contrastValue,
            null);
    }

    public static PptxSceneImageRecolor Duotone(RgbColor dark, RgbColor light, string? kindValue)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.Duotone, 0d, 0d, dark, light, 0d, kindValue, null, null, null);
    }

    public static PptxSceneImageRecolor Grayscale(string? kindValue)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.Grayscale, 0d, 0d, default, default, 0d, kindValue, null, null, null);
    }

    // Dark holds the source colour, Light the replacement colour and Threshold the
    // replacement alpha; the colour-change recolor needs no other fields.
    public static PptxSceneImageRecolor ColorChange(RgbColor from, RgbColor to, double toAlpha, string? kindValue)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.ColorChange, 0d, 0d, from, to, Math.Clamp(toAlpha, 0d, 1d), kindValue, null, null, null);
    }

    public static PptxSceneImageRecolor BiLevel(double threshold, string? kindValue, string? thresholdValue)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.BiLevel, 0d, 0d, default, default, Math.Clamp(threshold, 0d, 1d), kindValue, null, null, thresholdValue);
    }
}

internal sealed record PptxSceneTextBody(
    XElement? BodyProperties,
    XElement? ListStyle,
    bool HasUnsupportedTextOrientation,
    bool HasUnsupportedVerticalOverflow,
    IReadOnlyList<PptxSceneTextParagraph> Paragraphs);

internal sealed record PptxSceneCascadeLayer(
    string Name,
    string Kind,
    XElement? Source);

internal sealed record PptxSceneTextParagraph(
    XElement? Properties,
    XElement? EndParagraphProperties,
    int Level,
    IReadOnlyList<PptxSceneTextRun> Runs,
    IReadOnlyList<PptxSceneCascadeLayer> CascadeLayers,
    XElement? Source);

internal sealed record PptxSceneTextRun(
    PptxSceneTextRunKind Kind,
    string Text,
    XElement? Properties,
    XElement Source);

internal enum PptxSceneTextRunKind
{
    Text,
    Break,
    Field
}

internal sealed record PptxSceneBounds(
    long XEmu,
    long YEmu,
    long WidthEmu,
    long HeightEmu,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical)
{
    public double X => OoxUnits.EmuToPoints(XEmu);
    public double Y => OoxUnits.EmuToPoints(YEmu);
    public double Width => OoxUnits.EmuToPoints(WidthEmu);
    public double Height => OoxUnits.EmuToPoints(HeightEmu);
}

internal enum PptxSceneNodeKind
{
    Shape,
    Picture,
    Table,
    Chart,
    Group,
    Connector,
    UnknownGraphicFrame,
    Unknown
}
