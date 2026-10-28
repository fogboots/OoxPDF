using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private const int OfficeArrowTailConnectorSamplesPerSegment = 61;
    private const int OfficeTriangleTailConnectorSamplesPerSegment = 54;
    private const int OfficeStealthTailConnectorSamplesPerSegment = 66;
    private const double OfficeArrowheadLengthFactor = 4.423333d;
    private const double OfficeTriangleTailMinimumLength = 5d;
    private const double OfficeTriangleTailLengthFactor = 3.5d;
    private const double OfficeTriangleTailHalfWidthFactor = 0.45d;
    private const double OfficeStraightTriangleLineEndLengthFactor = 4d;
    private const double OfficeStraightTriangleLineEndHalfWidthFactor = 1.5d;
    private const double OfficeStraightConnectorTriangleLineEndHalfWidthFactor = 2d;
    private const double OfficeStraightTriangleLineEndOverlapFactor = 2d / 3d;
    private const double OfficeStraightStealthLineEndLengthFactor = 3d;
    private const double OfficeStraightStealthLineEndWidthFactor = 3d;
    private const double OfficeStraightStealthLineEndMinimumSize = 6d;
    private const double OfficeStraightStealthLineEndNotchFactor = 2d / 3d;
    private const double OfficePresetArcDefaultStartGuide = 16200000d;
    private const double OfficePresetArcDefaultEndGuide = 0d;
    private const double OfficePresetArcFlatteningTolerance = 0.015d;
    private const double OfficePresetArcMaximumFlatteningDegrees = 4.5d;
    private const double OfficeGlowRasterPixelsPerPoint = 1d;
    private const int OfficeGlowRasterMaxPixelsPerSide = 2048;
    private const double OfficeOuterShadowRasterExtentFactor = 1.25d;

    private static bool RenderBackground(PptxRenderContext context, PptxSceneBackground background, PdfGraphicsBuilder graphics, bool defaultWhenMissing)
    {
        if (background.HasFill)
        {
            graphics.SaveState();
            ClipSlideBoundsEvenOdd(context.Document, graphics);
            if (background.Alpha < 0.999d)
            {
                graphics.SetAlpha(background.Alpha, 1d);
            }

            graphics.SetFillRgb(background.Color.Red, background.Color.Green, background.Color.Blue);
            graphics.FillRectangleEvenOdd(0, 0, context.Document.SlideWidthPoints, context.Document.SlideHeightPoints);
            graphics.RestoreState();
            return true;
        }

        if (!defaultWhenMissing)
        {
            return false;
        }

        graphics.SaveState();
        ClipSlideBoundsEvenOdd(context.Document, graphics);
        graphics.SetFillRgb(255, 255, 255);
        graphics.FillRectangleEvenOdd(0, 0, context.Document.SlideWidthPoints, context.Document.SlideHeightPoints);
        graphics.RestoreState();
        return true;
    }

    private static void RenderShape(
        PptxSceneNode shape,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        PptxTheme theme,
        GroupTransform groupTransform,
        List<PdfImageResource>? images,
        Dictionary<string, PdfImageXObject?>? imageCache,
        ref int imageIndex)
    {
        if (shape.Shape is null || shape.Bounds is null)
        {
            return;
        }

        RenderShape(
            shape.Source,
            document,
            graphics,
            diagnosticSink,
            slideIndex,
            theme,
            groupTransform,
            images,
            imageCache,
            ref imageIndex,
            ToShapeBounds(shape.Bounds),
            shape.Shape.Preset,
            shape.Shape.PresetAdjustments,
            shape.Shape.HasCustomGeometry,
            shape.Shape.CustomGeometry,
            shape.Shape.NoFill,
            ToFillStyle(shape.Shape.Fill),
            ToGradientFill(shape.Shape.GradientFill),
            ToShapePatternFill(shape.Shape.PatternFill),
            ToShapePictureFill(shape.Shape.PictureFill),
            ToGlow(shape.Shape.Glow),
            ToOuterShadow(shape.Shape.OuterShadow),
            ToLineStyle(shape.Shape.Line),
            ToLineEndStyle(shape.Shape.HeadEnd),
            ToLineEndStyle(shape.Shape.TailEnd));
    }

    private static void RenderShape(
        XElement shape,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        PptxTheme theme,
        GroupTransform groupTransform,
        List<PdfImageResource>? images,
        Dictionary<string, PdfImageXObject?>? imageCache,
        ref int imageIndex,
        ShapeBounds rawBounds,
        string preset,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride,
        bool hasCustomGeometry,
        PptxSceneCustomGeometry? customGeometryOverride,
        bool noFillOverride,
        FillStyle? fillOverride,
        GradientFill? gradientFillOverride,
        ShapePatternFill? patternFillOverride,
        ShapePictureFill? pictureFillOverride,
        Glow? glowOverride,
        OuterShadow? outerShadowOverride,
        LineStyle? lineOverride,
        LineEndStyle? headEndOverride,
        LineEndStyle? tailEndOverride)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        if (shapeProperties is null)
        {
            return;
        }

        ShapeBounds bounds = groupTransform.Apply(rawBounds);

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        bool transformed = bounds.RotationDegrees != 0d || bounds.FlipHorizontal || bounds.FlipVertical;

        RgbColor fill;
        double fillAlpha;
        bool hasFill;
        if (fillOverride is { } resolvedFill)
        {
            fill = resolvedFill.Color;
            fillAlpha = resolvedFill.Alpha;
            hasFill = resolvedFill.HasFill;
        }
        else if (noFillOverride)
        {
            fill = default;
            fillAlpha = 1d;
            hasFill = false;
        }
        else
        {
            fill = default;
            fillAlpha = 1d;
            hasFill = false;
        }
        GradientFill? gradientFill = gradientFillOverride;
        bool hasPatternFill = patternFillOverride is not null;
        ShapePatternFill patternFill;
        patternFill = patternFillOverride ?? default;
        RgbColor stroke;
        double lineWidth;
        double strokeAlpha;
        bool hasStroke;
        if (lineOverride is { } line)
        {
            stroke = line.Color;
            lineWidth = line.Width;
            strokeAlpha = line.Alpha;
            hasStroke = line.HasLine;
        }
        else
        {
            hasStroke = TryReadShapeLine(shape, shapeProperties, theme, out stroke, out lineWidth, out strokeAlpha);
        }
        bool hasDash;
        IReadOnlyList<double> dashPattern;
        int? lineCap;
        int? lineJoin;
        if (lineOverride is { } resolvedLine)
        {
            hasDash = resolvedLine.HasDash;
            dashPattern = resolvedLine.DashPattern;
            lineCap = resolvedLine.Cap;
            lineJoin = resolvedLine.Join;
        }
        else
        {
            hasDash = TryReadPresetDash(shapeProperties, lineWidth, out dashPattern);
            lineCap = ReadLineCap(shapeProperties) switch
            {
                "rnd" => 1,
                "sq" => 2,
                _ => null
            };
            lineJoin = ReadLineJoin(shapeProperties);
        }
        LineEndStyle headEnd = headEndOverride ?? ReadLineEnd(shapeProperties, "headEnd");
        LineEndStyle tailEnd = tailEndOverride ?? ReadLineEnd(shapeProperties, "tailEnd");
        bool hasPictureFill = TryReadShapePictureFill(
            shapeProperties,
            diagnosticSink,
            slideIndex,
            images,
            imageCache,
            ref imageIndex,
            pictureFillOverride,
            out string? pictureFillName,
            out PdfImageXObject? pictureFillImage,
            out CropRect pictureFillCrop,
            out FillRect pictureFillRect,
            out double pictureFillAlpha);
        PptxSceneCustomGeometry? sceneCustomGeometry = customGeometryOverride is { HasGeometry: true } ? customGeometryOverride : null;

        if (transformed)
        {
            graphics.SaveState();
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }

        bool hasOuterShadow = outerShadowOverride is not null;
        OuterShadow outerShadow;
        outerShadow = outerShadowOverride ?? default;

        if (glowOverride is { } glow &&
            CanRenderGlowPreset(preset) &&
            images is not null)
        {
            double glowCornerRadius = preset == "roundRect" ? ReadRoundRectangleRadius(shapeProperties, presetAdjustmentsOverride, width, height) : 0d;
            DrawRasterGlow(graphics, x, y, width, height, glow, images, ref imageIndex, preset, glowCornerRadius);
        }

        if (hasOuterShadow &&
            preset is not ("line" or "straightConnector1" or "curvedConnector2" or "curvedConnector3") &&
            sceneCustomGeometry is null && !hasCustomGeometry)
        {
            if (outerShadow.BlurRadius > 0d &&
                preset == "rect" &&
                images is not null)
            {
                DrawRasterOuterShadow(graphics, x, y, width, height, outerShadow, images, ref imageIndex);
            }
            else
            {
                DrawOuterShadow(graphics, shapeProperties, preset, x, y, width, height, outerShadow, presetAdjustmentsOverride);
            }
        }

        if (sceneCustomGeometry is not null && TryRenderCustomGeometry(
                sceneCustomGeometry,
                graphics,
                x,
                y,
                width,
                height,
                hasFill,
                fill,
                fillAlpha,
                hasStroke,
                stroke,
                lineWidth,
                strokeAlpha,
                hasDash,
                dashPattern,
                lineCap,
                lineJoin,
                headEnd,
                tailEnd))
        {
            if (transformed)
            {
                graphics.RestoreState();
            }

            return;
        }

        if (preset is "line" or "straightConnector1")
        {
            if (hasStroke)
            {
                bool transparentStroke = strokeAlpha < 0.999d;
                if (transparentStroke)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(strokeAlpha, strokeAlpha);
                }

                double x1 = x;
                double y1 = document.SlideHeightPoints - yTop;
                double x2 = x + width;
                double y2 = document.SlideHeightPoints - yTop - height;
                bool hasHeadArrow = IsFilledTriangleArrow(headEnd);
                bool hasTailArrow = IsFilledTriangleArrow(tailEnd);
                bool hasStealthEnd = headEnd.Kind == LineEndKind.Stealth || tailEnd.Kind == LineEndKind.Stealth;
                // Office fills stealth-tailed straight bodies (seven renders); unflipped straight
                // connectors take the filled body too while transformed (grouped or rotated)
                // shapes keep the legacy stroked path for lack of Office evidence there.
                bool useFilledStealthBody = string.Equals(preset, "line", StringComparison.Ordinal) ||
                    rawBounds.FlipHorizontal ||
                    rawBounds.FlipVertical ||
                    (string.Equals(preset, "straightConnector1", StringComparison.Ordinal) && !transformed);
                if (hasStealthEnd && useFilledStealthBody && headEnd.Kind is LineEndKind.None or LineEndKind.Stealth && tailEnd.Kind is LineEndKind.None or LineEndKind.Stealth && !hasDash && lineCap is null)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    FillStealthEndedLine(graphics, x1, y1, x2, y2, lineWidth, headEnd, tailEnd);
                }
                else if ((hasHeadArrow || hasTailArrow) && headEnd.Kind is LineEndKind.None or LineEndKind.Triangle or LineEndKind.Arrow && tailEnd.Kind is LineEndKind.None or LineEndKind.Triangle or LineEndKind.Arrow && !hasDash && lineCap is null)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    bool usesOfficeArrowType = headEnd.Kind == LineEndKind.Arrow || tailEnd.Kind == LineEndKind.Arrow;
                    if (usesOfficeArrowType)
                    {
                        FillOfficeArrowedLine(graphics, x1, y1, x2, y2, lineWidth, hasHeadArrow, hasTailArrow);
                    }
                    else
                    {
                        double arrowHalfWidthFactor = string.Equals(preset, "straightConnector1", StringComparison.Ordinal)
                            ? OfficeStraightConnectorTriangleLineEndHalfWidthFactor
                            : OfficeStraightTriangleLineEndHalfWidthFactor;
                        FillArrowedLine(graphics, x1, y1, x2, y2, lineWidth, hasHeadArrow, hasTailArrow, arrowHalfWidthFactor);
                    }
                }
                else
                {
                    graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
                    graphics.SetLineWidth(lineWidth);
                    if (hasDash)
                    {
                        graphics.SetLineDash(dashPattern);
                    }

                    if (lineCap is { } cap)
                    {
                        graphics.SetLineCap(cap);
                        graphics.SetLineJoin(1);
                    }

                    graphics.StrokeLine(x1, y1, x2, y2);
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    FillLineEndMarker(graphics, headEnd, x1, y1, x1 - x2, y1 - y2, lineWidth);
                    FillLineEndMarker(graphics, tailEnd, x2, y2, x2 - x1, y2 - y1, lineWidth);
                    if (hasDash)
                    {
                        graphics.ClearLineDash();
                    }

                    if (lineCap is not null)
                    {
                        graphics.SetLineCap(0);
                        graphics.SetLineJoin(0);
                    }
                }

                if (transparentStroke)
                {
                    graphics.RestoreState();
                }
            }

            if (transformed)
            {
                graphics.RestoreState();
            }

            return;
        }

        if (preset is "curvedConnector2" or "curvedConnector3")
        {
            if (hasStroke)
            {
                bool transparentStroke = strokeAlpha < 0.999d;
                if (transparentStroke)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(strokeAlpha, strokeAlpha);
                }

                graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
                graphics.SetLineWidth(lineWidth);
                if (hasDash)
                {
                    graphics.SetLineDash(dashPattern);
                }

                if (lineCap is { } cap)
                {
                    graphics.SetLineCap(cap);
                    graphics.SetLineJoin(1);
                }

                if (!hasDash &&
                    lineCap is null &&
                    headEnd.IsNone &&
                    tailEnd.Kind is LineEndKind.Triangle or LineEndKind.Arrow or LineEndKind.Stealth)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    if (!TryFillCurvedConnectorPreset(
                        graphics,
                        shapeProperties,
                        preset,
                        x,
                        yTop,
                        width,
                        height,
                        document.SlideHeightPoints,
                        lineWidth,
                        tailEnd,
                        presetAdjustmentsOverride))
                    {
                        DrawCurvedConnectorPreset(
                            graphics,
                            shapeProperties,
                            preset,
                            x,
                            yTop,
                            width,
                            height,
                            document.SlideHeightPoints,
                            stroke,
                            lineWidth,
                            headEnd,
                            tailEnd,
                            presetAdjustmentsOverride);
                    }
                }
                else
                {
                    DrawCurvedConnectorPreset(
                        graphics,
                        shapeProperties,
                        preset,
                        x,
                        yTop,
                        width,
                        height,
                        document.SlideHeightPoints,
                        stroke,
                        lineWidth,
                        headEnd,
                        tailEnd,
                        presetAdjustmentsOverride);
                }

                if (hasDash)
                {
                    graphics.ClearLineDash();
                }

                if (lineCap is not null)
                {
                    graphics.SetLineCap(0);
                    graphics.SetLineJoin(0);
                }

                if (transparentStroke)
                {
                    graphics.RestoreState();
                }
            }

            if (transformed)
            {
                graphics.RestoreState();
            }

            return;
        }

        bool repeatSlideClipBeforeFill = hasStroke && hasFill && !transformed;
        if (repeatSlideClipBeforeFill)
        {
            ClipSlideBoundsEvenOdd(document, graphics);
        }

        if (gradientFill is not null)
        {
            graphics.SaveState();
            ClipToPresetShape(graphics, shapeProperties, preset, x, y, width, height, presetAdjustmentsOverride);

            DrawLinearGradientFill(graphics, gradientFill, x, y, width, height);

            graphics.RestoreState();
        }
        else if (hasPictureFill && pictureFillName is not null && pictureFillImage is not null)
        {
            CropRect crop = pictureFillCrop;
            FillRect fillRect = pictureFillRect;
            double imageX = x + fillRect.Left * width;
            double imageY = y + fillRect.Bottom * height;
            double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
            double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));
            bool transparentPictureFill = pictureFillAlpha < 0.999d;
            if (transparentPictureFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(pictureFillAlpha, 1d);
            }

            bool clippedToShape = preset != "rect";
            if (clippedToShape)
            {
                graphics.SaveState();
                ClipToPresetShape(graphics, shapeProperties, preset, x, y, width, height, presetAdjustmentsOverride);
            }

            DrawImageFill(graphics, pictureFillName, imageX, imageY, imageWidth, imageHeight, crop);

            if (clippedToShape)
            {
                graphics.RestoreState();
            }

            if (transparentPictureFill)
            {
                graphics.RestoreState();
            }

            images?.Add(new PdfImageResource(pictureFillName, pictureFillImage));
        }
        else if (hasPatternFill)
        {
            bool transparentFill = patternFill.Alpha < 0.999d;
            if (transparentFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(patternFill.Alpha, 1d);
            }

            graphics.SetFillRgb(patternFill.Background.Red, patternFill.Background.Green, patternFill.Background.Blue);
            DrawPresetFill(graphics, shapeProperties, preset, x, y, width, height, presetAdjustmentsOverride);
            StrokeShapePatternFill(graphics, shapeProperties, preset, x, y, width, height, patternFill, presetAdjustmentsOverride);

            if (transparentFill)
            {
                graphics.RestoreState();
            }
        }
        else if (hasFill)
        {
            bool transparentFill = fillAlpha < 0.999d;
            if (transparentFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(fillAlpha, 1d);
            }

            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
            DrawPresetFill(graphics, shapeProperties, preset, x, y, width, height, presetAdjustmentsOverride);

            if (transparentFill)
            {
                graphics.RestoreState();
            }
        }

        if (hasStroke)
        {
            if (hasFill && !repeatSlideClipBeforeFill)
            {
                if (transformed)
                {
                    graphics.RestoreState();
                    ClipSlideBoundsEvenOdd(document, graphics);
                    graphics.SaveState();
                    ApplyShapeTransform(graphics, x, y, width, height, bounds);
                }
                else
                {
                    ClipSlideBoundsEvenOdd(document, graphics);
                }
            }

            bool transparentStroke = strokeAlpha < 0.999d;
            if (transparentStroke)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, strokeAlpha);
            }

            graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
            graphics.SetLineWidth(lineWidth);
            if (hasDash)
            {
                graphics.SetLineDash(dashPattern);
            }

            int? appliedLineJoin = lineJoin ?? (lineCap is null ? null : 1);
            if (lineCap is { } cap)
            {
                graphics.SetLineCap(cap);
            }

            if (appliedLineJoin is { } join)
            {
                graphics.SetLineJoin(join);
            }

            if (preset == "arc")
            {
                bool useFilledLineEndOutline =
                    !hasDash &&
                    lineCap is null &&
                    headEnd.IsNone &&
                    tailEnd.Kind is LineEndKind.Triangle or LineEndKind.Arrow or LineEndKind.Stealth;
                if (useFilledLineEndOutline)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                }

                if (!useFilledLineEndOutline ||
                    !TryFillPresetArcLineEndOutline(graphics, shapeProperties, x, y, width, height, lineWidth, tailEnd, presetAdjustmentsOverride))
                {
                    DrawPresetArcStroke(graphics, shapeProperties, x, y, width, height, presetAdjustmentsOverride);
                }
            }
            else
            {
                StrokePresetShape(graphics, shapeProperties, preset, x, y, width, height, presetAdjustmentsOverride);
            }

            if (hasDash)
            {
                graphics.ClearLineDash();
            }

            if (lineCap is not null)
            {
                graphics.SetLineCap(0);
            }

            if (appliedLineJoin is not null)
            {
                graphics.SetLineJoin(0);
            }

            if (transparentStroke)
            {
                graphics.RestoreState();
            }
        }

        if (transformed)
        {
            graphics.RestoreState();
        }
    }
}
