using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void RenderPicture(
        PptxSceneNode picture,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        GroupTransform transform,
        List<PdfImageResource> images,
        ref int index,
        string? sourcePartName)
    {
        if (picture.Picture is null || picture.Bounds is null)
        {
            return;
        }

        if (picture.Picture.TargetPartName is null)
        {
            context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_RELATIONSHIP", OoxPdfSeverity.Error,
                "Picture image relationship could not be resolved and the image was ignored.",
                sourcePartName ?? context.SlidePartName, SlideIndex: context.SlideNumber,
                Feature: "image", Fallback: "Ignored"));
            return;
        }

        RenderPicture(
            context.Document,
            graphics,
            context.DiagnosticSink,
            context.SlideNumber,
            transform,
            images,
            context.ImageCache,
            ref index,
            picture.Picture.TargetPartName,
            picture.Picture.Resource,
            ToShapeBounds(picture.Bounds),
            ToCropRect(picture.Picture.Crop),
            ToFillRect(picture.Picture.Fill),
            picture.Picture.Alpha,
            picture.Picture.Recolor,
            ToLineStyle(picture.Picture.Line),
            ToOuterShadow(picture.Picture.OuterShadow),
            picture.Picture,
            context.FontResolver.MetafileFonts(context.CancellationToken),
            context.CancellationToken);
    }

    private static void RenderPicture(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        GroupTransform transform,
        List<PdfImageResource> images,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int index,
        string targetPartName,
        PptxSceneImageResource? imageResource,
        ShapeBounds rawBounds,
        CropRect crop,
        FillRect fillRect,
        double alpha,
        PptxSceneImageRecolor recolor,
        LineStyle line,
        OuterShadow? outerShadow,
        PptxScenePicture geometry,
        EmfFontLookup metafileFonts,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (imageResource is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_PART",
                OoxPdfSeverity.Error,
                "Referenced image part was missing and the image was ignored.",
                targetPartName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "image",
                Fallback: "Ignored"));
            return;
        }

        string preset = geometry.Preset;
        PptxSceneCustomGeometry? customGeometry = geometry.HasCustomGeometry && geometry.CustomGeometry is { HasGeometry: true, HasUnsupportedGeometry: false } authored
            ? authored
            : null;
        bool supportedGeometry = customGeometry is not null || (!geometry.HasCustomGeometry && CanRenderPictureFillPreset(preset));
        if (!supportedGeometry)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic("PPTX_PICTURE_GEOMETRY_UNSUPPORTED", OoxPdfSeverity.Warning,
                "Picture shape geometry is unsupported; its rectangular frame is retained.", targetPartName,
                SlideIndex: slideIndex, Feature: "picture geometry", Fallback: "Rectangular frame"));
            preset = "rect";
        }
        ShapeBounds transformedBounds = transform.Apply(rawBounds);
        double x = OoxUnits.EmuToPoints(transformedBounds.X);
        double yTop = OoxUnits.EmuToPoints(transformedBounds.Y);
        double width = OoxUnits.EmuToPoints(transformedBounds.Width);
        double height = OoxUnits.EmuToPoints(transformedBounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        bool hasTransform = Math.Abs(transformedBounds.RotationDegrees) > 0.001d || transformedBounds.FlipHorizontal || transformedBounds.FlipVertical;
        RenderPictureOuterShadow(document, graphics, transformedBounds, x, y, width, height, hasTransform, outerShadow, images, ref index, preset, geometry.PresetAdjustments, cancellationToken);
        if (imageResource.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
        {
            RenderSvgPicture(graphics, document, transformedBounds, imageResource.Bytes, crop, fillRect, diagnosticSink, slideIndex, targetPartName, cancellationToken, preset, geometry.PresetAdjustments);
            StrokePictureFrame(document, graphics, transformedBounds, x, y, width, height, line, hasTransform, preset, geometry.PresetAdjustments, customGeometry);
            return;
        }

        // Metafiles that are more than one picture-sized bitmap are drawn as vectors;
        // bitmap-only EMF shares the raster pipeline below.
        EmfScene? emf = TryReadVectorEmf(imageResource, recolor, imageCache, diagnosticSink, slideIndex, metafileFonts, cancellationToken, out bool emfFailed);
        if (emfFailed)
        {
            return;
        }

        PdfImageXObject? image = null;
        if (emf is null && !crop.IsEmpty)
        {
            image = GetOrCreateCroppedImage(crop);
            if (image is not null)
            {
                crop = default;
            }
        }

        PdfImageXObject? GetOrCreateCroppedImage(CropRect crop)
        {
            string cacheKey = imageResource.PartName + "\u001f" + ImageRecolorCacheKey(recolor) + "\u001fcrop:" +
                crop.Left.ToString("R", CultureInfo.InvariantCulture) + "," +
                crop.Top.ToString("R", CultureInfo.InvariantCulture) + "," +
                crop.Right.ToString("R", CultureInfo.InvariantCulture) + "," +
                crop.Bottom.ToString("R", CultureInfo.InvariantCulture);
            if (imageCache is not null && imageCache.TryGetValue(cacheKey, out PdfImageXObject? cached))
            {
                return cached;
            }

            PdfImageXObject? croppedImage = CreateCroppedImage(
                imageResource.PartName,
                imageResource.ContentType,
                imageResource.Bytes);
            imageCache?.TryAdd(cacheKey, croppedImage);
            return croppedImage;

            PdfImageXObject? CreateCroppedImage(string partName, string contentType, byte[] bytes)
            {
                try
                {
                    // D02: shared pixel decoding; unsupported types keep the silent
                    // PDF-clipping fallback (no diagnostic), exactly as before.
                    if (!OoxImageDecoder.IsSupportedContentType(contentType))
                    {
                        return null;
                    }

                    // A JPEG the pixel decoder does not take (progressive, CMYK, or past the
                    // pixel budget) is placed unchanged by DCTDecode and clipped to the crop
                    // window, which shows the same pixels, so nothing is lost.
                    if (OoxImageDecoder.IsJpegContentType(contentType) && !OoxImageDecoder.CanDecodeJpegPixels(bytes))
                    {
                        return null;
                    }

                    using DecodedPixels decoded = OoxImageDecoder.DecodePixelsOwned(contentType, bytes, cancellationToken);
                    return CreateCroppedRgbImage(decoded.Width, decoded.Height, decoded.Rgb, decoded.Alpha, recolor, crop, cancellationToken);
                }
                catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or IndexOutOfRangeException)
                {
                    diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                        "IMAGE_CROP_UNSUPPORTED_FORMAT",
                        OoxPdfSeverity.Warning,
                        $"Image '{partName}' could not be decoded for Office-style cropped image embedding on slide {slideIndex}; falling back to PDF clipping.",
                        partName,
                        PageIndex: null,
                        SlideIndex: slideIndex,
                        Feature: "image crop",
                        Fallback: "PDF clipping"));
                }

                return null;
            }
        }

        if (emf is null)
        {
            image ??= GetOrCreateImage(imageResource, recolor, imageCache, diagnosticSink, slideIndex, cancellationToken);
            if (image is null)
            {
                return;
            }
        }

        string name = emf is null ? "Im" + index++ : string.Empty;
        double imageX = x + fillRect.Left * width;
        double imageY = y + fillRect.Bottom * height;
        double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
        double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));
        bool transparent = alpha < 0.999d;
        if (hasTransform)
        {
            graphics.SaveState();
            ClipSlideBoundsEvenOdd(document, graphics);
            ApplyShapeTransform(graphics, x, y, width, height, transformedBounds);
        }

        if (transparent)
        {
            graphics.SaveState();
            graphics.SetAlpha(alpha, 1d);
        }

        graphics.SaveState();
        double clipX = imageX;
        double clipY = imageY;
        double clipWidth = imageWidth;
        double clipHeight = imageHeight;
        if (!hasTransform &&
            !TryIntersectWithSlideBounds(imageX, imageY, imageWidth, imageHeight, document, out clipX, out clipY, out clipWidth, out clipHeight))
        {
            graphics.RestoreState();
            if (transparent)
            {
                graphics.RestoreState();
            }

            return;
        }

        // Source crop/fill controls image sampling. The authored mask belongs
        // to the picture frame, in the same transformed coordinate system.
        if (customGeometry is not null)
            ClipToCustomGeometry(graphics, customGeometry, x, y, width, height);
        else if (preset != "rect")
            ClipToPresetShape(graphics, null, preset, x, y, width, height, geometry.PresetAdjustments);
        graphics.ClipRectangleEvenOdd(clipX, clipY, clipWidth, clipHeight);
        if (emf is not null)
        {
            DrawEmfPicture(graphics, emf, imageX, imageY, imageWidth, imageHeight, crop, images, ref index, imageResource.PartName, imageResource.ContentType, diagnosticSink, slideIndex, cancellationToken);
        }
        else if (crop.IsEmpty)
        {
            graphics.DrawImage(name, imageX, imageY, imageWidth, imageHeight);
        }
        else
        {
            graphics.DrawImageCropped(name, imageX, imageY, imageWidth, imageHeight, crop.Left, crop.Top, crop.Right, crop.Bottom);
        }

        graphics.RestoreState();

        if (transparent)
        {
            graphics.RestoreState();
        }

        if (hasTransform)
        {
            graphics.RestoreState();
        }

        if (image is not null)
        {
            images.Add(new PdfImageResource(name, image));
        }

        StrokePictureFrame(document, graphics, transformedBounds, x, y, width, height, line, hasTransform, preset, geometry.PresetAdjustments, customGeometry);
    }

    private static void RenderPictureOuterShadow(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        ShapeBounds bounds,
        double x,
        double y,
        double width,
        double height,
        bool transformed,
        OuterShadow? outerShadow,
        List<PdfImageResource> images,
        ref int imageIndex,
        string preset,
        IReadOnlyDictionary<string, double> adjustments,
        CancellationToken cancellationToken)
    {
        if (outerShadow is not { } shadow)
        {
            return;
        }

        graphics.SaveState();
        ClipSlideBoundsEvenOdd(document, graphics);
        if (transformed)
        {
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }

        if (shadow.BlurRadius > 0d)
        {
            double radius = preset == "roundRect" ? ReadRoundRectangleRadius(null, adjustments, width, height) : 0d;
            DrawRasterOuterShadow(graphics, x, y, width, height, shadow, images, ref imageIndex,
                preset, radius, cancellationToken);
        }
        else
        {
            DrawOuterShadow(graphics, null, preset, x, y, width, height, shadow, adjustments);
        }

        graphics.RestoreState();
    }

    private static void StrokePictureFrame(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        ShapeBounds bounds,
        double x,
        double y,
        double width,
        double height,
        LineStyle line,
        bool transformed,
        string preset,
        IReadOnlyDictionary<string, double> adjustments,
        PptxSceneCustomGeometry? customGeometry = null)
    {
        if (!line.HasLine)
        {
            return;
        }

        if (transformed)
        {
            graphics.SaveState();
            ClipSlideBoundsEvenOdd(document, graphics);
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }
        else
        {
            ClipSlideBoundsEvenOdd(document, graphics);
        }

        bool transparentStroke = line.Alpha < 0.999d;
        if (transparentStroke)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, line.Alpha);
        }

        graphics.SetStrokeRgb(line.Color.Red, line.Color.Green, line.Color.Blue);
        graphics.SetLineWidth(line.Width);
        if (line.DashPattern is { Count: > 0 })
        {
            graphics.SetLineDash(line.DashPattern);
        }

        if (line.Cap is { } cap)
        {
            graphics.SetLineCap(cap);
        }

        if (line.Join is { } join)
        {
            graphics.SetLineJoin(join);
        }

        double outlineOutset = line.Width / 2d;
        // Office's picture outline surrounds the frame: its centerline grows
        // by half the authored width on each side, including nonrectangular masks.
        if (customGeometry is not null)
        {
            foreach (PptxSceneCustomPath path in customGeometry.Paths.Where(path => path.AllowsStroke))
            {
                AppendCustomGeometryPath(graphics, customGeometry, path, x - outlineOutset, y - outlineOutset, width + line.Width, height + line.Width);
                graphics.StrokeCurrentPath();
            }
        }
        else
        {
            StrokePresetShape(graphics, null, preset, x - outlineOutset, y - outlineOutset,
                width + line.Width, height + line.Width, adjustments);
        }

        if (line.DashPattern is { Count: > 0 })
        {
            graphics.ClearLineDash();
        }

        if (line.Cap is not null)
        {
            graphics.SetLineCap(0);
        }

        if (line.Join is not null)
        {
            graphics.SetLineJoin(0);
        }

        if (transparentStroke)
        {
            graphics.RestoreState();
        }

        if (transformed)
        {
            graphics.RestoreState();
        }
    }

    private static ShapeBounds ToShapeBounds(PptxSceneBounds bounds)
    {
        return new ShapeBounds(
            bounds.XEmu,
            bounds.YEmu,
            bounds.WidthEmu,
            bounds.HeightEmu,
            bounds.RotationDegrees,
            bounds.FlipHorizontal,
            bounds.FlipVertical);
    }

    private static CropRect ToCropRect(PptxSceneRect rect)
    {
        return new CropRect(rect.Left, rect.Top, rect.Right, rect.Bottom);
    }

    private static FillRect ToFillRect(PptxSceneRect rect)
    {
        return new FillRect(rect.Left, rect.Top, rect.Right, rect.Bottom);
    }

    private static PdfImageXObject? GetOrCreateImage(
        PptxSceneImageResource imageResource,
        PptxSceneImageRecolor recolor,
        Dictionary<string, PdfImageXObject?>? imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        CancellationToken cancellationToken = default)
    {
        return GetOrCreateImage(imageResource.PartName, imageResource.ContentType, imageResource.Bytes, recolor, imageCache, diagnosticSink, slideIndex, cancellationToken);
    }

    private static PdfImageXObject? GetOrCreateImage(
        string partName,
        string contentType,
        byte[] bytes,
        PptxSceneImageRecolor recolor,
        Dictionary<string, PdfImageXObject?>? imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        CancellationToken cancellationToken = default)
    {
        string cacheKey = partName + "\u001f" + ImageRecolorCacheKey(recolor);
        if (imageCache is not null && imageCache.TryGetValue(cacheKey, out PdfImageXObject? cached))
        {
            return cached;
        }

        PdfImageXObject? image = CreateImage(partName, contentType, bytes, recolor, diagnosticSink, slideIndex, cancellationToken);
        imageCache?.TryAdd(cacheKey, image);
        return image;
    }

    private static PdfImageXObject? CreateImage(
        string partName,
        string contentType,
        byte[] bytes,
        PptxSceneImageRecolor recolor,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        CancellationToken cancellationToken = default)
    {
        try
        {
            // D02: without recolor the dispatch is the shared decoder; the recolor
            // fallback policy below stays PPTX-specific.
            if (IsNoImageRecolor(recolor))
            {
                return OoxImageDecoder.Decode(contentType, bytes, static rgb => rgb, cancellationToken);
            }

            if (OoxImageDecoder.IsJpegContentType(contentType))
            {
                JpegInfo info = JpegInfo.Read(bytes);
                if (info.IsBaselineDct && info.BitsPerComponent == 8 && info.ComponentCount is 1 or 3)
                {
                    try
                    {
                        using DecodedPixels recolorSource = OoxImageDecoder.DecodePixelsOwned(contentType, bytes, cancellationToken);
                        byte[]? jpegAlpha = recolorSource.Alpha;
                        byte[] rgb = ApplyImageRecolor(recolorSource.Rgb, ref jpegAlpha, recolor, cancellationToken);
                        return PdfImageXObject.RgbPng(recolorSource.Width, recolorSource.Height, rgb, jpegAlpha);
                    }
                    catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or IndexOutOfRangeException)
                    {
                    }
                }

                diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                    "PPTX_UNSUPPORTED_IMAGE_RECOLOR",
                    OoxPdfSeverity.Warning,
                    $"PPTX {ImageRecolorKindName()} image recolor could not be applied to {contentType} {info.FrameProfileName} image data and was ignored.",
                    partName,
                    PageIndex: null,
                    SlideIndex: slideIndex,
                    Feature: "image recolor",
                    Fallback: "Original image"));

                return PdfImageXObject.Jpeg(info.Width, info.Height, bytes, info.ComponentCount, info.BitsPerComponent, info.InvertedCmyk);
            }

            // PNG/BMP share the common pixel decoder with the recolor map applied;
            // unknown types throw with the same message the explicit branch used to
            // emit, so the catch below reports byte-identical diagnostics.
            using DecodedPixels decoded = OoxImageDecoder.DecodePixelsOwned(contentType, bytes, cancellationToken);
            byte[]? decodedAlpha = decoded.Alpha;
            byte[] recoloredRgb = ApplyImageRecolor(decoded.Rgb, ref decodedAlpha, recolor, cancellationToken);
            return PdfImageXObject.RgbPng(decoded.Width, decoded.Height, recoloredRgb, decodedAlpha);
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{contentType}' could not be rendered and was ignored: {ex.Message}",
                partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: contentType,
                Fallback: "Ignored"));
        }

        return null;

        string ImageRecolorKindName()
        {
            return recolor.Kind switch
            {
                PptxSceneImageRecolorKind.Luminance => "luminance",
                PptxSceneImageRecolorKind.Duotone => "duotone",
                PptxSceneImageRecolorKind.Grayscale => "grayscale",
                PptxSceneImageRecolorKind.BiLevel => "bi-level",
                PptxSceneImageRecolorKind.ColorChange => "color change",
                _ => "none"
            };
        }
    }

    private static PdfImageXObject? CreateCroppedRgbImage(
        int width,
        int height,
        byte[] rgb,
        byte[]? alpha,
        PptxSceneImageRecolor recolor,
        CropRect crop,
        CancellationToken cancellationToken)
    {
        int left = (int)Math.Floor(width * Math.Clamp(crop.Left, 0d, 1d));
        int top = (int)Math.Floor(height * Math.Clamp(crop.Top, 0d, 1d));
        int right = (int)Math.Ceiling(width * (1d - Math.Clamp(crop.Right, 0d, 1d)));
        int bottom = (int)Math.Ceiling(height * (1d - Math.Clamp(crop.Bottom, 0d, 1d)));

        left = Math.Clamp(left, 0, width);
        top = Math.Clamp(top, 0, height);
        right = Math.Clamp(right, left + 1, width);
        bottom = Math.Clamp(bottom, top + 1, height);

        int croppedWidth = right - left;
        int croppedHeight = bottom - top;
        if (croppedWidth <= 0 || croppedHeight <= 0)
        {
            return null;
        }

        byte[] croppedRgb = new byte[croppedWidth * croppedHeight * 3];
        byte[]? croppedAlpha = alpha is null ? null : new byte[croppedWidth * croppedHeight];
        for (int y = 0; y < croppedHeight; y++)
        {
            if ((y & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            int sourceY = top + y;
            Buffer.BlockCopy(rgb, (sourceY * width + left) * 3, croppedRgb, y * croppedWidth * 3, croppedWidth * 3);
            if (alpha is not null && croppedAlpha is not null)
            {
                Buffer.BlockCopy(alpha, sourceY * width + left, croppedAlpha, y * croppedWidth, croppedWidth);
            }
        }
        if (IsNoImageRecolor(recolor))
        {
            return PdfImageXObject.RgbPng(croppedWidth, croppedHeight, croppedRgb, croppedAlpha);
        }
        byte[] recoloredRgb = ApplyImageRecolor(croppedRgb, ref croppedAlpha, recolor, cancellationToken);
        return PdfImageXObject.RgbPng(croppedWidth, croppedHeight, recoloredRgb, croppedAlpha);
    }

    private static byte[] ApplyImageRecolor(byte[] rgb, ref byte[]? alpha, PptxSceneImageRecolor recolor, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (IsNoImageRecolor(recolor))
        {
            return rgb;
        }

        if (recolor.Kind == PptxSceneImageRecolorKind.ColorChange)
        {
            return ApplyImageColorChange(rgb, ref alpha, recolor, cancellationToken);
        }

        byte[] transformed = new byte[rgb.Length];
        for (int i = 0; i < rgb.Length; i += 3)
        {
            if (i % 12288 == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            double red = rgb[i];
            double green = rgb[i + 1];
            double blue = rgb[i + 2];
            if (recolor.Kind == PptxSceneImageRecolorKind.Luminance)
            {
                transformed[i] = ApplyBrightnessContrast(red);
                transformed[i + 1] = ApplyBrightnessContrast(green);
                transformed[i + 2] = ApplyBrightnessContrast(blue);
                continue;
            }

            double luma = (0.2126d * red + 0.7152d * green + 0.0722d * blue) / 255d;
            if (recolor.Kind == PptxSceneImageRecolorKind.Grayscale)
            {
                byte gray = PptxColorResolver.ToByte(luma * 255d);
                transformed[i] = gray;
                transformed[i + 1] = gray;
                transformed[i + 2] = gray;
                continue;
            }

            if (recolor.Kind == PptxSceneImageRecolorKind.BiLevel)
            {
                byte value = luma >= recolor.Threshold ? (byte)255 : (byte)0;
                transformed[i] = value;
                transformed[i + 1] = value;
                transformed[i + 2] = value;
                continue;
            }

            transformed[i] = Interpolate(recolor.Dark.Red, recolor.Light.Red, luma);
            transformed[i + 1] = Interpolate(recolor.Dark.Green, recolor.Light.Green, luma);
            transformed[i + 2] = Interpolate(recolor.Dark.Blue, recolor.Light.Blue, luma);
        }

        return transformed;

        byte ApplyBrightnessContrast(double channel)
        {
            double value = channel / 255d;
            value = ApplyContrast(value, recolor.Contrast);
            value = ApplyBrightness(value, recolor.Brightness);
            return PptxColorResolver.ToByte(value * 255d);

            double ApplyContrast(double value, double contrast)
            {
                value = Math.Clamp(value, 0d, 1d);
                if (contrast < 0d)
                {
                    return value * (1d + contrast);
                }

                if (contrast > 0d)
                {
                    double scale = 1d - contrast;
                    return value < 0.5d
                        ? value * scale
                        : 1d - (1d - value) * scale;
                }

                return value;
            }

            double ApplyBrightness(double value, double brightness)
            {
                value = Math.Clamp(value, 0d, 1d);
                return brightness < 0d
                    ? value * (1d + brightness)
                    : value + brightness;
            }
        }
    }

    // Office replaces pixels near the source colour (the same per-channel tolerance
    // LibreOffice applies) with the replacement colour at the replacement alpha.
    private static byte[] ApplyImageColorChange(byte[] rgb, ref byte[]? alpha, PptxSceneImageRecolor recolor, CancellationToken cancellationToken)
    {
        const int Tolerance = 9;
        byte[] changed = (byte[])rgb.Clone();
        byte[]? changedAlpha = alpha is null ? null : (byte[])alpha.Clone();
        double replacementAlpha = recolor.Threshold;
        for (int i = 0; i < rgb.Length; i += 3)
        {
            if (i % 12288 == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            if (Math.Abs(rgb[i] - recolor.Dark.Red) > Tolerance ||
                Math.Abs(rgb[i + 1] - recolor.Dark.Green) > Tolerance ||
                Math.Abs(rgb[i + 2] - recolor.Dark.Blue) > Tolerance)
            {
                continue;
            }

            changed[i] = recolor.Light.Red;
            changed[i + 1] = recolor.Light.Green;
            changed[i + 2] = recolor.Light.Blue;
            if (replacementAlpha >= 1d)
            {
                continue;
            }

            changedAlpha ??= Enumerable.Repeat((byte)255, rgb.Length / 3).ToArray();
            changedAlpha[i / 3] = (byte)Math.Round(changedAlpha[i / 3] * replacementAlpha);
        }

        alpha = changedAlpha;
        return changed;
    }

    private static bool IsNoImageRecolor(PptxSceneImageRecolor recolor)
    {
        return recolor.Kind == PptxSceneImageRecolorKind.None;
    }

    private static string ImageRecolorCacheKey(PptxSceneImageRecolor recolor)
    {
        return recolor.Kind switch
        {
            PptxSceneImageRecolorKind.Luminance => FormattableString.Invariant($"lum:{recolor.Brightness:0.#####}:{recolor.Contrast:0.#####}"),
            PptxSceneImageRecolorKind.Duotone => FormattableString.Invariant($"duo:{recolor.Dark.Red:X2}{recolor.Dark.Green:X2}{recolor.Dark.Blue:X2}:{recolor.Light.Red:X2}{recolor.Light.Green:X2}{recolor.Light.Blue:X2}"),
            PptxSceneImageRecolorKind.Grayscale => "gray",
            PptxSceneImageRecolorKind.BiLevel => FormattableString.Invariant($"bi:{recolor.Threshold:0.#####}"),
            PptxSceneImageRecolorKind.ColorChange => FormattableString.Invariant($"clr:{recolor.Dark.Red:X2}{recolor.Dark.Green:X2}{recolor.Dark.Blue:X2}:{recolor.Light.Red:X2}{recolor.Light.Green:X2}{recolor.Light.Blue:X2}:{recolor.Threshold:0.#####}"),
            _ => "none"
        };
    }

    private static byte Interpolate(byte from, byte to, double ratio)
    {
        return PptxColorResolver.ToByte(from + (to - from) * Math.Clamp(ratio, 0d, 1d));
    }

    private static bool TryReadShapePictureFill(
        XElement shapeProperties,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        List<PdfImageResource>? images,
        Dictionary<string, PdfImageXObject?>? imageCache,
        ref int imageIndex,
        ShapePictureFill? pictureFillOverride,
        out string? name,
        out PdfImageXObject? image,
        out CropRect crop,
        out FillRect fillRect,
        out double alpha)
    {
        name = null;
        image = null;
        crop = default;
        fillRect = default;
        alpha = 1d;
        if (images is null || pictureFillOverride is null || !CanRenderPictureFillPreset(ReadPreset(shapeProperties)))
        {
            return false;
        }

        ShapePictureFill resolvedPictureFill = pictureFillOverride.Value;
        if (resolvedPictureFill.Resource is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_PART",
                OoxPdfSeverity.Error,
                "Referenced image part was missing and the image was ignored.",
                resolvedPictureFill.TargetPartName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "image",
                Fallback: "Ignored"));
            return false;
        }

        image = GetOrCreateImage(resolvedPictureFill.Resource, PptxSceneImageRecolor.None, imageCache, diagnosticSink, slideIndex);
        if (image is null)
        {
            return false;
        }

        crop = resolvedPictureFill.Crop;
        fillRect = resolvedPictureFill.Fill;
        alpha = resolvedPictureFill.Alpha;
        name = "Im" + imageIndex++;
        return true;
    }

    private static bool CanRenderPictureFillPreset(string preset)
    {
        return preset is "rect" or "ellipse" or "roundRect" ||
            TryCreatePresetPolygonPoints(preset, 0d, 0d, 1d, 1d, out _) ||
            HasPresetGeometryDefinition(preset);
    }

    private static void StrokePresetShape(PdfGraphicsBuilder graphics, XElement? shapeProperties, string preset,
        double x, double y, double width, double height, IReadOnlyDictionary<string, double>? adjustments)
    {
        if (preset == "ellipse") graphics.StrokeEllipse(x, y, width, height);
        else if (preset == "roundRect")
            graphics.StrokeRoundedRectangle(x, y, width, height, ReadRoundRectangleRadius(shapeProperties, adjustments, width, height));
        else if (TryCreatePresetPolygonPoints(preset, x, y, width, height, out (double X, double Y)[] points))
            graphics.StrokePolygon(points);
        else if (TryAppendPresetBracketPath(graphics, shapeProperties, preset, x, y, width, height, adjustments))
            graphics.StrokeCurrentPath();
        else if (!TryStrokePresetGeometry(graphics, shapeProperties, preset, x, y, width, height, adjustments))
            graphics.StrokeRectangle(x, y, width, height);
    }

    private static void ClipToCustomGeometry(PdfGraphicsBuilder graphics, PptxSceneCustomGeometry geometry, double x, double y, double width, double height)
    {
        foreach (PptxSceneCustomPath path in geometry.Paths.Where(path => path.AllowsFill))
        {
            AppendCustomGeometryPath(graphics, geometry, path, x, y, width, height);
        }

        graphics.ClipCurrentPath();
    }

    private static void ClipToPresetShape(
        PdfGraphicsBuilder graphics,
        XElement? shapeProperties,
        string preset,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride)
    {
        if (preset == "ellipse")
        {
            graphics.ClipEllipse(x, y, width, height);
        }
        else if (preset == "roundRect")
        {
            graphics.ClipRoundedRectangle(x, y, width, height, ReadRoundRectangleRadius(shapeProperties, presetAdjustmentsOverride, width, height));
        }
        else if (TryCreatePresetPolygonPoints(preset, x, y, width, height, out (double X, double Y)[] polygonPoints))
        {
            graphics.ClipPolygon(polygonPoints);
        }
        else if (!TryClipToPresetGeometry(graphics, shapeProperties, preset, x, y, width, height, presetAdjustmentsOverride))
        {
            graphics.ClipRectangle(x, y, width, height);
        }
    }

    private static void DrawImageFill(PdfGraphicsBuilder graphics, string imageName, double x, double y, double width, double height, CropRect crop)
    {
        if (crop.IsEmpty)
        {
            graphics.DrawImage(imageName, x, y, width, height);
        }
        else
        {
            graphics.DrawImageCropped(imageName, x, y, width, height, crop.Left, crop.Top, crop.Right, crop.Bottom);
        }
    }
}
