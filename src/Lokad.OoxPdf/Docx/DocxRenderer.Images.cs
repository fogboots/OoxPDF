using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static DocxInlineImageLayout PlaceInlineImageForPaint(DocxInlineImageLayout image, DocxMarkupContext context)
    {
        double x = ResolveTextEmissionXOffset(context);
        double y = image.StaticStoryUsesPhysicalY ? 0d : ResolveTextEmissionBaselineOffset(context);
        return x == 0d && y == 0d ? image : image.Translate(x, -y);
    }

    private static DocxInlineImageLayout PlaceCellImageForPaint(DocxInlineImageLayout image,
        double xOffset, double yOffset, DocxTableCellClipSnapshot clip)
    {
        var placed = xOffset == 0d && yOffset == 0d ? image : image.Translate(xOffset, -yOffset);
        return placed with { LineFrame = new DocxPictureLineFrame(clip.Y, clip.Height) };
    }

    private static void RenderImage(DocxInlineImage image, double x, double y, double width, double height,
        PdfGraphicsBuilder graphics, List<PdfImageResource> pageImages, Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageIndex, CancellationToken cancellationToken, Dictionary<string, PdfImageXObject?> imageCache, ref int imageIndex,
        EmfFontLookup metafileFonts, DocxPictureLineFrame? lineFrame = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxPictureCrop crop = image.Crop;
        bool hasFlip = image.FlipHorizontal || image.FlipVertical;
        bool rotates = image.RotationDegrees != 0d;
        double paintX = x, paintY = y, paintWidth = width, paintHeight = height;
        if (rotates)
        {
            bool resolved = TryResolveRotatedPictureFrame(image, x, y, width, height, out PdfRectangle frame);
            if (resolved) { paintX = frame.X; paintY = frame.Y; paintWidth = frame.Width; paintHeight = frame.Height; }
            if (!resolved)
            {
                diagnosticSink?.Invoke(new OoxPdfDiagnostic("DOCX_UNSUPPORTED_IMAGE_TRANSFORM", OoxPdfSeverity.Warning,
                    "The picture rotation has no usable source or line frame; unrotated paint is retained.", image.PartName,
                    PageIndex: pageIndex, Feature: "picture rotation", Fallback: "UnrotatedImage"));
                rotates = false; paintY = y; paintWidth = width; paintHeight = height;
            }
        }
        if (crop.IsEmpty && !hasFlip && !rotates)
        {
            RenderImageCore(image, x, y, width, height, graphics, pageImages, diagnosticSink,
                pageIndex, cancellationToken, imageCache, ref imageIndex, metafileFonts);
            return;
        }
        double sourceWidth = paintWidth / (1d - crop.Left - crop.Right);
        double sourceHeight = paintHeight / (1d - crop.Top - crop.Bottom);
        double sourceX = paintX - crop.Left * sourceWidth;
        double sourceY = paintY - crop.Bottom * sourceHeight;
        if (!crop.IsUsable || !double.IsFinite(sourceWidth) || !double.IsFinite(sourceHeight) ||
            !double.IsFinite(sourceX) || !double.IsFinite(sourceY))
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic("DOCX_UNSUPPORTED_IMAGE_CROP", OoxPdfSeverity.Warning,
                "The picture source rectangle is unusable; the uncropped image is retained.", image.PartName,
                PageIndex: pageIndex, Feature: "image crop", Fallback: "UncroppedImage"));
            sourceX = paintX; sourceY = paintY; sourceWidth = paintWidth; sourceHeight = paintHeight;
            if (!hasFlip && !rotates)
            {
                RenderImageCore(image, x, y, width, height, graphics, pageImages, diagnosticSink,
                    pageIndex, cancellationToken, imageCache, ref imageIndex, metafileFonts);
                return;
            }
        }
        graphics.SaveState();
        try
        {
            double radians = rotates ? -(image.RotationDegrees % 360d) * Math.PI / 180d : 0d;
            double cos = Math.Cos(radians), sin = Math.Sin(radians);
            double centerX = paintX + paintWidth / 2d, centerY = paintY + paintHeight / 2d;
            if (!rotates) graphics.ClipRectangle(paintX, paintY, paintWidth, paintHeight);
            if (rotates && lineFrame is { IsUsable: true } line)
            {
                // Paragraph clipping stays in page space; the picture viewport rotates.
                double rotatedWidth = Math.Abs(cos) * paintWidth + Math.Abs(sin) * paintHeight;
                graphics.ClipRectangle(centerX - rotatedWidth / 2d, line.Bottom, rotatedWidth, line.Height);
            }
            if (hasFlip || rotates)
            {
                double sx = image.FlipHorizontal ? -1d : 1d, sy = image.FlipVertical ? -1d : 1d;
                double a = cos * sx, b = rotates ? sin * sx : 0d, c = rotates ? -sin * sy : 0d, d = cos * sy;
                graphics.Transform(a, b, c, d, centerX - a * centerX - c * centerY, centerY - b * centerX - d * centerY);
            }
            if (rotates) graphics.ClipRectangle(paintX, paintY, paintWidth, paintHeight);
            RenderImageCore(image, sourceX, sourceY, sourceWidth, sourceHeight, graphics, pageImages,
                diagnosticSink, pageIndex, cancellationToken, imageCache, ref imageIndex, metafileFonts);
        }
        finally { graphics.RestoreState(); }
    }

    private static bool TryResolveRotatedPictureFrame(DocxInlineImage image, double x, double y, double width, double height,
        out PdfRectangle frame)
    {
        frame = default;
        // Without a paragraph line frame the picture still rotates in its own box;
        // only the line clip is skipped.
        if (image.RotationDegrees == 0d || !double.IsFinite(image.RotationDegrees) ||
            image.PictureFrame is not { IsUsable: true } source) return false;
        double scale = Math.Min(width / source.Width, height / source.Height);
        double paintWidth = source.Width * scale, paintHeight = source.Height * scale;
        double paintY = y + height - paintHeight;
        if (!double.IsFinite(x) || !double.IsFinite(paintY) || !double.IsFinite(paintWidth) ||
            !double.IsFinite(paintHeight) || paintWidth <= 0d || paintHeight <= 0d) return false;
        frame = new PdfRectangle(x, paintY, paintWidth, paintHeight);
        return true;
    }

    private static void RenderImageCore(DocxInlineImage image, double x, double y, double width, double height,
        PdfGraphicsBuilder graphics, List<PdfImageResource> pageImages, Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageIndex, CancellationToken cancellationToken, Dictionary<string, PdfImageXObject?> imageCache, ref int imageIndex,
        EmfFontLookup metafileFonts)
    {
        if (image.Bytes is { } bytes && image.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
        {
            Action<OoxPdfDiagnostic>? svgSink = diagnosticSink is null ? null : diagnostic =>
                diagnosticSink(image.RasterFallback is not null && diagnostic.Severity == OoxPdfSeverity.Error && diagnostic.Fallback == "Ignored"
                    ? diagnostic with { Severity = OoxPdfSeverity.Warning, Fallback = "RasterImage", Message = diagnostic.Message + " The authored raster preview is used." }
                    : diagnostic);
            if (PptxRenderer.TryRenderSvgImage(graphics, x, y, width, height, bytes, svgSink, pageIndex, image.PartName, cancellationToken)) return;
            if (image.RasterFallback is not { } preview) return;
            image = preview;
        }
        if (TryDrawVectorEmf(image, x, y, width, height, graphics, pageImages, diagnosticSink, pageIndex, metafileFonts, cancellationToken, imageCache, ref imageIndex)) return;
        PdfImageXObject? resource = CreateImage(image, imageCache, diagnosticSink, pageIndex, cancellationToken);
        if (resource is null) return;
        string name = "Im" + imageIndex++;
        graphics.DrawImage(name, x, y, width, height);
        pageImages.Add(new PdfImageResource(name, resource));
    }

    // Metafiles that are more than one picture-sized bitmap are drawn as vectors;
    // bitmap-only EMF decodes to pixels through CreateImage. A failure is cached like
    // any other undecodable image so a repeated metafile is diagnosed once.
    private static bool TryDrawVectorEmf(DocxInlineImage image, double x, double y, double width, double height,
        PdfGraphicsBuilder graphics, List<PdfImageResource> pageImages, Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageIndex, EmfFontLookup metafileFonts, CancellationToken cancellationToken, Dictionary<string, PdfImageXObject?> imageCache, ref int imageIndex)
    {
        if (image.Bytes is not { } bytes || !OoxImageDecoder.IsMetafileContentType(image.ContentType)) return false;
        string? cacheKey = GetImageCacheKey(image);
        if (cacheKey is not null && imageCache.ContainsKey(cacheKey)) return false;
        int nextIndex = imageIndex;
        try
        {
            EmfScene scene = EmfScene.Read(bytes, cancellationToken, metafileFonts);
            if (scene.IsBitmapOnly) return false;
            EmfPdfWriter.Draw(graphics, scene, x, y, width, height, resource =>
            {
                string name = "Im" + nextIndex++;
                pageImages.Add(new PdfImageResource(name, resource));
                return name;
            }, cancellationToken);
            imageIndex = nextIndex;
            return true;
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            if (cacheKey is not null) imageCache.TryAdd(cacheKey, null);
            EmitImageDiagnostic(diagnosticSink, image, pageIndex, ex.Message);
            return true;
        }
    }
}
