using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Returns the scene of a metafile that needs vector drawing; null for other image
    // types and for bitmap-only EMF, which the raster pipeline handles. A metafile that
    // cannot be rendered reports IMAGE_UNSUPPORTED_FORMAT once (the failure is cached
    // with the image cache entry the raster pipeline would use) and sets failed.
    private static EmfScene? TryReadVectorEmf(
        PptxSceneImageResource imageResource,
        PptxSceneImageRecolor recolor,
        Dictionary<string, PdfImageXObject?>? imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        EmfFontLookup metafileFonts,
        CancellationToken cancellationToken,
        out bool failed)
    {
        failed = false;
        if (!OoxImageDecoder.IsMetafileContentType(imageResource.ContentType) || !IsNoImageRecolor(recolor))
        {
            return null;
        }

        string cacheKey = imageResource.PartName + "\u001f" + ImageRecolorCacheKey(recolor);
        if (imageCache is not null && imageCache.ContainsKey(cacheKey))
        {
            // Already decoded (or already diagnosed) through the raster pipeline.
            return null;
        }

        try
        {
            EmfScene scene = EmfScene.Read(imageResource.Bytes, cancellationToken, metafileFonts);
            return scene.IsBitmapOnly ? null : scene;
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            imageCache?.TryAdd(cacheKey, null);
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{imageResource.ContentType}' could not be rendered and was ignored: {ex.Message}",
                imageResource.PartName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: imageResource.ContentType,
                Fallback: "Ignored"));
            failed = true;
            return null;
        }
    }

    // The metafile frame fills the picture rectangle; a source crop enlarges the drawn
    // frame so that only the cropped part falls inside the caller's clip.
    private static void DrawEmfPicture(
        PdfGraphicsBuilder graphics,
        EmfScene scene,
        double x,
        double y,
        double width,
        double height,
        CropRect crop,
        List<PdfImageResource> images,
        ref int index,
        string partName,
        string contentType,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        CancellationToken cancellationToken)
    {
        double visibleWidth = Math.Max(0.001d, 1d - crop.Left - crop.Right);
        double visibleHeight = Math.Max(0.001d, 1d - crop.Top - crop.Bottom);
        double fullWidth = width / visibleWidth;
        double fullHeight = height / visibleHeight;
        int nextIndex = index;
        try
        {
            EmfPdfWriter.Draw(graphics, scene, x - crop.Left * fullWidth, y - crop.Bottom * fullHeight, fullWidth, fullHeight,
                image =>
                {
                    string name = "Im" + nextIndex++;
                    images.Add(new PdfImageResource(name, image));
                    return name;
                }, cancellationToken);
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

        index = nextIndex;
    }
}
