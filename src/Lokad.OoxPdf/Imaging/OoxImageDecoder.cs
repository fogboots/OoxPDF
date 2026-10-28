using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Imaging;

/// <summary>
/// Shared Office content-image decoding (D02). JPEG passes through as DCTDecode;
/// PNG, GIF, BMP, TIFF and bitmap-only EMF decode to RGB(A) pixels. A caller-supplied recolor map covers the
/// PPTX recolor case; DOCX passes the identity. Per-format diagnostics (with
/// page/slide location and severity) and PPTX crop/recolor policy stay in the
/// renderers. Throws <see cref="InvalidDataException"/> for corrupt data and
/// <see cref="NotSupportedException"/> for unknown content types.
/// </summary>
internal static class OoxImageDecoder
{
    public static bool IsTiffContentType(string contentType) =>
        contentType.Equals("image/tiff", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/x-tiff", StringComparison.OrdinalIgnoreCase);

    public static bool IsJpegContentType(string contentType) =>
        contentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase);

    public static bool IsPngContentType(string contentType) =>
        contentType.Equals("image/png", StringComparison.OrdinalIgnoreCase);

    public static bool IsBmpContentType(string contentType) =>
        contentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase);

    public static bool IsGifContentType(string contentType) =>
        contentType.Equals("image/gif", StringComparison.OrdinalIgnoreCase);

    public static bool IsEmfContentType(string contentType) =>
        contentType.Equals("image/x-emf", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/emf", StringComparison.OrdinalIgnoreCase);

    public static bool IsSupportedContentType(string contentType) =>
        IsJpegContentType(contentType) || IsPngContentType(contentType) || IsBmpContentType(contentType) || IsTiffContentType(contentType) ||
        IsGifContentType(contentType) || IsEmfContentType(contentType);

    public static PdfImageXObject Decode(string contentType, byte[] bytes, Func<byte[], byte[]> recolorRgb, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (IsJpegContentType(contentType))
        {
            JpegInfo info = JpegInfo.Read(bytes);
            return PdfImageXObject.Jpeg(info.Width, info.Height, bytes, info.ComponentCount, info.BitsPerComponent, info.InvertedCmyk);
        }

        using DecodedPixels decoded = DecodePixelsOwned(contentType, bytes, cancellationToken);
        return PdfImageXObject.RgbPng(decoded.Width, decoded.Height, recolorRgb(decoded.Rgb), decoded.Alpha);
    }

    public static DecodedPixels DecodePixelsOwned(string contentType, byte[] bytes, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (IsTiffContentType(contentType)) return TiffImage.ReadOwned(bytes, cancellationToken);

        if (IsPngContentType(contentType))
        {
            return PngImage.ReadOwned(bytes, cancellationToken);
        }

        if (IsBmpContentType(contentType))
        {
            return BmpImage.ReadOwned(bytes, cancellationToken);
        }

        if (IsGifContentType(contentType))
        {
            return GifImage.ReadOwned(bytes, cancellationToken);
        }

        // Only bitmap-only metafiles reduce to pixels; vector EMF content is drawn by
        // the renderers through EmfScene and EmfPdfWriter.
        if (IsEmfContentType(contentType))
        {
            return EmfScene.Read(bytes, cancellationToken).DecodeSingleBitmap(cancellationToken);
        }

        if (IsJpegContentType(contentType))
        {
            return JpegImage.ReadOwned(bytes, cancellationToken);
        }

        throw new NotSupportedException("Unsupported image content type.");
    }
}
