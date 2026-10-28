namespace Lokad.OoxPdf.Pdf;

internal readonly record struct PdfPage
{
    public PdfPage(double width, double height)
        : this(width, height, string.Empty)
    {
    }

    public PdfPage(double width, double height, string content)
        : this(width, height, content, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts)
        : this(width, height, content, fonts, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images)
        : this(width, height, content, fonts, images, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images, IReadOnlyList<PdfExtGStateResource> extGStates)
        : this(width, height, content, fonts, images, extGStates, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images, IReadOnlyList<PdfExtGStateResource> extGStates, IReadOnlyList<PdfShadingResource> shadings)
        : this(width, height, content, fonts, images, extGStates, shadings, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images, IReadOnlyList<PdfExtGStateResource> extGStates, IReadOnlyList<PdfShadingResource> shadings, IReadOnlyList<PdfTilingPatternResource> patterns)
        : this(width, height, content, fonts, images, extGStates, shadings, patterns, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images, IReadOnlyList<PdfExtGStateResource> extGStates, IReadOnlyList<PdfShadingResource> shadings, IReadOnlyList<PdfTilingPatternResource> patterns, IReadOnlyList<PdfLinkAnnotation> annotations, IReadOnlyList<PdfFallbackFontResource>? fallbackFonts = null, IReadOnlyList<PdfTransparencyGroupResource>? groups = null, IReadOnlyList<PdfShadingPatternResource>? shadingPatterns = null)
    {
        Width = width;
        Height = height;
        Content = content;
        Fonts = fonts;
        Images = images;
        ExtGStates = extGStates;
        Shadings = shadings;
        Patterns = patterns;
        Annotations = annotations;
        FallbackFonts = fallbackFonts ?? [];
        Groups = groups ?? [];
        ShadingPatterns = shadingPatterns ?? [];
    }

    public double Width { get; }

    public double Height { get; }

    public string Content { get; init; }

    public IReadOnlyList<PdfFontResource> Fonts { get; }

    public IReadOnlyList<PdfFallbackFontResource> FallbackFonts { get; }

    public IReadOnlyList<PdfImageResource> Images { get; }

    public IReadOnlyList<PdfExtGStateResource> ExtGStates { get; }

    public IReadOnlyList<PdfShadingResource> Shadings { get; }

    public IReadOnlyList<PdfTilingPatternResource> Patterns { get; }

    public IReadOnlyList<PdfShadingPatternResource> ShadingPatterns { get; }

    public IReadOnlyList<PdfTransparencyGroupResource> Groups { get; init; }

    public IReadOnlyList<PdfLinkAnnotation> Annotations { get; }
}

internal readonly record struct PdfLinkDestination(int PageIndex, double? Left, double? Top, double? Zoom);

// T02: exactly one target is always set. The constructor is private so the
// both/neither combinations cannot be built through normal builders; only the
// validated factories below can construct values.
internal readonly record struct PdfLinkAnnotation
{
    public double X { get; }

    public double Y { get; }

    public double Width { get; }

    public double Height { get; }

    public string? Uri { get; }

    public PdfLinkDestination? Destination { get; }

    public bool IsUri => Uri is not null && Destination is null;

    public bool IsDestination => Destination is not null && Uri is null;

    private PdfLinkAnnotation(double x, double y, double width, double height, string? uri, PdfLinkDestination? destination)
    {
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Uri = uri;
        Destination = destination;
    }

    internal static PdfLinkAnnotation ToUri(double x, double y, double width, double height, string uri)
    {
        ArgumentException.ThrowIfNullOrEmpty(uri);
        return new PdfLinkAnnotation(x, y, width, height, uri, null);
    }

    internal static PdfLinkAnnotation ToDestination(double x, double y, double width, double height, PdfLinkDestination destination)
    {
        return new PdfLinkAnnotation(x, y, width, height, null, destination);
    }
}
