using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
namespace Lokad.OoxPdf;

internal static class PageManifestBuilder
{
    public static OoxPageManifest Failed(int page, int? source) => new(page, source, "", [],
        [new("shapes", 0, false, true)], [], ["manifest_page_analysis_failed"]);

    public static OoxPageManifest Pptx(PdfPage pdf, PdfGraphicsBuilder graphics, PptxSceneSlide slide,
        PresentationFontResolver resolver, IEnumerable<string?> requestedFonts, int sourceIndex)
    {
        var objects = new List<OoxPageObject>(); var warnings = new List<string>();
        string[] names = requestedFonts.Where(n => !string.IsNullOrWhiteSpace(n)).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        try
        {
            if (slide.ShowMasterShapes) Walk(slide.MasterNodes, true, 0, 0, 1, 1, false);
            Walk(slide.LayoutNodes, true, 0, 0, 1, 1, false);
            Walk(slide.SlideNodes, false, 0, 0, 1, 1, false);
            return new OoxPageManifest(slide.Index+1, sourceIndex, PdfManifestText.Read(graphics.ManifestText, pdf.Fonts, warnings), [],
                objects, MissingFonts(names, resolver.Resolve), warnings.Distinct().ToArray()) { RequestedFonts = names };
        }
        catch { return Failed(slide.Index+1, sourceIndex); }

        void Walk(IReadOnlyList<PptxSceneNode> nodes, bool template, double ox, double oy, double sx, double sy, bool approximate)
        {
            foreach (var node in nodes)
            {
                if (template && node.IsPlaceholder) continue;
                // A diagram drawn from its saved drawing is one semantic object, not its pieces.
                if (node.IsSmartArtGraphicFrame && node.Kind is PptxSceneNodeKind.Group or PptxSceneNodeKind.Shape && node.Bounds is { } frame)
                {
                    objects.Add(new("smartart", Share((ox+frame.XEmu*sx)/12700d, (oy+frame.YEmu*sy)/12700d,
                        frame.WidthEmu*sx/12700d, frame.HeightEmu*sy/12700d, pdf.Width, pdf.Height), template, approximate || node.Kind == PptxSceneNodeKind.Shape));
                    continue;
                }
                if (node.Kind == PptxSceneNodeKind.Group)
                {
                    var g = node.GroupTransform;
                    Walk(node.Children, template, ox+(g.OffsetX-g.ChildOffsetX*g.ScaleX)*sx,
                        oy+(g.OffsetY-g.ChildOffsetY*g.ScaleY)*sy, sx*g.ScaleX, sy*g.ScaleY,
                        approximate || g.RotationDegrees != 0 || g.FlipHorizontal || g.FlipVertical);
                    continue;
                }
                string? kind = node.Kind switch {
                    PptxSceneNodeKind.Picture => "picture", PptxSceneNodeKind.Chart => "chart",
                    PptxSceneNodeKind.Table => "table", PptxSceneNodeKind.Connector => "shapes",
                    PptxSceneNodeKind.UnknownGraphicFrame => node.IsSmartArtGraphicFrame ? "smartart" : "shapes", _ => null };
                bool fallback = approximate || node.Kind == PptxSceneNodeKind.UnknownGraphicFrame;
                var xml = node.Source;
                if (xml.Descendants(PresentationNamespace + "oleObj").Any()) { kind = "ole"; fallback = true; }
                if (xml.Descendants().Any(e => e.Name.LocalName is "videoFile" or "audioFile" or "media")) { kind = "media"; fallback = true; }
                if (xml.Annotation<PptxNativeMath>() is not null || xml.Descendants(MathNamespace+"oMath").Any()) { kind = "equation"; fallback = true; }
                if (kind is null && node.Shape is { } shape &&
                    (shape.HasCustomGeometry || shape.Preset.Contains("Arrow", StringComparison.OrdinalIgnoreCase))) kind = "shapes";
                if (kind is null) continue;
                if (node.Bounds is not { } b) { objects.Add(new(kind,0,template,true)); warnings.Add("manifest_object_bounds_unavailable"); continue; }
                double x = (ox+b.XEmu*sx)/12700d, y = (oy+b.YEmu*sy)/12700d;
                double w = b.WidthEmu*sx/12700d, h = b.HeightEmu*sy/12700d;
                if (b.RotationDegrees != 0) { fallback = true; warnings.Add("manifest_rotated_bounds_approximate"); }
                if (node.Picture is { Resource: null }) fallback = true;
                objects.Add(new(kind, Share(x,y,w,h,pdf.Width,pdf.Height),template,fallback));
            }
        }
    }

    public static OoxPageManifest Docx(PdfPage pdf, PdfGraphicsBuilder graphics, DocxLayoutPage layout, int page,
        IFontResolver resolver)
    {
        var objects = new List<OoxPageObject>(); var headings = new List<string>(); var warnings = new List<string>();
        var fonts = new List<string?>();
        try
        {
            Walk(layout.Items,false);
            Walk(layout.StaticTextLines,true); Walk(layout.StaticInlineImages,true); Walk(layout.StaticTableRows,true);
            Walk(layout.StaticInlineTextBoxes,true);
            if (pdf.Images.Count > objects.Count(o => o.Kind == "picture"))
            { objects.Add(new("picture",1,false,true)); warnings.Add("manifest_floating_picture_bounds_unavailable"); }
            return new OoxPageManifest(page,null,PdfManifestText.Read(graphics.ManifestText,pdf.Fonts,warnings),headings,
                objects,MissingFonts(fonts,resolver.Resolve),warnings.Distinct().ToArray()) { RequestedFonts = fonts.Where(n=>n is not null).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase).ToArray() };
        }
        catch { return Failed(page,null); }
        void Walk(IEnumerable<DocxLayoutItem> items,bool template)
        {
            foreach(var item in items)
            {
                switch(item)
                {
                    case DocxTextLineLayout line:
                        fonts.Add(line.StyleRun.FontFamily ?? "Calibri");
                        foreach(var segment in line.Segments) fonts.Add(segment.StyleRun.FontFamily ?? "Calibri");
                        if (line.IsFirstParagraphLine == true && line.SourceParagraph?.OutlineLevel is >= 0 and <= 8)
                            headings.Add(ManifestUnicode.Plain(string.Concat(line.SourceParagraph.Runs.Select(r => r.Text))));
                        if(line.Math is not null || line.InlineMath.Count>0)
                            objects.Add(new("equation", Share(line.X, layout.Height-line.BaselineY, line.Width, line.LineHeight ?? line.FontSize,pdf.Width,pdf.Height),template,true));
                        break;
                    case DocxInlineImageLayout image:
                        objects.Add(new("picture",Share(image.X,image.Y,image.Width,image.Height,pdf.Width,pdf.Height),template,image.Image.Bytes is null)); break;
                    case DocxEquationLayout math:
                        objects.Add(new("equation",Math.Clamp(math.Math.Box.Width * (math.Math.Box.AboveBaseline+math.Math.Box.BelowBaseline)*math.Scale*math.Scale/(pdf.Width*pdf.Height),0,1),template,true)); break;
                    case DocxTableRowLayout row:
                        double width = row.Cells.Sum(c => c.Width);
                        objects.Add(new("table", Math.Clamp(width*row.Height/(pdf.Width*pdf.Height),0,1),template,false));
                        foreach(var cell in row.Cells) { Walk(cell.TextLines,template); Walk(cell.InlineImages,template); Walk(cell.NestedRows,template); Walk(cell.InlineTextBoxes,template); }
                        break;
                    case DocxInlineTextBoxLayout box:
                        Walk(box.TextLines,template); Walk(box.TableRows,template); break;
                    default: objects.Add(new("shapes",0,template,true)); warnings.Add("manifest_unclassified_layout_item"); break;
                }
            }
        }
    }
    private static string[] MissingFonts(IEnumerable<string?> names,Func<FontRequest,FontFaceResolution> resolve)
    {
        var missing = new List<string>();
        foreach(string name in names.Where(n=>!string.IsNullOrWhiteSpace(n)).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase))
        { try { if(resolve(new FontRequest(name)).IsFallback) missing.Add(name); } catch { missing.Add(name); } }
        return missing.ToArray();
    }
    private static double Share(double x,double y,double w,double h,double pw,double ph) =>
        Math.Clamp(Math.Max(0,Math.Min(pw,x+w)-Math.Max(0,x))*Math.Max(0,Math.Min(ph,y+h)-Math.Max(0,y))/(pw*ph),0,1);
}
