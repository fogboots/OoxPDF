using System.Xml.Linq;

using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Ooxml;

// ISO 29500 Markup Compatibility subset (O02): exactly one AlternateContent
// representation survives load, so Descendants traversals can neither duplicate
// Choice and Fallback content nor silently keep both. The first Choice whose
// Requires prefixes all resolve to understood namespaces wins, else the Fallback,
// else the block is removed. Declared ProcessContent wrappers expose their children
// before format readers run. Other extension content keeps its existing fallback;
// this is a bounded subset, not a complete MCE/schema processor.
internal static class OoxMarkupCompatibility
{
    private static readonly HashSet<string> UnderstoodNamespaces = new(
        [
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/chart",

            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",

            // RV05 Word-authored floating textboxes (edge-floatbox.docx, Word COM): the
            // anchor hides in mc:Choice Requires="wps" with a VML Fallback. Keeping the
            // wps representation routes the box through floating layout instead of leaking
            // its text into body flow; other 2010-era namespaces stay on fallback.
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",

            "urn:schemas-microsoft-com:vml"
        ],
        StringComparer.Ordinal);

    // MustUnderstand scan (O02): true when an element demands a namespace outside
    // the understood set. Readers warn once per document; slides and DOCX story parts warn with their part names.
    public static bool HasUnrecognizedMustUnderstand(XDocument? document)
    {
        if (document?.Root is null)
        {
            return false;
        }

        // Normalization may remove the declaration's carrier or namespace scope.
        // Retain only the diagnostic fact, resolved while that scope still exists.
        return document.Annotation<MustUnderstandResult>()?.HasUnrecognized
            ?? document.Descendants().Any(HasUnrecognizedMustUnderstand);
    }

    private static bool HasUnrecognizedMustUnderstand(XElement element)
    {
        string? demand = (string?)element.Attribute(OoxNamespaces.MarkupCompatibilityNamespace + "MustUnderstand");
        return demand is not null && demand.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
            .Any(prefix => !UnderstoodNamespaces.Contains(element.GetNamespaceOfPrefix(prefix)?.NamespaceName ?? string.Empty));
    }

    private sealed record MustUnderstandResult(bool HasUnrecognized);

    // O02: one must-understand warning per part per document. Slides warn with
    // their slide index; shared master/layout parts warn once no matter how many
    // slides use them; DOCX story parts warn without one. Repeat loads of the
    // same part stay quiet through the gate.
    public static void WarnMustUnderstandOnce(
        XDocument? document,
        string? partName,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<string>? warnedParts,
        int? slideIndex = null)
    {
        if (diagnosticSink is null || warnedParts is null || !HasUnrecognizedMustUnderstand(document))
        {
            return;
        }

        if (!warnedParts.Add(partName ?? ""))
        {
            return;
        }

        diagnosticSink(new OoxPdfDiagnostic(
            "OOXML_MUST_UNDERSTAND",
            OoxPdfSeverity.Warning,
            "A must-understand declaration requires unsupported or unresolved namespaces; conversion continues with partial compatibility handling.",
            partName,
            SlideIndex: slideIndex,
            PageIndex: null,
            Feature: "must-understand",
            Fallback: "Partial compatibility"));
    }

    public static void Normalize(XDocument? document, CancellationToken cancellationToken = default)
    {
        if (document?.Root is null)
        {
            return;
        }

        bool unrecognized = document.Annotation<MustUnderstandResult>()?.HasUnrecognized ?? false;
        Visit(document.Root);
        document.RemoveAnnotations<MustUnderstandResult>();
        document.AddAnnotation(new MustUnderstandResult(unrecognized));

        void Visit(XElement element)
        {
            cancellationToken.ThrowIfCancellationRequested();
            XNamespace compatibility = OoxNamespaces.MarkupCompatibilityNamespace;
            if (element.Name == compatibility + "AlternateContent")
            {
                unrecognized |= HasUnrecognizedMustUnderstand(element);
                XElement? selected = element.Elements(compatibility + "Choice").FirstOrDefault(IsUnderstoodChoice);
                // The VML branch of an OLE object references a legacy drawing shape; the
                // fallback carries the same object with its preview picture, which is
                // what gets painted.
                if (selected is not null && HasOlePreview(element.Element(compatibility + "Fallback")) && !HasOlePreview(selected))
                {
                    selected = null;
                }
                if (selected is null)
                {
                    selected = element.Element(compatibility + "Fallback");
                    if (selected is null) { element.Remove(); return; }
                    PreservePreferredChartGalleryStyle(element, selected);
                    Lokad.OoxPdf.Pptx.PptxNativeMath.Preserve(element, selected);
                }
                unrecognized |= HasUnrecognizedMustUnderstand(selected);
                VisitChildren(selected);
                Unwrap(element, selected);
                return;
            }

            bool ignored = IsIgnorableExtension(element, out bool processContent, cancellationToken);
            if (ignored && !processContent) { return; }
            unrecognized |= HasUnrecognizedMustUnderstand(element);
            VisitChildren(element);
            if (ignored && processContent) { Unwrap(element, element); }
        }

        void VisitChildren(XElement parent)
        {
            // Child normalization can replace its container; retain the next sibling
            // without allocating a snapshot for every ordinary XML element.
            for (XNode? child = parent.FirstNode; child is not null;)
            {
                XNode? next = child.NextNode;
                if (child is XElement element) { Visit(element); }
                child = next;
            }
        }
    }

    private static bool HasOlePreview(XElement? branch)
    {
        return branch?.Elements(OoxNamespaces.PresentationNamespace + "oleObj").Elements(OoxNamespaces.PresentationNamespace + "pic").Any() == true;
    }

    private static bool IsIgnorableExtension(XElement element, out bool processContent, CancellationToken cancellationToken)
    {
        processContent = false;
        if (UnderstoodNamespaces.Contains(element.Name.NamespaceName)) { return false; }
        bool ignorable = false;
        XNamespace compatibility = OoxNamespaces.MarkupCompatibilityNamespace;
        // Resolve inherited declarations at their owner, before any wrapper is
        // removed. Rebinding a prefix below that owner must not retarget its rule.
        // The XML depth bound limits this walk; no per-element scope copies remain.
        foreach (XElement scope in element.AncestorsAndSelf())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (scope.Attribute(compatibility + "Ignorable") is XAttribute ignored)
            {
                foreach (string prefix in ignored.Value.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
                {
                    ignorable |= scope.GetNamespaceOfPrefix(prefix)?.NamespaceName == element.Name.NamespaceName;
                }
            }
            if (scope.Attribute(compatibility + "ProcessContent") is XAttribute content)
            {
                foreach (string name in content.Value.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
                {
                    int colon = name.IndexOf(':');
                    if (colon <= 0 || colon == name.Length - 1) { continue; }
                    processContent |= scope.GetNamespaceOfPrefix(name[..colon])?.NamespaceName == element.Name.NamespaceName &&
                        (name[(colon + 1)..] == element.Name.LocalName || name[(colon + 1)..] == "*");
                }
            }
            if (ignorable && processContent) { break; }
        }
        return ignorable;
    }

    private static void Unwrap(XElement wrapper, XElement content)
    {
        XNode[] children = content.Nodes().ToArray();
        // Detach instead of cloning so existing semantic annotations, including
        // preferred chart-gallery styles, survive container removal.
        foreach (XNode child in children) { child.Remove(); }
        wrapper.ReplaceWith(children);
    }

    private static void PreservePreferredChartGalleryStyle(XElement alternate, XElement fallback)
    {
        XElement? chartSpace = alternate.Parent;
        if (chartSpace?.Name != OoxNamespaces.ChartNamespace + "chartSpace")
        {
            return;
        }

        XNamespace compatibility = OoxNamespaces.MarkupCompatibilityNamespace;
        // Preserve the preferred chart-gallery declaration without opting into
        // other unsupported c14 content or changing the selected fallback tree.
        XElement? preferred = alternate.Elements(compatibility + "Choice").FirstOrDefault();
        XElement? preferredStyle = preferred is not null && !preferred.Elements().Skip(1).Any() ? preferred.Elements().FirstOrDefault() : null;
        XElement? fallbackStyle = !fallback.Elements().Skip(1).Any() ? fallback.Elements().FirstOrDefault() : null;
        string[] required = ((string?)preferred?.Attribute("Requires") ?? string.Empty)
            .Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        if (required.Length == 1 &&
            preferred?.GetNamespaceOfPrefix(required[0]) == OoxNamespaces.ChartStyle2010Namespace &&
            preferredStyle?.Name == OoxNamespaces.ChartStyle2010Namespace + "style" &&
            fallbackStyle?.Name == OoxNamespaces.ChartNamespace + "style" &&
            int.TryParse((string?)preferredStyle.Attribute("val"), out int style) &&
            style is >= 101 and <= 148)
        {
            chartSpace.AddAnnotation(new ChartGalleryChoice(style));
        }

    }

    private sealed record ChartGalleryChoice(int Value);

    internal static int? ReadPreferredChartGalleryStyle(XElement? chartSpace)
    {
        XElement? direct = chartSpace?.Element(OoxNamespaces.ChartStyle2010Namespace + "style");
        return direct is not null && int.TryParse((string?)direct.Attribute("val"), out int value) && value is >= 101 and <= 148
            ? value
            : chartSpace?.Annotation<ChartGalleryChoice>()?.Value;
    }

    private static bool IsUnderstoodChoice(XElement choice)
    {
        string requires = (string?)choice.Attribute("Requires") ?? string.Empty;
        string[] prefixes = requires.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        if (prefixes.Length == 0)
        {
            return true;
        }

        foreach (string prefix in prefixes)
        {
            if (!UnderstoodNamespaces.Contains(choice.GetNamespaceOfPrefix(prefix)?.NamespaceName ?? string.Empty))
            {
                return false;
            }
        }

        return true;
    }
}
