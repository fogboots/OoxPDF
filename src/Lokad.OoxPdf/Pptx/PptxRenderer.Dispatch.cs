using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Q02: orphan image/chart-font registrations from rewound nodes are inert but
    // would still serialize as unreferenced PDF objects. Prune each page to the
    // names referenced by surviving content. References always emit "/{name} "
    // (Do/Tf), so the trailing space guards prefix collisions (Im1 vs Im12).
    // Conservative: a name that never appears is dropped; anything else stays.
    internal static List<PdfImageResource> PruneUnreferencedImages(string content, List<PdfImageResource> images, CancellationToken cancellationToken = default)
    {
        return PruneUnreferencedResources(content, images, static image => image.ResourceName, cancellationToken);
    }

    internal static List<PdfFontResource> PruneUnreferencedChartFonts(string content, List<PdfFontResource> fonts, CancellationToken cancellationToken = default)
    {
        return PruneUnreferencedResources(content, fonts, static font => font.ResourceName, cancellationToken);
    }

    private static List<T> PruneUnreferencedResources<T>(string content, List<T> resources, Func<T, string> resourceName, CancellationToken cancellationToken)
    {
        if (resources.Count == 0)
        {
            return resources;
        }

        var wanted = new HashSet<string>(StringComparer.Ordinal);
        foreach (T resource in resources)
        {
            cancellationToken.ThrowIfCancellationRequested();
            wanted.Add(PdfEmbeddedFont.SanitizeName(resourceName(resource)));
        }

        var referenced = new HashSet<string>(StringComparer.Ordinal);
        int position = 0;
        while (position < content.Length)
        {
            if ((position & 4095) == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            if (content[position] != '/')
            {
                position++;
                continue;
            }

            int tokenStart = position + 1;
            int tokenEnd = tokenStart;
            while (tokenEnd < content.Length && !char.IsWhiteSpace(content[tokenEnd]))
            {
                tokenEnd++;
            }

            if (tokenEnd < content.Length && content[tokenEnd] == ' ' && tokenEnd > tokenStart)
            {
                string token = content.Substring(tokenStart, tokenEnd - tokenStart);
                if (wanted.Contains(token))
                {
                    referenced.Add(token);
                    if (referenced.Count == wanted.Count)
                    {
                        break;
                    }
                }
            }

            position = tokenEnd + 1;
        }

        List<T>? pruned = null;
        for (int i = 0; i < resources.Count; i++)
        {
            if ((i & 63) == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            if (referenced.Contains(PdfEmbeddedFont.SanitizeName(resourceName(resources[i]))))
            {
                if (pruned is not null)
                {
                    pruned.Add(resources[i]);
                }

                continue;
            }

            pruned ??= resources.Take(i).ToList();
        }

        return pruned ?? resources;
    }

    private static bool ShouldRenderSceneNode(PptxSceneNode node, bool renderPlaceholders)
    {
        return node.Kind != PptxSceneNodeKind.Shape || renderPlaceholders || !node.IsPlaceholder;
    }

    private static IEnumerable<PptxSceneNode> EnumerateRenderedSceneNodes(IReadOnlyList<PptxSceneNode> nodes, bool renderPlaceholders)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (!ShouldRenderSceneNode(node, renderPlaceholders))
                continue;
            yield return node;
            if (node.Kind == PptxSceneNodeKind.Group)
            {
                foreach (PptxSceneNode child in EnumerateRenderedSceneNodes(node.Children, renderPlaceholders))
                    yield return child;
            }
        }
    }

    private static void RenderOrderedSceneNodes(
        IReadOnlyList<PptxSceneNode> nodes,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<FontRequest, RenderedFont> fonts,
        List<PdfImageResource> images,
        List<PdfFontResource> chartFonts,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        string? sourcePartName,
        PptxColorMap sourceColorMap,
        ref int imageIndex,
        GroupTransform transform,
        bool renderPlaceholders,
        CancellationToken cancellationToken)
    {
        foreach (PptxSceneNode node in nodes)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PdfGraphicsBuilder.ContentMark contentMark = graphics.MarkContent();
            int annotationCount = linkAnnotations.Count;
            try
            {
                RenderOrderedSceneNode(
                    node,
                    context,
                    graphics,
                    fonts,
                    images,
                    chartFonts,
                    linkAnnotations,
                    reportedHyperlinkIds,
                    sourcePartName,
                    sourceColorMap,
                    ref imageIndex,
                    transform,
                    renderPlaceholders,
                    cancellationToken);
            }
            catch (Exception ex) when (IsRecoverableNodeRenderException(ex))
            {
                graphics.TruncateContent(contentMark);
                linkAnnotations.RemoveRange(annotationCount, linkAnnotations.Count - annotationCount);
                EmitNodeRenderFailureDiagnostic(node, ex);
            }
        }

        // Q02: preserve the failure category (exception type only, never the message,
        // which can carry document text) so repeated failures stay triageable without
        // leaking content. Budget and cancellation failures never reach here: they are
        // outside the recoverable whitelist below and abort the conversion.
        void EmitNodeRenderFailureDiagnostic(PptxSceneNode node, Exception cause)
        {
            context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "PPTX_NODE_RENDER_FAILED",
                OoxPdfSeverity.Warning,
                "PPTX node rendering failed and the node was ignored while rendering continued. Cause: " + cause.GetType().Name + ".",
                sourcePartName ?? context.SlidePartName,
                PageIndex: null,
                SlideIndex: context.SlideNumber,
                Feature: node.Kind.ToString(),
                Fallback: "Ignored"));
        }

        bool IsRecoverableNodeRenderException(Exception exception)
        {
            if (exception is OoxPdfLimitExceededException)
            {
                return false;
            }

            return exception is FormatException or InvalidDataException or NotSupportedException or ArgumentException;
        }
    }

    private static void RenderOrderedSceneNode(
        PptxSceneNode node,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<FontRequest, RenderedFont> fonts,
        List<PdfImageResource> images,
        List<PdfFontResource> chartFonts,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        string? sourcePartName,
        PptxColorMap sourceColorMap,
        ref int imageIndex,
        GroupTransform transform,
        bool renderPlaceholders,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (node.Kind == PptxSceneNodeKind.Shape && (renderPlaceholders || !node.IsPlaceholder) &&
            node.Source.Annotation<PptxNativeMath>() is not null && TryRenderNativeMath(node, context, graphics, chartFonts, transform)) return;
        switch (node.Kind)
        {
            case PptxSceneNodeKind.Shape:
                if (ShouldRenderSceneNode(node, renderPlaceholders))
                {
                    if (HasDrawableShape())
                    {
                        BeginSlideNodeClip(context, graphics);
                        RenderShape(
                            node,
                            context.Document,
                            graphics,
                            context.DiagnosticSink,
                            context.SlideNumber,
                            context.Theme,
                            transform,
                            images,
                            context.ImageCache,
                            ref imageIndex);
                        EndSlideNodeClip(graphics);
                    }

                    // Node-level links go first so run-level links sit above them.
                    AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                    RenderTextNode(
                        node,
                        context,
                        graphics,
                        fonts,
                        sourceColorMap,
                        renderPlaceholders,
                        linkAnnotations,
                        reportedHyperlinkIds,
                        sourcePartName);
                }

                break;
            case PptxSceneNodeKind.Connector:
                BeginSlideNodeClip(context, graphics);
                RenderShape(
                    node,
                    context.Document,
                    graphics,
                    context.DiagnosticSink,
                    context.SlideNumber,
                    context.Theme,
                    transform,
                    images,
                    context.ImageCache,
                    ref imageIndex);
                EndSlideNodeClip(graphics);
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                break;
            case PptxSceneNodeKind.Picture:
                RenderPicture(
                    node,
                    context,
                    graphics,
                    transform,
                    images,
                    ref imageIndex,
                    sourcePartName);
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                break;
            case PptxSceneNodeKind.Table:
                BeginSlideNodeClip(context, graphics);
                // Frame-level links go first so cell-run links sit above them.
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = RenderTableFrame(context, node, graphics, transform, sourceColorMap);
                var tableHyperlinkScope = new PptxTextHyperlinkScope(context, sourcePartName, linkAnnotations, reportedHyperlinkIds);
                DrawTextSpansWithFonts(tableTextSpans, graphics, fonts, tableHyperlinkScope);
                EndSlideNodeClip(graphics);
                break;
            case PptxSceneNodeKind.Chart:
                BeginSlideNodeClip(context, graphics);
                cancellationToken.ThrowIfCancellationRequested();
                // Node-level links go first so chart run links sit above them.
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                RenderChartFrame(context, graphics, chartFonts, node, transform, linkAnnotations, reportedHyperlinkIds);
                EndSlideNodeClip(graphics);
                break;
            case PptxSceneNodeKind.UnknownGraphicFrame:
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                RenderUnsupportedGraphicFrame(node, context, sourcePartName);
                break;
            case PptxSceneNodeKind.Group:
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                RenderOrderedSceneNodes(
                    node.Children,
                    context,
                    graphics,
                    fonts,
                    images,
                    chartFonts,
                    linkAnnotations,
                    reportedHyperlinkIds,
                    sourcePartName,
                    sourceColorMap,
                    ref imageIndex,
                    transform.Combine(ToGroupTransform(node.GroupTransform)),
                    renderPlaceholders,
                    cancellationToken);
                break;
        }

        bool HasDrawableShape()
        {
            return node.Shape is
            {
                Fill.HasFill: true
            } or
            {
                GradientFill.HasGradient: true
            } or
            {
                PatternFill.HasPattern: true
            } or
            {
                PictureFill.HasPicture: true
            } or
            {
                Glow.HasGlow: true
            } or
            {
                OuterShadow.HasShadow: true
            } or
            {
                Line.HasLine: true
            };
        }
    }

    private static void BeginSlideNodeClip(PptxRenderContext context, PdfGraphicsBuilder graphics)
    {
        graphics.SaveState();
        ClipSlideBoundsEvenOdd(context.Document, graphics);
    }

    private static void EndSlideNodeClip(PdfGraphicsBuilder graphics)
    {
        graphics.RestoreState();
    }

    private static void RenderTextNode(
        PptxSceneNode node,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<FontRequest, RenderedFont> fonts,
        PptxColorMap sourceColorMap,
        bool renderPlaceholders,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        string? sourcePartName)
    {
        if (node.TextBody is null)
        {
            return;
        }

        var hyperlinkScope = new PptxTextHyperlinkScope(context, sourcePartName, linkAnnotations, reportedHyperlinkIds);
        DrawTextSpansWithFonts(ReadTextSpansForSceneNode(node, context, sourceColorMap, renderPlaceholders), graphics, fonts, hyperlinkScope);
    }

    private static void RenderUnsupportedGraphicFrame(PptxSceneNode node, PptxRenderContext context, string? sourcePartName)
    {
        string effectivePartName = sourcePartName ?? context.SlidePartName;
        if (context.DiagnosticSink is null || effectivePartName == context.SlidePartName || node.IsSmartArtGraphicFrame)
        {
            return;
        }

        context.DiagnosticSink(new OoxPdfDiagnostic(
            "PPTX_UNSUPPORTED_GRAPHIC_FRAME",
            OoxPdfSeverity.Warning,
            "Unsupported PPTX graphic frame was detected and ignored.",
            effectivePartName,
            PageIndex: null,
            SlideIndex: context.SlideNumber,
            Feature: "graphic frame",
            Fallback: "Ignored"));
    }

}

