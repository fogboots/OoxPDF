using System.Globalization;
using System.Xml.Linq;

using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static PptxTextBodyProperties ReadTextBodyProperties(XElement textBody, XElement? inheritedTextBody)
    {
        (int columnCount, double columnSpacing, PptxTextBodyPropertySource columnCountSource, PptxTextBodyPropertySource columnSpacingSource, string? columnCountValue, string? columnSpacingValue) =
            ReadTextColumns(textBody, inheritedTextBody);
        PptxTextBodyPropertySource columnSource = MergeTextBodyPropertySources(columnCountSource, columnSpacingSource);
        (TextInsets insets, TextInsetSources insetSources, TextInsetValues insetValues) = ReadTextInsets(textBody, inheritedTextBody);
        (XElement? autofit, string autofitMode, PptxTextBodyPropertySource autofitModeSource) = ReadTextAutofit(textBody, inheritedTextBody);
        (double fontScale, PptxTextBodyPropertySource fontScaleSource, string? fontScaleValue) = ReadNormAutofitFontScale(autofit, autofitModeSource);
        (double lineSpacingScale, PptxTextBodyPropertySource lineSpacingScaleSource, string? lineSpacingReductionValue) = ReadNormAutofitLineSpacingScale(autofit, autofitModeSource);
        (string? orientation, PptxTextBodyPropertySource orientationSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "vert", inherit: true);
        (string? verticalAnchor, PptxTextBodyPropertySource verticalAnchorSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "anchor", inherit: true);
        (string? anchorCenter, PptxTextBodyPropertySource anchorCenterSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "anchorCtr", inherit: true);
        (string? wrap, PptxTextBodyPropertySource wrapSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "wrap", inherit: true);
        (string? verticalOverflow, PptxTextBodyPropertySource verticalOverflowSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "vertOverflow", inherit: true);
        (string? horizontalOverflow, _) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "horzOverflow", inherit: true);
        (string? compatibleLineSpacing, PptxTextBodyPropertySource compatibleLineSpacingSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "compatLnSpc", inherit: true);
        (string? rotation, PptxTextBodyPropertySource rotationSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "rot", inherit: true);
        return new PptxTextBodyProperties(
            insets,
            insetSources,
            insetValues,
            ParseTextOrientation(orientation),
            orientation,
            orientationSource,
            ParseTextVerticalAnchor(verticalAnchor),
            verticalAnchor,
            verticalAnchorSource,
            anchorCenter is null ? null : OoxBoolean.IsTrue(anchorCenter),
            anchorCenter,
            anchorCenterSource,
            ParseTextWrapMode(wrap),
            wrap,
            wrapSource,
            ParseTextVerticalOverflow(verticalOverflow),
            verticalOverflow,
            verticalOverflowSource,
            columnCount,
            columnSpacing,
            columnSource,
            columnCountSource,
            columnSpacingSource,
            columnCountValue,
            columnSpacingValue,
            autofitMode,
            autofitModeSource,
            fontScale,
            fontScaleValue,
            fontScaleSource,
            lineSpacingScale,
            lineSpacingReductionValue,
            lineSpacingScaleSource,
            compatibleLineSpacing is not null && OoxBoolean.IsTrue(compatibleLineSpacing),
            compatibleLineSpacing,
            compatibleLineSpacingSource,
            ParseTextBodyRotationDegrees(rotation),
            rotation,
            rotationSource,
            ExplicitWrapWidth: null)
        {
            AutofitMode = ParseTextAutofitMode(autofitMode),
            ClipsHorizontalOverflow = string.Equals(horizontalOverflow, "clip", StringComparison.OrdinalIgnoreCase)
        };

        static PptxTextBodyPropertySource MergeTextBodyPropertySources(
            PptxTextBodyPropertySource first,
            PptxTextBodyPropertySource second)
        {
            if (first == second)
            {
                return first;
            }

            if (first == PptxTextBodyPropertySource.DirectBodyPr || second == PptxTextBodyPropertySource.DirectBodyPr)
            {
                return PptxTextBodyPropertySource.DirectBodyPr;
            }

            if (first == PptxTextBodyPropertySource.InheritedBodyPr || second == PptxTextBodyPropertySource.InheritedBodyPr)
            {
                return PptxTextBodyPropertySource.InheritedBodyPr;
            }

            if (first == PptxTextBodyPropertySource.TableCellStyle || second == PptxTextBodyPropertySource.TableCellStyle)
            {
                return PptxTextBodyPropertySource.TableCellStyle;
            }

            return PptxTextBodyPropertySource.DefaultValue;
        }
    }

    private static (string? Value, PptxTextBodyPropertySource Source) ReadTextBodyAttributeWithSource(
        XElement textBody,
        XElement? inheritedTextBody,
        string attributeName,
        bool inherit)
    {
        XElement? bodyPr = textBody.Element(DrawingNamespace + "bodyPr");
        if (bodyPr?.Attribute(attributeName) is { } directAttribute)
        {
            return (directAttribute.Value, PptxTextBodyPropertySource.DirectBodyPr);
        }

        XElement? inheritedBodyPr = inheritedTextBody?.Element(DrawingNamespace + "bodyPr");
        if (inherit && inheritedBodyPr?.Attribute(attributeName) is { } inheritedAttribute)
        {
            return (inheritedAttribute.Value, PptxTextBodyPropertySource.InheritedBodyPr);
        }

        return (null, PptxTextBodyPropertySource.DefaultValue);
    }

    private static TextInsets ReadPresetTextRectInsets(XElement shape, double width, double height)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        string? preset = shapeProperties
            ?.Element(DrawingNamespace + "prstGeom")
            ?.Attribute("prst")
            ?.Value;

        return preset switch
        {
            "ellipse" => new TextInsets(
                width * PptxTextMetricRules.EllipseTextRectInsetRatio,
                width * PptxTextMetricRules.EllipseTextRectInsetRatio,
                height * PptxTextMetricRules.EllipseTextRectInsetRatio,
                height * PptxTextMetricRules.EllipseTextRectInsetRatio),
            "roundRect" when shapeProperties is not null => RoundRectTextRectInsets(shapeProperties, width, height),
            _ when preset is not null && TryReadPresetTextRectInsets(preset, shapeProperties, width, height, out TextInsets presetInsets) => presetInsets,
            _ => TextInsets.Empty
        };
    }

    private static TextInsets RoundRectTextRectInsets(XElement shapeProperties, double width, double height)
    {
        double adjustment = Math.Clamp(
            ReadPresetGeometryGuide(
                shapeProperties,
                presetAdjustmentsOverride: null,
                "adj",
                PptxTextMetricRules.RoundRectDefaultAdjustment),
            0d,
            50000d) / 100000d;
        double radius = Math.Min(width, height) * adjustment;
        double inset = radius * PptxTextMetricRules.RoundRectTextRectRadiusInsetFactor;
        return new TextInsets(inset, inset, inset, inset);
    }
}
