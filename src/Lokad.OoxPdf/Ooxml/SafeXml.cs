using System.Xml;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal static class SafeXml
{
    // Generous bounds: genuine Office parts nest shallowly (< 40 levels) and
    // arrive through the 64 MB package part cap. Anything beyond fails fast
    // instead of amplifying a small input into a huge DOM. Element count is capped
    // separately so broad shallow parts cannot inflate the object graph without bound.
    internal const long DefaultMaxCharacters = 64L * 1024L * 1024L;
    internal const int DefaultMaxDepth = 256;
    internal const long DefaultMaxNodes = 2000000L;
    internal const long DefaultMaxAttributes = 2000000L;

    internal static XmlReaderSettings CreateReaderSettings(long maxCharacters)
    {
        return new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = maxCharacters
        };
    }

    public static XDocument Load(Stream stream, CancellationToken cancellationToken)
    {
        return Load(stream, cancellationToken, DefaultMaxCharacters, DefaultMaxDepth, DefaultMaxNodes);
    }

    internal static XDocument Load(Stream stream, CancellationToken cancellationToken, long maxCharacters, int maxDepth, long maxNodes = DefaultMaxNodes, long maxAttributes = DefaultMaxAttributes)
    {
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            using XmlReader reader = XmlReader.Create(stream, CreateReaderSettings(maxCharacters));
            using var bounded = new DepthBoundReader(reader, maxDepth, maxNodes, maxAttributes, cancellationToken);
            XDocument document = XDocument.Load(bounded, LoadOptions.None);
            OoxMarkupCompatibility.Normalize(document, cancellationToken);
            OoxVmlDrawing.Normalize(document, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return document;
        }
        catch (XmlException ex)
        {
            throw new InvalidDataException($"Malformed XML content: {ex.Message}", ex);
        }
    }

    // XmlReaderSettings exposes no depth quota, so enforce it here while also
    // surfacing cancellation during long parses. All members delegate.
    private sealed class DepthBoundReader(XmlReader inner, int maxDepth, long maxNodes, long maxAttributes, CancellationToken cancellationToken) : XmlReader
    {
        private int depth;
        private long nodes;
        private long elementCount;
        private long attributeCount;
        public override int AttributeCount => inner.AttributeCount;
        public override string BaseURI => inner.BaseURI;
        public override int Depth => inner.Depth;
        public override bool EOF => inner.EOF;
        public override bool IsEmptyElement => inner.IsEmptyElement;
        public override string LocalName => inner.LocalName;
        public override string? LookupNamespace(string prefix) => inner.LookupNamespace(prefix);
        public override bool MoveToAttribute(string name) => inner.MoveToAttribute(name);
        public override bool MoveToAttribute(string name, string? ns) => inner.MoveToAttribute(name, ns);
        public override bool MoveToElement() => inner.MoveToElement();
        public override bool MoveToFirstAttribute() => inner.MoveToFirstAttribute();
        public override bool MoveToNextAttribute() => inner.MoveToNextAttribute();
        public override XmlNameTable NameTable => inner.NameTable;
        public override string NamespaceURI => inner.NamespaceURI;
        public override XmlNodeType NodeType => inner.NodeType;
        public override string Prefix => inner.Prefix;
        public override bool ReadAttributeValue() => inner.ReadAttributeValue();
        public override ReadState ReadState => inner.ReadState;
        public override string Value => inner.Value;

        public override bool Read()
        {
            bool result = inner.Read();
            if (result)
            {
                if (NodeType == XmlNodeType.Element)
                {
                    elementCount++;
                    if (elementCount > maxNodes)
                    {
                        throw new OoxPdfLimitExceededException($"XML element count exceeds the maximum supported count of {maxNodes}.");
                    }

                    // the element cap alone misses attribute-driven object
                    // amplification (one element with 1000 attributes passed with maxNodes=3).
                    // Count attributes per element toward a separate object budget.
                    int currentAttributes = inner.AttributeCount;
                    if (currentAttributes > 0)
                    {
                        long next;
                        try
                        {
                            next = checked(attributeCount + currentAttributes);
                        }
                        catch (OverflowException ex)
                        {
                            throw new OoxPdfLimitExceededException("XML attribute count overflows.", ex);
                        }

                        attributeCount = next;
                        if (attributeCount > maxAttributes)
                        {
                            throw new OoxPdfLimitExceededException($"XML attribute count exceeds the maximum supported count of {maxAttributes}.");
                        }
                    }
                }

                if (NodeType == XmlNodeType.Element && !IsEmptyElement)
                {
                    depth++;
                    if (depth > maxDepth)
                    {
                        throw new OoxPdfLimitExceededException($"XML element nesting exceeds the maximum supported depth of {maxDepth}.");
                    }
                }
                else if (NodeType == XmlNodeType.EndElement)
                {
                    depth--;
                }

                if (++nodes % 4096 == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }
            }

            return result;
        }

        public override string? GetAttribute(string name) => inner.GetAttribute(name);
        public override string? GetAttribute(string name, string? namespaceURI) => inner.GetAttribute(name, namespaceURI);
        public override string GetAttribute(int i) => inner.GetAttribute(i);
        public override void ResolveEntity() => inner.ResolveEntity();
    }
}
