using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal sealed class OoxContentTypes
{

    private readonly Dictionary<string, string> defaults;
    private readonly Dictionary<string, string> overrides;

    private OoxContentTypes(Dictionary<string, string> defaults, Dictionary<string, string> overrides)
    {
        this.defaults = defaults;
        this.overrides = overrides;
    }

    public static OoxContentTypes Parse(Stream stream, CancellationToken cancellationToken)
    {
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var conflicted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Two different types for one extension ("xml" and "XML") or one part leave neither
        // guess in force, and an incomplete or unsafe declaration is skipped: the package
        // still opens, and a part without a type reads as absent. Failing the whole document
        // here rejected files Word and LibreOffice open.
        void Declare(Dictionary<string, string> types, string key, string contentType)
        {
            if (conflicted.Contains(key))
            {
                return;
            }

            if (types.TryGetValue(key, out string? existing) && existing != contentType)
            {
                types.Remove(key);
                conflicted.Add(key);
                return;
            }

            types[key] = contentType;
        }

        foreach (XElement element in document.Root?.Elements() ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? contentType = (string?)element.Attribute("ContentType");
            if (contentType is null)
            {
                continue;
            }

            if (element.Name == OoxNamespaces.ContentTypesNamespace + "Default" &&
                (string?)element.Attribute("Extension") is { } extension)
            {
                // Conflict keys keep defaults and overrides apart: an override name starts with '/'.
                Declare(defaults, extension.TrimStart('.'), contentType);
            }
            else if (element.Name == OoxNamespaces.ContentTypesNamespace + "Override" &&
                (string?)element.Attribute("PartName") is { } partName)
            {
                string normalized;
                try
                {
                    normalized = OoxPath.NormalizePartName(partName);
                }
                catch (Exception ex) when (ex is InvalidDataException or ArgumentException)
                {
                    continue;
                }

                Declare(overrides, normalized, contentType);
            }
        }

        return new OoxContentTypes(defaults, overrides);
    }

    public string? GetContentType(string partName)
    {
        string normalized = OoxPath.NormalizePartName(partName);
        if (overrides.TryGetValue(normalized, out string? contentType))
        {
            return contentType;
        }

        string extension = Path.GetExtension(normalized).TrimStart('.');
        return defaults.TryGetValue(extension, out contentType) ? contentType : null;
    }
}
