using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
namespace Lokad.OoxPdf;

public sealed record OoxPageObject(string Kind, double Share, bool Template, bool Fallback);
public sealed record OoxPageManifest(int Page, int? SourceIndex, string Text, IReadOnlyList<string> Headings,
    IReadOnlyList<OoxPageObject> Objects, IReadOnlyList<string> MissingFonts, IReadOnlyList<string> Warnings)
{
    [JsonIgnore] internal IReadOnlyList<string> RequestedFonts { get; init; } = [];
}
public sealed record OoxPagesManifest(int Version, string Source, IReadOnlyList<OoxPageManifest> Pages);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(OoxPagesManifest))]
[JsonSerializable(typeof(OoxPageManifest))]
public partial class OoxManifestJsonContext : JsonSerializerContext { }

internal static class ManifestUnicode
{
    public static string Plain(string text)
    {
        if (!text.EnumerateRunes().Any(r => r.Value is >= 0x1D400 and <= 0x1D7FF)) return text;
        var result = new StringBuilder(text.Length);
        foreach (Rune rune in text.EnumerateRunes())
            result.Append(rune.Value is >= 0x1D400 and <= 0x1D7FF ? rune.ToString().Normalize(NormalizationForm.FormKC) : rune.ToString());
        return result.ToString();
    }
}
