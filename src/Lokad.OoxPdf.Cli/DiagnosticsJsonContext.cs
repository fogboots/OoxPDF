using System.Text.Json.Serialization;
using Lokad.OoxPdf.Diagnostics;

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(IReadOnlyList<OoxPdfDiagnostic>))]
internal partial class DiagnosticsJsonContext : JsonSerializerContext
{
}
