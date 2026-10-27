using System.Text.Json;
using Lokad.OoxPdf;

internal sealed class PageManifestWriter : IDisposable
{
    private readonly string destination;
    private readonly string staging;
    private FileStream? stream;
    private Utf8JsonWriter? json;
    private bool failed;
    private bool published;
    public PageManifestWriter(string path, string source)
    {
        destination = Path.GetFullPath(path);
        staging = destination + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            stream = new FileStream(staging, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            json = new Utf8JsonWriter(stream);
            json.WriteStartObject(); json.WriteNumber("version",1); json.WriteString("source",source);
            json.WritePropertyName("pages"); json.WriteStartArray();
        }
        catch { failed = true; }
    }
    public void Add(OoxPageManifest page)
    {
        if (failed || json is null) return;
        try { JsonSerializer.Serialize(json,page,OoxManifestJsonContext.Default.OoxPageManifest); json.Flush(); }
        catch { failed = true; }
    }
    public void Publish()
    {
        try
        {
            // A previous successful PDF's metadata must never describe a new PDF.
            File.Delete(destination);
            if (failed || json is null) { Console.Error.WriteLine("Page manifest unavailable; PDF conversion succeeded."); return; }
            json.WriteEndArray(); json.WriteEndObject(); json.Flush(); json.Dispose(); json = null;
            stream!.Dispose(); stream = null;
            File.Move(staging,destination,overwrite:true); published = true;
        }
        catch { Console.Error.WriteLine("Page manifest unavailable; PDF conversion succeeded."); }
    }
    public void Dispose()
    {
        try { json?.Dispose(); } catch { }
        try { stream?.Dispose(); } catch { }
        if (!published) try { File.Delete(staging); } catch { }
    }
}
