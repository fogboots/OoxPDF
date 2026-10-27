using System.Globalization;
using System.Text;
namespace Lokad.OoxPdf.Pdf;

internal sealed record PdfManifestTextRun(string Font, string Operator, double X, double Y, double Size,
    string? Actual, double CharacterSpacing = 0);
internal static class PdfManifestText
{
    public static string Read(IReadOnlyList<PdfManifestTextRun> runs, IReadOnlyList<PdfFontResource> fonts, List<string> warnings)
    {
        var resources = fonts.GroupBy(f => f.ResourceName).ToDictionary(g => g.Key, g => g.First().Font);
        var widths = resources.ToDictionary(p => p.Key, p => p.Value.ManifestWidths());
        var text = new StringBuilder();
        double? y = null, endX = null;
        foreach (var run in runs)
        {
            bool newLine = y is not null && Math.Abs(run.Y-y.Value) > Math.Max(2, run.Size * .6);
            if (newLine) text.Append('\n');
            else if (endX is not null && run.X-endX.Value > run.Size * .12 && text.Length>0 && !char.IsWhiteSpace(text[^1])) text.Append(' ');
            y = run.Y;
            if (run.Actual is not null) { text.Append(ManifestUnicode.Plain(run.Actual)); endX = null; continue; }
            if (!resources.TryGetValue(run.Font, out var font)) { warnings.Add("manifest_text_unmapped_font"); endX=null; continue; }
            var map = font.UnicodeByCid;
            var advances = widths[run.Font];
            double x = run.X;
            ReadOnlySpan<char> op = run.Operator;
            for (int i = 0; i < op.Length; i++)
            {
                if (op[i] == '<')
                {
                    int end = op[(i+1)..].IndexOf('>');
                    if (end < 0) break;
                    end += i+1;
                    for (int at = i+1; at+4 <= end; at += 4)
                    {
                        if (!ushort.TryParse(op.Slice(at,4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ushort cid) ||
                            !map.TryGetValue(cid, out int point) || !Rune.IsValid(point))
                        { warnings.Add("manifest_text_unmapped_glyph"); text.Append('\uFFFD'); }
                        else text.Append(ManifestUnicode.Plain(new Rune(point).ToString()));
                        x += advances.GetValueOrDefault(cid) * run.Size + run.CharacterSpacing;
                    }
                    i = end;
                }
                else if (char.IsDigit(op[i]) || op[i] is '-' or '+' or '.')
                {
                    int start=i;
                    while(i+1<op.Length && (char.IsDigit(op[i+1]) || op[i+1] is '-' or '+' or '.')) i++;
                    if(double.TryParse(op[start..(i+1)],NumberStyles.Float,CultureInfo.InvariantCulture,out double adjustment))
                    {
                        x -= adjustment * run.Size / 1000d;
                        if(adjustment < -120 && text.Length>0 && !char.IsWhiteSpace(text[^1])) text.Append(' ');
                    }
                }
            }
            endX = x;
        }
        return text.ToString();
    }
}
