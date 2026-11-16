using System.Globalization;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Filee.Engines.Hwp.Hwpx;
using Filee.Engines.Infrastructure;

namespace Filee.Engines.Hwp;

internal static partial class RhwpInput
{
    private static IEnumerable<XProcessingInstruction> PageFieldMarkers(XDocument section) =>
        section.DescendantNodes().OfType<XProcessingInstruction>().Where(p => p.Target == HwpxWriter.PageFieldMarker);

    /// <summary>
    /// rhwp doesn't draw page fields in a table inside a text box. For first-page headers written by Filee,
    /// replace the cached results with numbers from rhwp's pagination of the prepared HWPX.
    /// </summary>
    internal static async Task<string> ResolveFirstPageFieldsAsync(string exe, string input, List<string> sections,
        string workDirectory, CancellationToken cancellationToken)
    {
        try
        {
            var layout = await ProcessRunner.RunAsync(exe, ["dump-pages", input, "--json"],
                RhwpConverter.TimeLimit(RhwpConverter.TextSize(input)), cancellationToken);
            if (layout.ExitCode != 0)
                return input; // Older renderers may lack the pagination command; the cached text remains readable.

            using var json = JsonDocument.Parse(layout.StandardOutput);
            if (json.RootElement.ValueKind != JsonValueKind.Object
                || !json.RootElement.TryGetProperty("pageCount", out var count) || count.ValueKind != JsonValueKind.Number
                || !count.TryGetInt32(out var total)
                || total <= 0 || !json.RootElement.TryGetProperty("pages", out var pages) || pages.ValueKind != JsonValueKind.Array)
                return input;

            var pagination = new Dictionary<int, (int DisplayPage, int Length)>();
            foreach (var page in pages.EnumerateArray())
            {
                if (page.ValueKind != JsonValueKind.Object
                    || !page.TryGetProperty("section", out var section) || section.ValueKind != JsonValueKind.Number
                    || !section.TryGetInt32(out var index)
                    || !page.TryGetProperty("displayPage", out var display) || display.ValueKind != JsonValueKind.Number
                    || !display.TryGetInt32(out var displayPage))
                    return input;
                if (!pagination.TryGetValue(index, out var first))
                    pagination.Add(index, (displayPage, 1));
                else
                    pagination[index] = (first.DisplayPage, first.Length + 1);
            }

            var copy = Path.Combine(workDirectory, $"rhwp-{Guid.NewGuid():N}.hwpx");
            File.Copy(input, copy);
            using var zip = ZipFile.Open(copy, ZipArchiveMode.Update);
            var parts = new BoundedXml();
            foreach (var name in sections)
            {
                const string prefix = "Contents/section";
                const string extension = ".xml";
                if (!int.TryParse(name.AsSpan(prefix.Length, name.Length - prefix.Length - extension.Length), out var index)
                    || zip.GetEntry(name) is not { } entry)
                    continue;
                var xml = Load(parts, entry);
                var restart = xml.Root?.Elements().FirstOrDefault()?.Elements()
                    .SelectMany(run => run.Elements().Where(e => e.Name.LocalName == "ctrl"))
                    .SelectMany(control => control.Elements())
                    .FirstOrDefault(control => control.Name.LocalName == "newNum" && (string?)control.Attribute("numType") == "PAGE");
                var start = int.TryParse((string?)restart?.Attribute("num"), out var firstNumber) ? firstNumber : 0;
                foreach (var marker in PageFieldMarkers(xml).ToList())
                {
                    if (marker.NextNode is XElement { Name.LocalName: "t" } text)
                    {
                        var section = pagination.GetValueOrDefault(index);
                        var number = marker.Data switch
                        {
                            nameof(HFieldKind.PageNumber) => start > 0 ? start : section.DisplayPage,
                            nameof(HFieldKind.SectionPages) => section.Length,
                            nameof(HFieldKind.TotalPages) => total,
                            _ => 0,
                        };
                        if (number > 0)
                            text.Value = number.ToString(CultureInfo.InvariantCulture);
                    }
                    marker.Remove();
                }
                entry.Delete();
                using var stream = zip.CreateEntry(name, CompressionLevel.Optimal).Open();
                xml.Save(stream, SaveOptions.DisableFormatting);
            }
            return copy;
        }
        catch (Exception ex) when (ex is TimeoutException or JsonException or InvalidDataException or System.Xml.XmlException)
        {
            return input;
        }
    }
}
