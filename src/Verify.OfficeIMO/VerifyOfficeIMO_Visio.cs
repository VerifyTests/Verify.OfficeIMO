using OfficeIMO.Visio;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterVisio()
    {
        VerifierSettings.RegisterStreamConverter(
            "vsdx",
            (_, stream, context) => Convert(VisioDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<VisioDocument>(Convert);
    }

    static ConversionResult Convert(VisioDocument document, IReadOnlyDictionary<string, object> context)
    {
        var pages = document.Pages;
        var info = new
        {
            document.Title,
            document.Author,
            Pages = pages.Select(_ => _.Name).ToList()
        };

        var derived = Pages(
            context,
            pages.Count,
            _ => string.Join(
                "\n",
                pages[_].Shapes
                    .Select(_ => _.Text)
                    .Where(_ => !string.IsNullOrEmpty(_))),
            _ => pages[_].ToImage(Render(new VisioImageExportOptions())).AsPng().Export().Bytes);
        return new(info, DeterministicSource("vsdx", document.ToBytes(), context), derived);
    }
}
