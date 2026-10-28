namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterDjVu()
    {
        VerifierSettings.RegisterStreamConverter(
            "djvu",
            (_, stream, context) => Convert(DjVuDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<DjVuDocument>(Convert);
    }

    // OfficeIMO reads DjVu but does not write it, so there is no document to verify, only what was read
    static ConversionResult Convert(DjVuDocument document, IReadOnlyDictionary<string, object> context)
    {
        var pages = document.Pages;
        var info = new
        {
            Pages = pages
                .Select(_ => new
                {
                    Title = NullIfEmpty(_.Title),
                    _.Width,
                    _.Height,
                    _.Dpi
                })
                .ToList(),
            Bookmarks = document.Bookmarks.Select(_ => _.Title).ToList()
        };

        var derived = Pages(
            context,
            pages.Count,
            _ => pages[_].GetText().Text,
            _ => OfficeRasterImageEncoder.Encode(pages[_].Render().Image, OfficeImageExportFormat.Png));
        return new(info, derived);
    }
}
