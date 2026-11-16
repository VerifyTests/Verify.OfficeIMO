namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterPdf()
    {
        VerifierSettings.RegisterStreamConverter(
            "pdf",
            (_, stream, context) => ConvertLoaded(PdfDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<PdfDocument>(Convert);
    }

    // A document is what is authored, so it is saved as the pdf, then read back for the rest
    static ConversionResult Convert(PdfDocument document, IReadOnlyDictionary<string, object> context)
    {
        var bytes = document.ToBytes();
        var result = ConvertLoaded(PdfDocument.Load(new MemoryStream(bytes)), context);
        return new(result.Info, Source("pdf", bytes, context), result.Targets.ToList());
    }

    static ConversionResult ConvertLoaded(PdfDocument document, IReadOnlyDictionary<string, object> context)
    {
        var properties = document.Properties;
        var pageCount = document.GetPageLayouts().Count;
        var info = new
        {
            properties.Title,
            properties.Author,
            properties.Subject,
            properties.Keywords,
            PageCount = pageCount
        };

        var images = new Lazy<IReadOnlyList<OfficeImageExportResult>>(
            () => document.ToImages(Render(new PdfImageExportOptions())));
        var derived = Pages(
            context,
            pageCount,
            _ => document.Pages.Extract(_ + 1).ExtractText(),
            _ => images.Value[_].Bytes);
        return new(info, derived);
    }
}
