using OfficeIMO.Pdf;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterPdf()
    {
        VerifierSettings.RegisterStreamConverter(
            "pdf",
            (_, stream, context) => Convert(PdfReadDocument.Open(stream), context));
        VerifierSettings.RegisterFileConverter<PdfReadDocument>(Convert);
        VerifierSettings.RegisterFileConverter<PdfDocument>(Convert);
    }

    // A document is what is authored, so it is saved as the pdf, then read back for the rest
    static ConversionResult Convert(PdfDocument document, IReadOnlyDictionary<string, object> context)
    {
        var bytes = document.ToBytes();
        var result = Convert(PdfReadDocument.Open(bytes), context);
        return new(result.Info, Source("pdf", bytes, context), result.Targets.ToList());
    }

    static ConversionResult Convert(PdfReadDocument document, IReadOnlyDictionary<string, object> context)
    {
        var metadata = document.Metadata;
        var pages = document.Pages;
        var info = new
        {
            metadata.Title,
            metadata.Author,
            metadata.Subject,
            metadata.Keywords,
            PageCount = pages.Count
        };

        var derived = Pages(
            context,
            pages.Count,
            _ => pages[_].ExtractText(),
            _ => pages[_].ExportImage(OfficeImageExportFormat.Png, Render(new PdfImageExportOptions())).Bytes);
        return new(info, derived);
    }
}
