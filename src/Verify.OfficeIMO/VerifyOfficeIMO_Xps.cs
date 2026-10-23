using OfficeIMO.Xps;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterXps()
    {
        VerifierSettings.RegisterStreamConverter(
            "xps",
            (_, stream, context) => Convert(XpsDocument.Load(stream), context));
        VerifierSettings.RegisterStreamConverter(
            "oxps",
            (_, stream, context) => Convert(XpsDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<XpsDocument>(Convert);
    }

    static ConversionResult Convert(XpsDocument document, IReadOnlyDictionary<string, object> context)
    {
        var pages = document.Pages;
        var info = new
        {
            Format = document.Format.ToString(),
            PageCount = pages.Count
        };

        var derived = Pages(
            context,
            pages.Count,
            _ => pages[_].ExtractText(),
            _ => pages[_].ExportImage(OfficeImageExportFormat.Png, Render(new OfficeImageExportOptions())).Bytes);
        return new(info, DeterministicSource("xps", document.Save(), context), derived);
    }
}
