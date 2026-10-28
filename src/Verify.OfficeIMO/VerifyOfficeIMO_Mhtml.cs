namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterMhtml()
    {
        VerifierSettings.RegisterStreamConverter(
            "mht",
            (_, stream, context) => Convert(MhtmlDocument.Load(stream), context));
        VerifierSettings.RegisterStreamConverter(
            "mhtml",
            (_, stream, context) => Convert(MhtmlDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<MhtmlDocument>(Convert);
    }

    static ConversionResult Convert(MhtmlDocument document, IReadOnlyDictionary<string, object> context)
    {
        var info = new
        {
            document.Subject,
            document.ContentLocation,
            Resources = document.Resources.Select(_ => _.ContentLocation).ToList()
        };

        List<Target> derived = [];
        if (!context.IsDerivedTargetExcluded("html"))
        {
            derived.Add(new("html", document.Html, "Page"));
        }

        if (!context.IsDerivedTargetExcluded("png"))
        {
            var options = Render(new HtmlRenderOptions());
            document.ConfigureRenderOptions(options);
            var image = document.HtmlDocument.ExportImage(OfficeImageExportFormat.Png, options);
            derived.Add(new("png", new MemoryStream(image.Bytes), "Page"));
        }

        return new(info, Source("mht", document.ToBytes(), context), derived);
    }
}
