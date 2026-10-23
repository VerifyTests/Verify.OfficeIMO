using OfficeIMO.Rtf;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterRtf()
    {
        VerifierSettings.RegisterStreamConverter(
            "rtf",
            (_, stream, context) => Convert(RtfDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<RtfDocument>(Convert);
    }

    // OfficeIMO does not render rtf itself (only via its pdf package), so there is no png
    static ConversionResult Convert(RtfDocument document, IReadOnlyDictionary<string, object> context)
    {
        var properties = document.Info;
        var paragraphs = document.Paragraphs;
        var info = new
        {
            properties.Title,
            properties.Subject,
            properties.Author,
            properties.Manager,
            properties.Company,
            properties.Category,
            properties.Keywords,
            properties.Comments,
            Paragraphs = paragraphs.Count
        };

        List<Target> derived = [];
        if (!context.IsDerivedTargetExcluded("txt"))
        {
            derived.Add(new("txt", string.Join("\n", paragraphs.Select(_ => _.ToPlainText())), "Text"));
        }

        return new(info, Source("rtf", document.ToBytes(), context), derived);
    }
}
