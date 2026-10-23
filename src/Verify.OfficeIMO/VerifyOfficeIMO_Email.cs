using OfficeIMO.Email;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterEmail()
    {
        VerifierSettings.RegisterStreamConverter(
            "eml",
            (_, stream, context) => Convert(EmailDocument.Load(stream), context));
        VerifierSettings.RegisterStreamConverter(
            "msg",
            (_, stream, context) => Convert(EmailDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<EmailDocument>(Convert);
    }

    // verified as eml, whatever it was loaded from, since eml is text and so diffs
    static ConversionResult Convert(EmailDocument document, IReadOnlyDictionary<string, object> context)
    {
        var info = new
        {
            document.Subject,
            From = Address(document.From),
            Recipients = document.Recipients
                .Select(_ => $"{_.Kind}: {Address(_.Address)}")
                .ToList(),
            document.Date,
            Attachments = document.Attachments.Select(_ => _.FileName).ToList()
        };

        List<Target> derived = [];
        var text = document.Body.Text;
        if (text != null &&
            !context.IsDerivedTargetExcluded("txt"))
        {
            derived.Add(new("txt", text, "Body"));
        }

        if (document.Body.Html != null &&
            !context.IsDerivedTargetExcluded("html"))
        {
            derived.Add(new("html", document.Body.Html, "Body"));
        }

        if (!context.IsDerivedTargetExcluded("png"))
        {
            derived.Add(new("png", new MemoryStream(document.ExportImage(OfficeImageExportFormat.Png).Bytes), "Body"));
        }

        var source = Source("eml", document.ToBytes(), context);
        return new(info, source, derived);
    }

    static string? Address(EmailAddress? address)
    {
        if (address == null)
        {
            return null;
        }

        if (string.IsNullOrEmpty(address.DisplayName))
        {
            return address.Address;
        }

        return $"{address.DisplayName} <{address.Address}>";
    }
}
