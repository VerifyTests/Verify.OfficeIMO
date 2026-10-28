namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterEpub()
    {
        VerifierSettings.RegisterStreamConverter(
            "epub",
            (_, stream, context) => Convert(EpubDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<EpubDocument>(Convert);
        VerifierSettings.RegisterFileConverter<EpubPublication>(Convert);
    }

    // A publication is what is authored, so it is saved as the epub, then read back for the rest
    static ConversionResult Convert(EpubPublication publication, IReadOnlyDictionary<string, object> context)
    {
        using var stream = new MemoryStream();
        // dcterms:modified is otherwise the time of saving, so would differ every run. Pinned, as
        // DeterministicPackage strips the modified date of an Office document.
        publication.Save(
            stream,
            new()
            {
                ModifiedAt = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero)
            });
        var result = Convert(publication.Read(), context);
        return new(result.Info, DeterministicSource("epub", stream.ToArray(), context), result.Targets.ToList());
    }

    // A chapter is a page of the book, so PagesToInclude limits the chapters
    static ConversionResult Convert(EpubDocument document, IReadOnlyDictionary<string, object> context)
    {
        var chapters = document.Chapters;
        var info = new
        {
            document.Title,
            document.Language,
            document.Creator,
            Chapters = chapters.Select(_ => _.Title ?? _.Path).ToList()
        };

        var derived = Pages(
            context,
            chapters.Count,
            _ => chapters[_].Text,
            Rendered(() => document.ExportImages(OfficeImageExportFormat.Png, Render(new EpubImageExportOptions()))),
            "Chapter");
        return new(info, derived);
    }
}
