using OfficeIMO.Publisher;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterPublisher()
    {
        VerifierSettings.RegisterStreamConverter(
            "pub",
            (_, stream, context) => Convert(PublisherDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<PublisherDocument>(Convert);
    }

    // OfficeIMO reads Publisher documents but does not write them, so there is no document to verify,
    // only what was read
    static ConversionResult Convert(PublisherDocument document, IReadOnlyDictionary<string, object> context)
    {
        var pages = document.Pages;
        var stories = document.TextStories.ToDictionary(_ => _.Id, _ => _.Text);
        var info = new
        {
            Pages = pages
                .Select(_ => new
                {
                    Name = NullIfEmpty(_.Name),
                    _.Width,
                    _.Height
                })
                .ToList(),
            Images = document.Images.Count
        };

        var derived = Pages(
            context,
            pages.Count,
            _ => string.Join(
                "\n",
                pages[_].TextFrames
                    .OrderBy(_ => _.Order)
                    .Select(_ => _.StoryId)
                    .Distinct()
                    .Select(_ => stories.TryGetValue(_, out var text) ? text : "")),
            _ => pages[_].Drawing.ExportImage(OfficeImageExportFormat.Png, Render(new OfficeImageExportOptions())).Bytes);
        return new(info, derived);
    }
}
