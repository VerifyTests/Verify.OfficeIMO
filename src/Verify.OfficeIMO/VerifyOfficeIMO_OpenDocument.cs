namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterOpenDocument()
    {
        foreach (var extension in new[] { "odt", "ods", "odp", "odg" })
        {
            VerifierSettings.RegisterStreamConverter(
                extension,
                (_, stream, context) => Convert(OdfDocument.Load(stream), context));
        }

        VerifierSettings.RegisterFileConverter<OdfDocument>(Convert);
    }

    // OfficeIMO does not render OpenDocument itself (only via its pdf packages), so there is no png
    static ConversionResult Convert(OdfDocument document, IReadOnlyDictionary<string, object> context)
    {
        var metadata = document.Metadata;
        List<Target> derived = [];
        string extension;
        object? content;
        switch (document)
        {
            case OdtDocument text:
                extension = "odt";
                content = new
                {
                    Paragraphs = text.Paragraphs.Count
                };
                if (!context.IsDerivedTargetExcluded("txt"))
                {
                    derived.Add(new("txt", string.Join('\n', text.Paragraphs.Select(_ => _.Text)), "Text"));
                }

                break;
            case OdsDocument spreadsheet:
                extension = "ods";
                content = new
                {
                    Sheets = spreadsheet.Sheets.Select(_ => _.Name).ToList()
                };
                if (!context.IsDerivedTargetExcluded("csv"))
                {
                    var page = 0;
                    foreach (var sheet in spreadsheet.Sheets)
                    {
                        page++;
                        if (context.IsPageIncluded(page))
                        {
                            derived.Add(new("csv", ToCsv(sheet), sheet.Name));
                        }
                    }
                }

                break;
            case OdpPresentation presentation:
                extension = "odp";
                var slides = presentation.Slides;
                content = new
                {
                    SlideCount = slides.Count
                };
                derived.AddRange(Pages(context, slides.Count, _ => ToText(slides[_]), null, "Slide"));
                break;
            default:
                extension = "odg";
                content = null;
                break;
        }

        var info = new
        {
            Title = NullIfEmpty(metadata.Title),
            Subject = NullIfEmpty(metadata.Subject),
            Description = NullIfEmpty(metadata.Description),
            Content = content
        };

        return new(info, Source(extension, document.ToBytes(), context), derived);
    }

    static string ToCsv(OdsSheet sheet)
    {
        var builder = new StringBuilder();
        if (sheet.UsedRange is not { } range)
        {
            return string.Empty;
        }

        for (var row = range.FirstRow; row <= range.LastRow; row++)
        {
            for (var column = range.FirstColumn; column <= range.LastColumn; column++)
            {
                if (column > range.FirstColumn)
                {
                    builder.Append(',');
                }

                builder.Append(Csv.Escape(sheet.Cell(row, column).Text));
            }

            builder.AppendLineN();
        }

        return builder.ToString();
    }

    static string ToText(OdpSlide slide)
    {
        var builder = new StringBuilder();
        foreach (var paragraph in slide.Shapes.OfType<OdpTextBox>().SelectMany(_ => _.Paragraphs))
        {
            builder.AppendLineN(paragraph.Text);
        }

        var notes = slide.SpeakerNotes?.Paragraphs ?? [];
        if (notes.Count > 0)
        {
            builder.AppendLineN();
            builder.AppendLineN("Notes:");
            foreach (var paragraph in notes)
            {
                builder.AppendLineN(paragraph.Text);
            }
        }

        return builder.ToString();
    }
}
