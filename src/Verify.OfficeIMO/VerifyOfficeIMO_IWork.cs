using OfficeIMO.IWork;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterIWork()
    {
        VerifierSettings.RegisterStreamConverter(
            "pages",
            (_, stream, context) => Convert(IWorkSourceDocument.Open(stream), context));
        VerifierSettings.RegisterStreamConverter(
            "numbers",
            (_, stream, context) => Convert(IWorkSourceDocument.Open(stream), context));
        VerifierSettings.RegisterStreamConverter(
            "key",
            (_, stream, context) => Convert(IWorkSourceDocument.Open(stream), context));
        VerifierSettings.RegisterFileConverter<IWorkSourceDocument>(Convert);
    }

    // OfficeIMO reads iWork documents but writes them only from its own authoring model, so there is
    // no document to verify, only what was read
    static ConversionResult Convert(IWorkSourceDocument document, IReadOnlyDictionary<string, object> context)
    {
        switch (document.Kind)
        {
            case IWorkDocumentKind.Pages:
                var pages = document.ReadPages();
                List<Target> text = [];
                if (!context.IsDerivedTargetExcluded("txt"))
                {
                    text.Add(new("txt", string.Join("\n", pages.Paragraphs), "Text"));
                }

                text.AddRange(Tables(pages.Tables, context, ""));
                return new(
                    new
                    {
                        Kind = "Pages",
                        Paragraphs = pages.Paragraphs.Count,
                        Tables = pages.Tables.Count
                    },
                    text);

            case IWorkDocumentKind.Numbers:
                var numbers = document.ReadNumbers();
                List<Target> sheets = [];
                foreach (var sheet in numbers.Sheets)
                {
                    sheets.AddRange(Tables(sheet.Tables, context, $"{sheet.Name}_"));
                }

                return new(
                    new
                    {
                        Kind = "Numbers",
                        Sheets = numbers.Sheets.Select(_ => _.Name).ToList()
                    },
                    sheets);

            default:
                var keynote = document.ReadKeynote();
                var slides = keynote.Slides;
                var skipped = slides
                    .Where(_ => _.IsSkipped)
                    .Select(_ => _.Index)
                    .ToList();
                return new(
                    new
                    {
                        Kind = "Keynote",
                        SlideCount = slides.Count,
                        SkippedSlides = skipped.Count == 0 ? null : skipped
                    },
                    Pages(context, slides.Count, _ => ToText(slides[_]), null, "Slide"));
        }
    }

    static IEnumerable<Target> Tables(IReadOnlyList<IWorkTable> tables, IReadOnlyDictionary<string, object> context, string prefix)
    {
        if (context.IsDerivedTargetExcluded("csv"))
        {
            yield break;
        }

        for (var index = 0; index < tables.Count; index++)
        {
            yield return new("csv", ToCsv(tables[index]), $"{prefix}Table{index + 1}");
        }
    }

    static string ToCsv(IWorkTable table)
    {
        var cells = table.Cells.ToDictionary(_ => (_.Row, _.Column), _ => _.DisplayText);
        var builder = new StringBuilder();
        for (var row = 0; row < table.RowCount; row++)
        {
            var values = Enumerable.Range(0, table.ColumnCount)
                .Select(column => Csv.Escape(cells.TryGetValue((row, column), out var text) ? text : ""));
            builder.AppendLineN(string.Join(",", values));
        }

        return builder.ToString();
    }

    static string ToText(IWorkKeynoteSlide slide)
    {
        var builder = new StringBuilder();
        if (slide.Title.Length > 0)
        {
            builder.AppendLineN(slide.Title);
        }

        foreach (var line in slide.Body)
        {
            builder.AppendLineN(line);
        }

        if (slide.PresenterNotes.Length > 0)
        {
            builder.AppendLineN();
            builder.AppendLineN("Notes:");
            builder.AppendLineN(slide.PresenterNotes);
        }

        return builder.ToString();
    }
}
