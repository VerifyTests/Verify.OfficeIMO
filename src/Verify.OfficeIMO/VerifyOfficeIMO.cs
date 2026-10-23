using WordParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace VerifyTests;

public static class VerifyOfficeIMO
{
    public static bool Initialized { get; private set; }

    public static void Initialize()
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;

        VerifierSettings.RegisterStreamConverter(
            "docx",
            (_, stream, context) => ConvertLoaded(WordDocument.Load(stream), context));
        VerifierSettings.RegisterStreamConverter(
            "xlsx",
            (_, stream, context) => ConvertLoaded(ExcelDocument.Load(stream), context));
        VerifierSettings.RegisterStreamConverter(
            "pptx",
            (_, stream, context) => ConvertLoaded(PowerPointPresentation.Load(stream), context));
        VerifierSettings.RegisterFileConverter<WordDocument>(Convert);
        VerifierSettings.RegisterFileConverter<ExcelDocument>(Convert);
        VerifierSettings.RegisterFileConverter<PowerPointPresentation>(Convert);
    }

    // A document loaded here belongs to the conversion, so it is disposed once Verify is done with it.
    // One the test passed in belongs to the test.
    static ConversionResult ConvertLoaded(WordDocument document, IReadOnlyDictionary<string, object> context) =>
        Convert(document, context, Dispose(document));

    static ConversionResult ConvertLoaded(ExcelDocument document, IReadOnlyDictionary<string, object> context) =>
        Convert(document, context, Dispose(document));

    static ConversionResult ConvertLoaded(PowerPointPresentation document, IReadOnlyDictionary<string, object> context) =>
        Convert(document, context, Dispose(document));

    static Func<Task> Dispose(IDisposable document) =>
        () =>
        {
            document.Dispose();
            return Task.CompletedTask;
        };

    static ConversionResult Convert(WordDocument document, IReadOnlyDictionary<string, object> context) =>
        Convert(document, context, null);

    static ConversionResult Convert(WordDocument document, IReadOnlyDictionary<string, object> context, Func<Task>? cleanup)
    {
        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;

        // read back from what was saved, so the text is of the same document as the docx
        List<string> paragraphs;
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            paragraphs = package.MainDocumentPart?.Document?.Body?
                             .Descendants<WordParagraph>()
                             .Select(_ => _.InnerText)
                             .ToList() ??
                         [];
        }

        var properties = document.BuiltinDocumentProperties;
        var info = new WordInfo
        {
            Title = properties.Title,
            Subject = properties.Subject,
            Description = properties.Description,
            Category = properties.Category,
            Keywords = properties.Keywords,
            Revision = properties.Revision,
            Version = properties.Version,
            Sections = document.Sections.Count,
            Paragraphs = paragraphs.Count,
            Tables = document.TablesIncludingNestedTables.Count
        };

        List<Target> derived = [];
        if (!context.IsDerivedTargetExcluded("txt"))
        {
            derived.Add(new("txt", string.Join('\n', paragraphs), "Text"));
        }

        // Word renders every page in one pass, so the pages not wanted are dropped after
        if (!context.IsDerivedTargetExcluded("png"))
        {
            var pages = document.ExportImages(OfficeImageExportFormat.Png);
            for (var index = 0; index < pages.Count; index++)
            {
                var page = index + 1;
                if (context.IsPageIncluded(page))
                {
                    derived.Add(Png(pages[index], $"Page{page}"));
                }
            }
        }

        return new(info, Source("docx", stream, context), derived, cleanup);
    }

    static ConversionResult Convert(ExcelDocument document, IReadOnlyDictionary<string, object> context) =>
        Convert(document, context, null);

    static ConversionResult Convert(ExcelDocument document, IReadOnlyDictionary<string, object> context, Func<Task>? cleanup)
    {
        using var stream = new MemoryStream();
        document.Save(stream);

        var properties = document.BuiltinDocumentProperties;
        var info = new ExcelInfo
        {
            Title = properties.Title,
            Subject = properties.Subject,
            Description = properties.Description,
            Category = properties.Category,
            Keywords = properties.Keywords,
            Revision = properties.Revision,
            Version = properties.Version,
            Sheets = document.Sheets.Select(_ => _.Name).ToList()
        };

        // A sheet is a page, numbered in tab order, so PagesToInclude limits the csv files as it
        // limits the pages of any other document. Named for the sheet even when it is the only one,
        // so that a second sheet adds a file rather than renaming the first.
        List<Target> derived = [];
        var includeCsv = !context.IsDerivedTargetExcluded("csv");
        var includePng = !context.IsDerivedTargetExcluded("png");
        var page = 0;
        foreach (var sheet in document.Sheets)
        {
            page++;
            if (!context.IsPageIncluded(page))
            {
                continue;
            }

            if (includeCsv)
            {
                derived.Add(new("csv", ToCsv(sheet), sheet.Name));
            }

            if (includePng)
            {
                derived.Add(Png(sheet.ExportImage(OfficeImageExportFormat.Png), sheet.Name));
            }
        }

        return new(info, Source("xlsx", stream, context), derived, cleanup);
    }

    static ConversionResult Convert(PowerPointPresentation presentation, IReadOnlyDictionary<string, object> context) =>
        Convert(presentation, context, null);

    static ConversionResult Convert(PowerPointPresentation presentation, IReadOnlyDictionary<string, object> context, Func<Task>? cleanup)
    {
        using var stream = new MemoryStream(presentation.ToBytes());

        var slides = presentation.Slides;
        var hidden = slides
            .Select((slide, index) => (slide, index))
            .Where(_ => _.slide.Hidden)
            .Select(_ => _.index + 1)
            .ToList();

        var properties = presentation.BuiltinDocumentProperties;
        var info = new PowerPointInfo
        {
            Title = properties.Title,
            Subject = properties.Subject,
            Description = properties.Description,
            Category = properties.Category,
            Keywords = properties.Keywords,
            Revision = properties.Revision,
            Version = properties.Version,
            SlideCount = slides.Count,
            HiddenSlides = hidden.Count == 0 ? null : hidden
        };

        List<Target> derived = [];
        var includeText = !context.IsDerivedTargetExcluded("txt");
        var includePng = !context.IsDerivedTargetExcluded("png");
        for (var index = 0; index < slides.Count; index++)
        {
            var page = index + 1;
            if (!context.IsPageIncluded(page))
            {
                continue;
            }

            var slide = slides[index];
            if (includeText)
            {
                derived.Add(new("txt", ToText(slide), $"Slide{page}"));
            }

            if (includePng)
            {
                derived.Add(Png(slide.ExportImage(OfficeImageExportFormat.Png), $"Slide{page}"));
            }
        }

        return new(info, Source("pptx", stream, context), derived, cleanup);
    }

    // the document is not made deterministic when ExcludeTargets says it is not wanted
    static Target? Source(string extension, MemoryStream stream, IReadOnlyDictionary<string, object> context)
    {
        if (context.IsTargetExcluded(extension))
        {
            return null;
        }

        stream.Position = 0;
        return new(extension, DeterministicPackage.Convert(stream));
    }

    static Target Png(OfficeImageExportResult image, string name) =>
        new("png", new MemoryStream(image.Bytes), name);

    static string ToCsv(ExcelSheet sheet)
    {
        var builder = new StringBuilder();
        if (!A1.TryParseRange(sheet.UsedRangeA1, out var firstRow, out var firstColumn, out var lastRow, out var lastColumn))
        {
            return string.Empty;
        }

        for (var row = firstRow; row <= lastRow; row++)
        {
            for (var column = firstColumn; column <= lastColumn; column++)
            {
                if (column > firstColumn)
                {
                    builder.Append(',');
                }

                if (sheet.TryGetCellText(row, column, out var text))
                {
                    builder.Append(Csv.Escape(text));
                }
            }

            builder.AppendLineN();
        }

        return builder.ToString();
    }

    static string ToText(PowerPointSlide slide)
    {
        var builder = new StringBuilder();
        foreach (var textBox in slide.TextBoxes)
        {
            builder.AppendLineN(textBox.Text);
        }

        if (slide.HasSpeakerNotes)
        {
            var notes = slide.Notes.Text;
            if (notes.Length > 0)
            {
                builder.AppendLineN();
                builder.AppendLineN("Notes:");
                builder.AppendLineN(notes);
            }
        }

        return builder.ToString();
    }
}
