using OfficeIMO.OneNote;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterOneNote()
    {
        VerifierSettings.RegisterStreamConverter(
            "one",
            (_, stream, context) => Convert(OneNoteSectionReader.Read(stream), context));
        VerifierSettings.RegisterFileConverter<OneNoteSection>(Convert);
    }

    // A section is written with fresh ids each time, so the .one file is not verified, only its content
    static ConversionResult Convert(OneNoteSection section, IReadOnlyDictionary<string, object> context)
    {
        var pages = section.Pages;
        var info = new
        {
            section.Name,
            Pages = pages.Select(_ => _.Title).ToList()
        };

        var derived = Pages(
            context,
            pages.Count,
            _ => ToText(pages[_]),
            Rendered(() => section.ExportImages(OfficeImageExportFormat.Png)));
        return new(info, derived);
    }

    static string ToText(OneNotePage page)
    {
        var builder = new StringBuilder();
        foreach (var outline in page.Outlines)
        {
            AppendText(builder, outline.Children, 0);
        }

        AppendText(builder, page.DirectContent, 0);
        return builder.ToString();
    }

    static void AppendText(StringBuilder builder, IEnumerable<OneNoteElement> elements, int depth)
    {
        foreach (var element in elements)
        {
            if (element is not OneNoteParagraph paragraph)
            {
                continue;
            }

            builder.Append(' ', depth * 2);
            builder.AppendLineN(string.Concat(paragraph.Runs.Select(_ => _.Text)));
            AppendText(builder, paragraph.Children, depth + 1);
        }
    }
}
