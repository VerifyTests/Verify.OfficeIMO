using OfficeIMO;
using OfficeIMO.Project;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterProject()
    {
        VerifierSettings.RegisterStreamConverter(
            "mpp",
            (_, stream, context) => ConvertLoaded(ProjectDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<ProjectDocument>(Convert);
    }

    static ConversionResult ConvertLoaded(ProjectDocument document, IReadOnlyDictionary<string, object> context) =>
        Convert(document, context, Dispose(document));

    static ConversionResult Convert(ProjectDocument document, IReadOnlyDictionary<string, object> context) =>
        Convert(document, context, null);

    // verified as Project xml, whatever it was loaded from, since xml is text and so diffs. A page of
    // the rendered schedule is a page of the document.
    static ConversionResult Convert(ProjectDocument document, IReadOnlyDictionary<string, object> context, Func<Task>? cleanup)
    {
        var info = new
        {
            document.Name,
            document.Title,
            document.Author,
            document.Subject,
            document.Company,
            document.Manager,
            TaskCount = document.AllTasks.Count()
        };

        List<Target> derived = [];
        if (!context.IsDerivedTargetExcluded("txt"))
        {
            var builder = new StringBuilder();
            AppendTasks(builder, document.Tasks, 0);
            derived.Add(new("txt", builder.ToString(), "Tasks"));
        }

        // a schedule can only be calculated, so rendered, against a calendar
        if (document.Calendar != null &&
            !context.IsDerivedTargetExcluded("png"))
        {
            var view = document.CreateView(document.CalculateSchedule());
            var pages = view.Render();
            derived.AddRange(
                Pages(
                    context,
                    pages.Count,
                    null,
                    _ => pages[_].Drawing.ExportImage(OfficeImageExportFormat.Png, Render(new OfficeImageExportOptions())).Bytes));
        }

        Target? source = null;
        if (!context.IsTargetExcluded("xml"))
        {
            // a native mpp holds views, filters, and other records that Project xml has no place for.
            // They are dropped rather than failing the verification.
            source = new(
                "xml",
                document.ToXml(
                    new()
                    {
                        LossPolicy = OfficeConversionLossPolicy.Allow
                    }));
        }

        return new(info, source, derived, cleanup);
    }

    static void AppendTasks(StringBuilder builder, IEnumerable<ProjectTask> tasks, int depth)
    {
        foreach (var task in tasks)
        {
            builder.Append(' ', depth * 2);
            builder.Append(task.Name);
            if (task.Start != null)
            {
                builder.Append($" ({DateFormatter.Convert(task.Start.Value)} - {DateFormatter.Convert(task.Finish!.Value)})");
            }

            builder.AppendLineN();
            AppendTasks(builder, task.Children, depth + 1);
        }
    }
}
