using OfficeIMO.Access;

namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    static void RegisterAccess()
    {
        VerifierSettings.RegisterStreamConverter(
            "accdb",
            (_, stream, context) => ConvertLoaded(AccessDocument.Load(stream), context));
        VerifierSettings.RegisterStreamConverter(
            "mdb",
            (_, stream, context) => ConvertLoaded(AccessDocument.Load(stream), context));
        VerifierSettings.RegisterFileConverter<AccessDocument>(Convert);
    }

    static ConversionResult ConvertLoaded(AccessDocument document, IReadOnlyDictionary<string, object> context) =>
        Convert(document, context, Dispose(document));

    static ConversionResult Convert(AccessDocument document, IReadOnlyDictionary<string, object> context) =>
        Convert(document, context, null);

    // A database file is not written deterministically, so it is described by its schema and a csv of
    // each table, rather than verified as a file.
    static ConversionResult Convert(AccessDocument document, IReadOnlyDictionary<string, object> context, Func<Task>? cleanup)
    {
        var tables = document.Tables
            .Where(_ => !_.IsSystem && !_.IsLinked)
            .ToList();

        var info = new
        {
            Format = document.Format.ToString(),
            Tables = tables
                .Select(_ => new
                {
                    _.Name,
                    _.RowCount,
                    Columns = _.Columns
                        .Select(_ => $"{_.Name}: {_.DataType}")
                        .ToList()
                })
                .ToList()
        };

        List<Target> derived = [];
        if (!context.IsDerivedTargetExcluded("csv"))
        {
            foreach (var table in tables)
            {
                derived.Add(new("csv", ToCsv(table), table.Name));
            }
        }

        return new(info, derived, cleanup);
    }

    static string ToCsv(AccessTable table)
    {
        var builder = new StringBuilder();
        using var reader = table.OpenDataReader();
        var fields = Enumerable.Range(0, reader.FieldCount).ToList();
        builder.AppendLineN(string.Join(",", fields.Select(_ => Csv.Escape(reader.GetName(_)))));
        while (reader.Read())
        {
            var values = fields.Select(_ => Csv.Escape(ToText(reader.GetValue(_))));
            builder.AppendLineN(string.Join(",", values));
        }

        return builder.ToString();
    }

    static string ToText(object value)
    {
        if (value is DateTime date)
        {
            return DateFormatter.Convert(date);
        }

        return System.Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
    }
}
