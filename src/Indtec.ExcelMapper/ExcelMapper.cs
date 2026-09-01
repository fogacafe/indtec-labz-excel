using ClosedXML.Excel;
using Indtec.ExcelMapper.Internal;
using Indtec.ExcelMapper.Mapping;
using Indtec.ExcelMapper.Styling;

namespace Indtec.ExcelMapper;

public sealed class ExcelMapper
{
    public IReadOnlyList<T> Import<T>(Stream stream) where T : new()
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));

        var map = GetMap<T>();
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.FirstOrDefault(x =>
            x.Name.Equals(map.SheetName, StringComparison.OrdinalIgnoreCase));

        if (worksheet is null)
            throw new ExcelMappingException($"Worksheet '{map.SheetName}' was not found.");

        var headers = worksheet.Row(1)
            .CellsUsed()
            .ToDictionary(x => x.GetString(), x => x.Address.ColumnNumber, StringComparer.OrdinalIgnoreCase);

        var result = new List<T>();

        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            var item = new T();

            foreach (var column in map.Columns)
            {
                if (!headers.TryGetValue(column.Header, out var columnNumber))
                    continue;

                if (column.Setter is null)
                    throw new ExcelMappingException(
                        $"Property mapped to '{column.Header}' is read-only and cannot be imported.");

                var value = ExcelCellConverter.Read(row.Cell(columnNumber), column.ValueType);
                column.Setter(item!, value);
            }

            result.Add(item);
        }

        return result;
    }

    public void Export<T>(IEnumerable<T> items, Stream stream) where T : new()
        => Export(items, stream, null);

    public void Export<T>(
        IEnumerable<T> items,
        Stream stream,
        Action<ExcelExportOptions<T>>? configure) where T : new()
    {
        if (items is null) throw new ArgumentNullException(nameof(items));
        if (stream is null) throw new ArgumentNullException(nameof(stream));

        var map = GetMap<T>();
        var options = new ExcelExportOptions<T>();
        configure?.Invoke(options);

        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(map.SheetName);

        for (var i = 0; i < map.Columns.Count; i++)
        {
            var cell = worksheet.Cell(1, i + 1);
            cell.Value = map.Columns[i].Header;
            ClosedXmlStyleApplier.Apply(cell.Style, options.HeaderStyle);

            if (options.Columns.TryGetValue(map.Columns[i].PropertyName, out var columnConfig) &&
                columnConfig.Width.HasValue)
            {
                worksheet.Column(i + 1).Width = columnConfig.Width.Value;
            }
        }

        var rowNumber = 2;
        foreach (var item in items)
        {
            var rowRules = options.RowRules.Where(x => x.Predicate(item)).ToArray();

            for (var i = 0; i < map.Columns.Count; i++)
            {
                var column = map.Columns[i];
                var cell = worksheet.Cell(rowNumber, i + 1);
                var value = column.Getter(item!);
                ExcelCellConverter.Write(cell, value);

                foreach (var rule in rowRules)
                    ClosedXmlStyleApplier.Apply(cell.Style, rule.Style);

                if (options.Columns.TryGetValue(column.PropertyName, out var columnConfig))
                {
                    ClosedXmlStyleApplier.Apply(cell.Style, columnConfig.Style);

                    foreach (var rule in columnConfig.Rules)
                    {
                        if (rule.Predicate(item))
                            ClosedXmlStyleApplier.Apply(cell.Style, rule.Style);
                    }
                }
            }

            rowNumber++;
        }

        if (options.FreezeHeader)
            worksheet.SheetView.FreezeRows(1);

        if (options.AutoFilter && worksheet.RangeUsed() is { } range)
            range.SetAutoFilter();

        workbook.SaveAs(stream);
    }

    public void Export<T>(IEnumerable<T> items, string path) where T : new()
        => Export(items, path, null);

    public void Export<T>(
        IEnumerable<T> items,
        string path,
        Action<ExcelExportOptions<T>>? configure) where T : new()
    {
        if (path is null) throw new ArgumentNullException(nameof(path));
        using var stream = File.Create(path);
        Export(items, stream, configure);
    }

    private static ExcelTypeMap GetMap<T>() where T : new()
    {
        var probe = new T();
        return probe is IGeneratedExcelMapProvider provider
            ? provider.ExcelMap
            : throw new GeneratedMapNotFoundException(typeof(T));
    }
}
