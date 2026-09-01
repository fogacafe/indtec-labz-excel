using ClosedXML.Excel;
using Indtec.ExcelMapper.Internal;
using Indtec.ExcelMapper.Mapping;

namespace Indtec.ExcelMapper;

public sealed class ExcelMapper
{
    public IReadOnlyList<T> Import<T>(Stream stream) where T : new()
    {
        ArgumentNullException.ThrowIfNull(stream);

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
    {
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(stream);

        var map = GetMap<T>();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(map.SheetName);

        for (var i = 0; i < map.Columns.Count; i++)
            worksheet.Cell(1, i + 1).Value = map.Columns[i].Header;

        var rowNumber = 2;
        foreach (var item in items)
        {
            for (var i = 0; i < map.Columns.Count; i++)
            {
                var value = map.Columns[i].Getter(item!);
                ExcelCellConverter.Write(worksheet.Cell(rowNumber, i + 1), value);
            }

            rowNumber++;
        }

        workbook.SaveAs(stream);
    }

    public void Export<T>(IEnumerable<T> items, string path) where T : new()
    {
        using var stream = File.Create(path);
        Export(items, stream);
    }

    private static ExcelTypeMap GetMap<T>() where T : new()
    {
        var probe = new T();
        return probe is IGeneratedExcelMapProvider provider
            ? provider.ExcelMap
            : throw new GeneratedMapNotFoundException(typeof(T));
    }
}
