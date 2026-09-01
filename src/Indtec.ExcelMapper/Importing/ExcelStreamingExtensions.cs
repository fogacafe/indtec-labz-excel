using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Indtec.ExcelMapper.Conversion;
using Indtec.ExcelMapper.Internal;
using Indtec.ExcelMapper.Mapping;

namespace Indtec.ExcelMapper.Importing;

public static class ExcelStreamingExtensions
{
    public static async Task ImportChunksAsync<T>(
        this ExcelMapper mapper,
        Stream stream,
        Func<ExcelImportChunk<T>, CancellationToken, Task> onChunk,
        int chunkSize = 1000,
        Action<ExcelImportOptions<T>>? configure = null,
        CancellationToken cancellationToken = default) where T : new()
    {
        if (mapper is null) throw new ArgumentNullException(nameof(mapper));
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (onChunk is null) throw new ArgumentNullException(nameof(onChunk));
        if (chunkSize < 1) throw new ArgumentOutOfRangeException(nameof(chunkSize), "Chunk size must be greater than zero.");
        if (!stream.CanRead) throw new ArgumentException("The stream must be readable.", nameof(stream));
        if (!stream.CanSeek) throw new ArgumentException("Streaming .xlsx import requires a seekable stream.", nameof(stream));

        var options = new ExcelImportOptions<T>();
        configure?.Invoke(options);

        if (options.BatchValidators.Count > 0)
            throw new ExcelMappingException(mapper.Messages.StreamingBatchValidatorsNotSupported());

        var map = GetMap<T>();
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart
            ?? throw new ExcelMappingException(mapper.Messages.WorksheetNotFound(map.SheetName));

        var sheet = workbookPart.Workbook.Sheets?
            .Elements<Sheet>()
            .FirstOrDefault(x => string.Equals(x.Name?.Value, map.SheetName, StringComparison.OrdinalIgnoreCase));

        if (sheet?.Id?.Value is not { } relationshipId)
            throw new ExcelMappingException(mapper.Messages.WorksheetNotFound(map.SheetName));

        var worksheetPart = workbookPart.GetPartById(relationshipId) as WorksheetPart
            ?? throw new ExcelMappingException(mapper.Messages.WorksheetNotFound(map.SheetName));

        var sharedStrings = ReadSharedStrings(workbookPart);
        Dictionary<string, int>? headers = null;
        var chunkRows = new List<(int RowNumber, T Value, IReadOnlyList<ExcelImportError> Errors)>(chunkSize);
        var chunkIndex = 1;

        using var reader = OpenXmlReader.Create(worksheetPart);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (reader.ElementType != typeof(Row))
                continue;

            var row = (Row)reader.LoadCurrentElement();
            var rowNumber = checked((int)(row.RowIndex?.Value ?? 0U));
            if (rowNumber == 0)
                continue;

            var cells = row.Elements<Cell>()
                .Where(HasValue)
                .ToDictionary(
                    cell => GetColumnIndex(cell.CellReference?.Value),
                    cell => cell);

            if (rowNumber == 1)
            {
                headers = BuildHeaders(cells, sharedStrings);
                ValidateRequiredHeaders(map, headers, mapper);
                continue;
            }

            if (cells.Count == 0)
                continue;

            headers ??= new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (headers.Count == 0)
                ValidateRequiredHeaders(map, headers, mapper);

            var item = new T();
            var errors = new List<ExcelImportError>();

            foreach (var column in map.Columns)
            {
                if (!headers.TryGetValue(column.Header, out var columnNumber))
                    continue;

                if (column.Setter is null)
                    throw new ExcelMappingException(mapper.Messages.ReadOnlyProperty(column.Header));

                cells.TryGetValue(columnNumber, out var cell);
                var address = cell?.CellReference?.Value ?? $"{GetColumnName(columnNumber)}{rowNumber}";

                try
                {
                    var value = cell is null
                        ? new ExcelValue(null)
                        : ReadOpenXmlValue(cell, sharedStrings);

                    var converted = column.Converter is null
                        ? ExcelCellConverter.Read(value, column.ValueType, address, mapper.Messages)
                        : column.Converter.Read(value, column.ValueType);

                    column.Setter(item!, converted);
                }
                catch (Exception ex)
                {
                    if (options.ErrorBehavior == ExcelImportErrorBehavior.Throw)
                        throw;

                    errors.Add(new ExcelImportError(rowNumber, column.Header, ex.Message));
                }
            }

            if (errors.Count == 0)
            {
                var validationErrors = options.Validators
                    .Where(rule => !rule.Predicate(item))
                    .Select(rule => new ExcelImportError(rowNumber, null, rule.Message))
                    .ToArray();

                if (validationErrors.Length > 0 && options.ErrorBehavior == ExcelImportErrorBehavior.Throw)
                    throw new ExcelMappingException(validationErrors[0].ToString());

                errors.AddRange(validationErrors);
            }

            chunkRows.Add((rowNumber, item, errors.ToArray()));

            if (chunkRows.Count >= chunkSize)
            {
                await DeliverChunkAsync(chunkRows, chunkIndex++, onChunk, cancellationToken).ConfigureAwait(false);
                chunkRows.Clear();
            }
        }

        if (headers is null)
        {
            headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            ValidateRequiredHeaders(map, headers, mapper);
        }

        if (chunkRows.Count > 0)
            await DeliverChunkAsync(chunkRows, chunkIndex, onChunk, cancellationToken).ConfigureAwait(false);
    }

    private static async Task DeliverChunkAsync<T>(
        IReadOnlyList<(int RowNumber, T Value, IReadOnlyList<ExcelImportError> Errors)> source,
        int index,
        Func<ExcelImportChunk<T>, CancellationToken, Task> onChunk,
        CancellationToken cancellationToken)
    {
        var rows = source
            .Select(x => new ExcelImportRow<T>(x.RowNumber, x.Value, x.Errors))
            .ToArray();

        var errors = rows.SelectMany(x => x.Errors).ToArray();
        var items = rows.Where(x => !x.HasErrors).Select(x => x.Value).ToArray();
        var chunk = new ExcelImportChunk<T>(
            index,
            rows[0].RowNumber,
            rows[rows.Length - 1].RowNumber,
            items,
            errors,
            rows);

        await onChunk(chunk, cancellationToken).ConfigureAwait(false);
    }

    private static Dictionary<string, int> BuildHeaders(
        IReadOnlyDictionary<int, Cell> cells,
        IReadOnlyList<string> sharedStrings)
    {
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in cells)
        {
            var header = ReadOpenXmlValue(pair.Value, sharedStrings).AsString();
            if (!string.IsNullOrWhiteSpace(header))
                headers[header] = pair.Key;
        }

        return headers;
    }

    private static void ValidateRequiredHeaders<T>(
        ExcelTypeMap map,
        IReadOnlyDictionary<string, int> headers,
        ExcelMapper mapper)
    {
        var missing = map.Columns
            .Where(x => x.Required && !headers.ContainsKey(x.Header))
            .Select(x => x.Header)
            .ToArray();

        if (missing.Length > 0)
            throw new ExcelMappingException(mapper.Messages.RequiredColumnsNotFound(map.SheetName, missing));
    }

    private static IReadOnlyList<string> ReadSharedStrings(WorkbookPart workbookPart)
        => workbookPart.SharedStringTablePart?.SharedStringTable?
            .Elements<SharedStringItem>()
            .Select(x => x.InnerText)
            .ToArray()
            ?? Array.Empty<string>();

    private static ExcelValue ReadOpenXmlValue(Cell cell, IReadOnlyList<string> sharedStrings)
    {
        var text = cell.CellValue?.Text;
        var type = cell.DataType?.Value;

        if (type == CellValues.SharedString)
        {
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                   && index >= 0
                   && index < sharedStrings.Count
                ? new ExcelValue(sharedStrings[index])
                : new ExcelValue(text);
        }

        if (type == CellValues.Boolean)
            return new ExcelValue(text == "1" || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase));

        if (type == CellValues.InlineString)
            return new ExcelValue(cell.InlineString?.InnerText ?? string.Empty);

        if (type == CellValues.String || type == CellValues.Error)
            return new ExcelValue(text);

        if (type == CellValues.Date)
        {
            return DateTime.TryParse(
                text,
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out var date)
                ? new ExcelValue(date)
                : new ExcelValue(text);
        }

        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return new ExcelValue(number);

        return new ExcelValue(text);
    }

    private static bool HasValue(Cell cell)
        => cell.CellValue is not null || cell.InlineString is not null;

    private static int GetColumnIndex(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            throw new ExcelMappingException("A streamed cell is missing its Excel reference.");

        var result = 0;
        foreach (var character in reference)
        {
            if (!char.IsLetter(character))
                break;

            result = checked(result * 26 + char.ToUpperInvariant(character) - 'A' + 1);
        }

        if (result == 0)
            throw new ExcelMappingException($"Invalid Excel cell reference '{reference}'.");

        return result;
    }

    private static string GetColumnName(int columnNumber)
    {
        var name = string.Empty;
        while (columnNumber > 0)
        {
            columnNumber--;
            name = (char)('A' + columnNumber % 26) + name;
            columnNumber /= 26;
        }

        return name;
    }

    private static ExcelTypeMap GetMap<T>() where T : new()
    {
        var probe = new T();
        return probe is IGeneratedExcelMapProvider provider
            ? provider.ExcelMap
            : throw new GeneratedMapNotFoundException(typeof(T));
    }
}
