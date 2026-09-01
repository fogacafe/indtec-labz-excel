using ClosedXML.Excel;
using Indtec.ExcelMapper.Importing;
using Indtec.ExcelMapper.Internal;
using Indtec.ExcelMapper.Mapping;
using Indtec.ExcelMapper.Styling;

namespace Indtec.ExcelMapper;

public sealed class ExcelMapper
{
    public IReadOnlyList<T> Import<T>(Stream stream) where T : new()
        => Import<T>(stream, null).Items;

    public ExcelImportResult<T> Import<T>(
        Stream stream,
        Action<ExcelImportOptions<T>>? configure) where T : new()
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));

        var map = GetMap<T>();
        var options = new ExcelImportOptions<T>();
        configure?.Invoke(options);

        if (options.BatchValidators.Count > 0)
            throw new ExcelMappingException(
                "Batch validators require ImportAsync because they may perform asynchronous work.");

        using var workbook = new XLWorkbook(stream);
        var worksheet = GetWorksheet(workbook, map);
        var headers = GetHeaders(worksheet, map);

        var items = new List<T>();
        var errors = new List<ExcelImportError>();

        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            var item = new T();
            var rowHasMappingErrors = false;

            foreach (var column in map.Columns)
            {
                if (!headers.TryGetValue(column.Header, out var columnNumber))
                    continue;

                if (column.Setter is null)
                    throw new ExcelMappingException(
                        $"Property mapped to '{column.Header}' is read-only and cannot be imported.");

                try
                {
                    ReadCell(row, columnNumber, column, item);
                }
                catch (Exception ex) when (options.ErrorBehavior == ExcelImportErrorBehavior.Collect)
                {
                    rowHasMappingErrors = true;
                    errors.Add(new ExcelImportError(row.RowNumber(), column.Header, ex.Message));
                }
            }

            if (rowHasMappingErrors)
                continue;

            var validationErrors = options.Validators
                .Where(rule => !rule.Predicate(item))
                .Select(rule => new ExcelImportError(row.RowNumber(), null, rule.Message))
                .ToArray();

            if (validationErrors.Length > 0)
            {
                if (options.ErrorBehavior == ExcelImportErrorBehavior.Throw)
                    throw new ExcelMappingException(validationErrors[0].ToString());

                errors.AddRange(validationErrors);
                continue;
            }

            items.Add(item);
        }

        return new ExcelImportResult<T>(items, errors);
    }

    public async ValueTask<ExcelImportResult<T>> ImportAsync<T>(
        Stream stream,
        Action<ExcelImportOptions<T>>? configure = null,
        CancellationToken cancellationToken = default) where T : new()
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));

        var map = GetMap<T>();
        var options = new ExcelImportOptions<T>();
        configure?.Invoke(options);

        using var workbook = new XLWorkbook(stream);
        var worksheet = GetWorksheet(workbook, map);
        var headers = GetHeaders(worksheet, map);

        var parsedRows = new List<(int RowNumber, T Value)>();
        var errors = new List<ExcelImportError>();

        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var rowNumber = row.RowNumber();
            var item = new T();
            var rowErrors = new List<ExcelImportError>();

            foreach (var column in map.Columns)
            {
                if (!headers.TryGetValue(column.Header, out var columnNumber))
                    continue;

                if (column.Setter is null)
                    throw new ExcelMappingException(
                        $"Property mapped to '{column.Header}' is read-only and cannot be imported.");

                try
                {
                    ReadCell(row, columnNumber, column, item);
                }
                catch (Exception ex)
                {
                    rowErrors.Add(new ExcelImportError(rowNumber, column.Header, ex.Message));
                }
            }

            if (rowErrors.Count == 0)
            {
                rowErrors.AddRange(options.Validators
                    .Where(rule => !rule.Predicate(item))
                    .Select(rule => new ExcelImportError(rowNumber, null, rule.Message)));
            }

            parsedRows.Add((rowNumber, item));
            errors.AddRange(rowErrors);
        }

        foreach (var validator in options.BatchValidators)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var contextRows = parsedRows
                .Select(row => new ExcelImportRow<T>(
                    row.RowNumber,
                    row.Value,
                    errors.Where(error => error.Row == row.RowNumber).ToArray()))
                .ToArray();

            var context = new ExcelBatchValidationContext<T>(map.SheetName, contextRows);
            var validatorErrors = await validator.ValidateAsync(context, cancellationToken).ConfigureAwait(false);

            if (validatorErrors is not null)
                errors.AddRange(validatorErrors);
        }

        if (options.ErrorBehavior == ExcelImportErrorBehavior.Throw && errors.Count > 0)
            throw new ExcelMappingException(errors[0].ToString());

        var invalidRows = new HashSet<int>(errors.Where(x => x.Row > 0).Select(x => x.Row));
        var items = parsedRows
            .Where(row => !invalidRows.Contains(row.RowNumber))
            .Select(row => row.Value)
            .ToArray();

        return new ExcelImportResult<T>(items, errors.ToArray());
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

        WriteHeaders(worksheet, map, options);

        var rowNumber = 2;
        foreach (var item in items)
        {
            var rowRules = options.RowRules.Where(x => x.Predicate(item)).ToArray();

            for (var i = 0; i < map.Columns.Count; i++)
            {
                var column = map.Columns[i];
                var cell = worksheet.Cell(rowNumber, i + 1);
                var value = column.Getter(item!);

                if (column.Converter is null)
                    ExcelCellConverter.Write(cell, value);
                else
                    ExcelCellConverter.Write(cell, column.Converter.Write(value));

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

        FinishWorksheet(worksheet, options);
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

    public void CreateTemplate<T>(Stream stream) where T : new()
        => CreateTemplate<T>(stream, null);

    public void CreateTemplate<T>(
        Stream stream,
        Action<ExcelExportOptions<T>>? configure) where T : new()
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));

        var map = GetMap<T>();
        var options = new ExcelExportOptions<T>();
        configure?.Invoke(options);

        if (options.TemplateRows < 1)
            throw new ArgumentOutOfRangeException(nameof(options.TemplateRows), "TemplateRows must be greater than zero.");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(map.SheetName);

        WriteHeaders(worksheet, map, options);
        ApplyTemplateValidations(worksheet, map, options);
        FinishWorksheet(worksheet, options);
        workbook.SaveAs(stream);
    }

    public void CreateTemplate<T>(string path) where T : new()
        => CreateTemplate<T>(path, null);

    public void CreateTemplate<T>(
        string path,
        Action<ExcelExportOptions<T>>? configure) where T : new()
    {
        if (path is null) throw new ArgumentNullException(nameof(path));
        using var stream = File.Create(path);
        CreateTemplate<T>(stream, configure);
    }

    private static IXLWorksheet GetWorksheet(XLWorkbook workbook, ExcelTypeMap map)
    {
        var worksheet = workbook.Worksheets.FirstOrDefault(x =>
            x.Name.Equals(map.SheetName, StringComparison.OrdinalIgnoreCase));

        return worksheet ?? throw new ExcelMappingException($"Worksheet '{map.SheetName}' was not found.");
    }

    private static Dictionary<string, int> GetHeaders(IXLWorksheet worksheet, ExcelTypeMap map)
    {
        var headers = worksheet.Row(1)
            .CellsUsed()
            .ToDictionary(x => x.GetString(), x => x.Address.ColumnNumber, StringComparer.OrdinalIgnoreCase);

        var missingRequired = map.Columns
            .Where(x => x.Required && !headers.ContainsKey(x.Header))
            .Select(x => x.Header)
            .ToArray();

        if (missingRequired.Length > 0)
            throw new ExcelMappingException(
                $"Required columns were not found in worksheet '{map.SheetName}': {string.Join(", ", missingRequired)}.");

        return headers;
    }

    private static void ReadCell<T>(IXLRow row, int columnNumber, ExcelColumnMap column, T item)
    {
        var cell = row.Cell(columnNumber);
        var value = column.Converter is null
            ? ExcelCellConverter.Read(cell, column.ValueType)
            : column.Converter.Read(ExcelCellConverter.ToExcelValue(cell), column.ValueType);

        column.Setter!(item!, value);
    }

    private static void WriteHeaders<T>(
        IXLWorksheet worksheet,
        ExcelTypeMap map,
        ExcelExportOptions<T> options)
    {
        for (var i = 0; i < map.Columns.Count; i++)
        {
            var cell = worksheet.Cell(1, i + 1);
            cell.Value = map.Columns[i].Header;
            ClosedXmlStyleApplier.Apply(cell.Style, options.HeaderStyle);

            if (options.Columns.TryGetValue(map.Columns[i].PropertyName, out var columnConfig))
            {
                if (columnConfig.Width.HasValue)
                    worksheet.Column(i + 1).Width = columnConfig.Width.Value;

                ClosedXmlStyleApplier.Apply(worksheet.Column(i + 1).Style, columnConfig.Style);
            }
        }
    }

    private static void ApplyTemplateValidations<T>(
        IXLWorksheet worksheet,
        ExcelTypeMap map,
        ExcelExportOptions<T> options)
    {
        for (var i = 0; i < map.Columns.Count; i++)
        {
            var column = map.Columns[i];
            options.Columns.TryGetValue(column.PropertyName, out var config);

            var values = config?.AllowedValues;
            if (values is null)
            {
                var enumType = Nullable.GetUnderlyingType(column.ValueType) ?? column.ValueType;
                if (enumType.IsEnum)
                    values = Enum.GetNames(enumType);
            }

            if (values is null || values.Count == 0)
                continue;

            var formula = string.Join(",", values);
            if (formula.Length > 255)
                throw new ExcelMappingException(
                    $"Allowed values for '{column.Header}' exceed Excel's 255-character inline validation limit.");

            var range = worksheet.Range(2, i + 1, options.TemplateRows + 1, i + 1);
            var validation = range.CreateDataValidation();
            validation.List(formula, true);
            validation.ErrorTitle = "Invalid value";
            validation.ErrorMessage = $"Choose one of the allowed values for {column.Header}.";
            validation.ShowErrorMessage = true;
        }
    }

    private static void FinishWorksheet<T>(IXLWorksheet worksheet, ExcelExportOptions<T> options)
    {
        if (options.FreezeHeader)
            worksheet.SheetView.FreezeRows(1);

        if (options.AutoFilter && worksheet.RangeUsed() is { } range)
            range.SetAutoFilter();
    }

    private static ExcelTypeMap GetMap<T>() where T : new()
    {
        var probe = new T();
        return probe is IGeneratedExcelMapProvider provider
            ? provider.ExcelMap
            : throw new GeneratedMapNotFoundException(typeof(T));
    }
}
