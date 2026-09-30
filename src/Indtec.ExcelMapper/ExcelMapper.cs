using ClosedXML.Excel;
using Indtec.ExcelMapper.Importing;
using Indtec.ExcelMapper.Internal;
using Indtec.ExcelMapper.Localization;
using Indtec.ExcelMapper.Mapping;
using Indtec.ExcelMapper.Styling;
using Indtec.ExcelMapper.Workbooks;

namespace Indtec.ExcelMapper;

/// <summary>Imports, exports and generates Excel workbooks for source-generated mapped models.</summary>
public sealed class ExcelMapper
{
    private readonly IExcelMessageProvider _messages;

    internal IExcelMessageProvider Messages => _messages;

    /// <summary>Creates a mapper using the default options and English built-in messages.</summary>
    public ExcelMapper()
        : this(new ExcelMapperOptions())
    {
    }

    /// <summary>Creates a mapper using the supplied localization and message options.</summary>
    public ExcelMapper(ExcelMapperOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        _messages = options.ResolveMessages();
    }

    /// <summary>Creates a mapper and configures localization and generated messages.</summary>
    public ExcelMapper(Action<ExcelMapperOptions> configure)
    {
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        var options = new ExcelMapperOptions();
        configure(options);
        _messages = options.ResolveMessages();
    }

    /// <summary>Imports valid items from the mapped worksheet in an XLSX stream.</summary>
    public IReadOnlyList<T> Import<T>(Stream stream) where T : new()
        => Import<T>(stream, null).Items;

    /// <summary>Imports a mapped worksheet and returns items plus structured validation errors.</summary>
    public ExcelImportResult<T> Import<T>(
        Stream stream,
        Action<ExcelImportOptions<T>>? configure) where T : new()
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));

        var map = GetMap<T>();
        var options = new ExcelImportOptions<T>();
        configure?.Invoke(options);
        options.ValidateConfiguration();

        if (options.BatchValidators.Count > 0)
            throw new ExcelMappingException(_messages.BatchValidatorsRequireAsync());

        using var workbook = new XLWorkbook(stream);
        var worksheet = GetWorksheet(workbook, map);
        var headers = GetHeaders(worksheet, map);

        var items = new List<T>();
        var errors = new List<ExcelImportError>();
        var reachedInvalidRowLimit = false;

        foreach (var row in GetDataRows(worksheet, options.EmptyRowBehavior))
        {
            if (IsMappedRowEmpty(row, headers, map))
            {
                if (options.EmptyRowBehavior == ExcelEmptyRowBehavior.Stop)
                    break;
                if (options.EmptyRowBehavior == ExcelEmptyRowBehavior.Ignore)
                    continue;
                if (options.EmptyRowBehavior == ExcelEmptyRowBehavior.Error)
                {
                    var error = new ExcelImportError(row.RowNumber(), null, _messages.EmptyRow(row.RowNumber()));
                    if (options.ErrorBehavior == ExcelImportErrorBehavior.Throw)
                        throw new ExcelMappingException(error.ToString());
                    errors.Add(error);
                    if (HasReachedInvalidRowLimit(errors, options.MaxInvalidRows))
                    {
                        reachedInvalidRowLimit = true;
                        break;
                    }
                    continue;
                }
            }

            var item = new T();
            var rowHasMappingErrors = false;

            foreach (var column in map.Columns)
            {
                if (!TryGetColumnNumber(headers, column, out var columnNumber))
                    continue;

                if (column.Setter is null)
                    throw new ExcelMappingException(_messages.ReadOnlyProperty(column.Header));

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
            {
                if (HasReachedInvalidRowLimit(errors, options.MaxInvalidRows))
                {
                    reachedInvalidRowLimit = true;
                    break;
                }
                continue;
            }

            NormalizeItem(item, map, options);

            var validationErrors = options.Validators
                .Where(rule => !rule.Predicate(item))
                .Select(rule => new ExcelImportError(row.RowNumber(), GetValidationColumn(map, rule.PropertyName), rule.Message))
                .ToArray();

            if (validationErrors.Length > 0)
            {
                if (options.ErrorBehavior == ExcelImportErrorBehavior.Throw)
                    throw new ExcelMappingException(validationErrors[0].ToString());

                errors.AddRange(validationErrors);
                if (HasReachedInvalidRowLimit(errors, options.MaxInvalidRows))
                {
                    reachedInvalidRowLimit = true;
                    break;
                }
                continue;
            }

            items.Add(item);
        }

        return new ExcelImportResult<T>(items, errors, reachedInvalidRowLimit: reachedInvalidRowLimit);
    }

    /// <summary>Imports a mapped worksheet asynchronously, including optional batch validation.</summary>
    public async Task<ExcelImportResult<T>> ImportAsync<T>(
        Stream stream,
        Action<ExcelImportOptions<T>>? configure = null,
        CancellationToken cancellationToken = default) where T : new()
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));

        var options = new ExcelImportOptions<T>();
        configure?.Invoke(options);
        options.ValidateConfiguration();

        using var workbook = new XLWorkbook(stream);
        return await ImportSheetAsync<T>(workbook, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Imports multiple registered worksheets and runs optional workbook-level validation.</summary>
    public async Task<ExcelWorkbookImportResult> ImportWorkbookAsync(
        Stream stream,
        Action<ExcelWorkbookImportOptions> configure,
        CancellationToken cancellationToken = default)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (configure is null) throw new ArgumentNullException(nameof(configure));

        var options = new ExcelWorkbookImportOptions();
        configure(options);

        if (options.Sheets.Count == 0)
            throw new ExcelMappingException(_messages.AtLeastOneSheetForImport());

        var duplicateModel = options.Sheets
            .GroupBy(x => x.ModelType)
            .FirstOrDefault(x => x.Count() > 1);

        if (duplicateModel is not null)
            throw new ExcelMappingException(_messages.DuplicateWorkbookModel(duplicateModel.Key.Name));

        using var workbook = new XLWorkbook(stream);
        var results = new Dictionary<Type, object>();

        foreach (var sheet in options.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            results[sheet.ModelType] = await sheet.ImportAsync(this, workbook, cancellationToken).ConfigureAwait(false);
        }

        foreach (var validator in options.Validators)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var context = new ExcelWorkbookValidationContext(results);
            var validatorErrors = await validator.ValidateAsync(context, cancellationToken).ConfigureAwait(false)
                ?? Array.Empty<ExcelWorkbookValidationError>();

            if (validatorErrors.Count == 0)
                continue;

            foreach (var error in validatorErrors)
            {
                if (!results.ContainsKey(error.ModelType))
                    throw new ExcelMappingException(_messages.UnregisteredWorkbookModel(error.ModelType.Name));
            }

            var firstThrowingError = validatorErrors.FirstOrDefault(error =>
                options.Sheets.Single(sheet => sheet.ModelType == error.ModelType).ShouldThrow);

            if (firstThrowingError is not null)
            {
                var importError = new ExcelImportError(
                    firstThrowingError.Row,
                    firstThrowingError.Column,
                    firstThrowingError.Message);

                throw new ExcelMappingException(importError.ToString());
            }

            foreach (var group in validatorErrors.GroupBy(error => error.ModelType))
            {
                var registration = options.Sheets.Single(sheet => sheet.ModelType == group.Key);
                results[group.Key] = registration.AddValidationErrors(results[group.Key], group.ToArray());
            }
        }

        return new ExcelWorkbookImportResult(results);
    }

    internal async Task<ExcelImportResult<T>> ImportSheetAsync<T>(
        XLWorkbook workbook,
        ExcelImportOptions<T> options,
        CancellationToken cancellationToken) where T : new()
    {
        options.ValidateConfiguration();
        var map = GetMap<T>();
        var worksheet = GetWorksheet(workbook, map);
        var headers = GetHeaders(worksheet, map);

        var parsedRows = new List<(int RowNumber, T Value)>();
        var errors = new List<ExcelImportError>();
        var reachedInvalidRowLimit = false;

        foreach (var row in GetDataRows(worksheet, options.EmptyRowBehavior))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var rowNumber = row.RowNumber();
            if (IsMappedRowEmpty(row, headers, map))
            {
                if (options.EmptyRowBehavior == ExcelEmptyRowBehavior.Stop)
                    break;
                if (options.EmptyRowBehavior == ExcelEmptyRowBehavior.Ignore)
                    continue;
                if (options.EmptyRowBehavior == ExcelEmptyRowBehavior.Error)
                {
                    errors.Add(new ExcelImportError(rowNumber, null, _messages.EmptyRow(rowNumber)));
                    if (HasReachedInvalidRowLimit(errors, options.MaxInvalidRows))
                    {
                        reachedInvalidRowLimit = true;
                        break;
                    }
                    continue;
                }
            }
            var item = new T();
            var rowErrors = new List<ExcelImportError>();

            foreach (var column in map.Columns)
            {
                if (!TryGetColumnNumber(headers, column, out var columnNumber))
                    continue;

                if (column.Setter is null)
                    throw new ExcelMappingException(_messages.ReadOnlyProperty(column.Header));

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
                NormalizeItem(item, map, options);

            if (rowErrors.Count == 0)
            {
                rowErrors.AddRange(options.Validators
                    .Where(rule => !rule.Predicate(item))
                    .Select(rule => new ExcelImportError(rowNumber, GetValidationColumn(map, rule.PropertyName), rule.Message)));
            }

            parsedRows.Add((rowNumber, item));
            errors.AddRange(rowErrors);

            if (rowErrors.Count > 0 && HasReachedInvalidRowLimit(errors, options.MaxInvalidRows))
            {
                reachedInvalidRowLimit = true;
                break;
            }
        }

        if (!reachedInvalidRowLimit)
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

        var rows = parsedRows
            .Select(row => new ExcelImportRow<T>(
                row.RowNumber,
                row.Value,
                errors.Where(error => error.Row == row.RowNumber).ToArray()))
            .ToArray();

        var invalidRows = new HashSet<int>(errors.Where(x => x.Row > 0).Select(x => x.Row));
        var items = rows
            .Where(row => !invalidRows.Contains(row.RowNumber))
            .Select(row => row.Value)
            .ToArray();

        return new ExcelImportResult<T>(items, errors.ToArray(), rows, reachedInvalidRowLimit);
    }

    /// <summary>Exports mapped items to an XLSX stream using default export options.</summary>
    public void Export<T>(IEnumerable<T> items, Stream stream) where T : new()
        => Export(items, stream, null);

    /// <summary>Exports mapped items to an XLSX stream using configurable worksheet options.</summary>
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
        var worksheet = workbook.AddWorksheet(ResolveSheetName(map, options));

        WriteHeaders(worksheet, map, options);

        WriteRows(worksheet, map, options, items);

        FinishWorksheet(worksheet, map, options);
        workbook.SaveAs(stream);
    }

    /// <summary>Exports mapped items to an XLSX file using default export options.</summary>
    public void Export<T>(IEnumerable<T> items, string path) where T : new()
        => Export(items, path, null);

    /// <summary>Exports mapped items to an XLSX file using configurable worksheet options.</summary>
    public void Export<T>(
        IEnumerable<T> items,
        string path,
        Action<ExcelExportOptions<T>>? configure) where T : new()
    {
        if (path is null) throw new ArgumentNullException(nameof(path));
        using var stream = File.Create(path);
        Export(items, stream, configure);
    }

    /// <summary>Exports multiple registered worksheets to an XLSX stream.</summary>
    public void ExportWorkbook(
        Stream stream,
        Action<ExcelWorkbookExportOptions> configure)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (configure is null) throw new ArgumentNullException(nameof(configure));

        var options = new ExcelWorkbookExportOptions();
        configure(options);

        if (options.Sheets.Count == 0)
            throw new ExcelMappingException(_messages.AtLeastOneSheetForExport());

        using var workbook = new XLWorkbook();
        foreach (var sheet in options.Sheets)
            sheet.AddSheet(this, workbook);

        workbook.SaveAs(stream);
    }

    /// <summary>Exports multiple registered worksheets to an XLSX file.</summary>
    public void ExportWorkbook(
        string path,
        Action<ExcelWorkbookExportOptions> configure)
    {
        if (path is null) throw new ArgumentNullException(nameof(path));
        using var stream = File.Create(path);
        ExportWorkbook(stream, configure);
    }

    internal void AddExportSheet<T>(
        XLWorkbook workbook,
        ExcelExportOptions<T> options,
        IEnumerable<T> items) where T : new()
    {
        var map = GetMap<T>();
        var sheetName = ResolveSheetName(map, options);
        if (workbook.Worksheets.Any(x => x.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
            throw new ExcelMappingException(_messages.DuplicateWorksheet(sheetName));

        var worksheet = workbook.AddWorksheet(sheetName);
        WriteHeaders(worksheet, map, options);
        WriteRows(worksheet, map, options, items);
        FinishWorksheet(worksheet, map, options);
    }

    /// <summary>Generates a blank mapped worksheet template in an XLSX stream.</summary>
    public void CreateTemplate<T>(Stream stream) where T : new()
        => CreateTemplate<T>(stream, null);

    /// <summary>Generates a configurable mapped worksheet template in an XLSX stream.</summary>
    public void CreateTemplate<T>(
        Stream stream,
        Action<ExcelExportOptions<T>>? configure) where T : new()
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));

        var options = new ExcelExportOptions<T>();
        configure?.Invoke(options);

        using var workbook = new XLWorkbook();
        AddTemplateSheet<T>(workbook, options);
        workbook.SaveAs(stream);
    }

    /// <summary>Generates a blank mapped worksheet template as an XLSX file.</summary>
    public void CreateTemplate<T>(string path) where T : new()
        => CreateTemplate<T>(path, null);

    /// <summary>Generates a configurable mapped worksheet template as an XLSX file.</summary>
    public void CreateTemplate<T>(
        string path,
        Action<ExcelExportOptions<T>>? configure) where T : new()
    {
        if (path is null) throw new ArgumentNullException(nameof(path));
        using var stream = File.Create(path);
        CreateTemplate<T>(stream, configure);
    }

    /// <summary>Generates a multi-sheet workbook template in an XLSX stream.</summary>
    public void CreateWorkbookTemplate(
        Stream stream,
        Action<ExcelWorkbookTemplateOptions> configure)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (configure is null) throw new ArgumentNullException(nameof(configure));

        var options = new ExcelWorkbookTemplateOptions();
        configure(options);

        if (options.Sheets.Count == 0)
            throw new ExcelMappingException(_messages.AtLeastOneSheetForTemplate());

        using var workbook = new XLWorkbook();
        foreach (var sheet in options.Sheets)
            sheet.AddSheet(this, workbook);

        workbook.SaveAs(stream);
    }

    /// <summary>Generates a multi-sheet workbook template as an XLSX file.</summary>
    public void CreateWorkbookTemplate(
        string path,
        Action<ExcelWorkbookTemplateOptions> configure)
    {
        if (path is null) throw new ArgumentNullException(nameof(path));
        using var stream = File.Create(path);
        CreateWorkbookTemplate(stream, configure);
    }

    internal void AddTemplateSheet<T>(
        XLWorkbook workbook,
        ExcelExportOptions<T> options,
        IReadOnlyList<T>? items = null) where T : new()
    {
        if (options.TemplateRows < 1)
            throw new ArgumentOutOfRangeException(nameof(options.TemplateRows), _messages.InvalidTemplateRows());

        var map = GetMap<T>();
        var sheetName = ResolveSheetName(map, options);
        if (workbook.Worksheets.Any(x => x.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
            throw new ExcelMappingException(_messages.DuplicateWorksheet(sheetName));

        var worksheet = workbook.AddWorksheet(sheetName);
        WriteHeaders(worksheet, map, options);

        if (items is { Count: > 0 })
            WriteRows(worksheet, map, options, items);

        ApplyTemplateValidations(
            workbook,
            worksheet,
            map,
            options,
            Math.Max(options.TemplateRows, items?.Count ?? 0));

        FinishWorksheet(worksheet, map, options);
    }

    private IXLWorksheet GetWorksheet(XLWorkbook workbook, ExcelTypeMap map)
    {
        var worksheet = workbook.Worksheets.FirstOrDefault(x =>
            x.Name.Equals(map.SheetName, StringComparison.OrdinalIgnoreCase));

        return worksheet ?? throw new ExcelMappingException(_messages.WorksheetNotFound(map.SheetName));
    }

    private Dictionary<string, int> GetHeaders(IXLWorksheet worksheet, ExcelTypeMap map)
    {
        var headers = worksheet.Row(1)
            .CellsUsed()
            .ToDictionary(x => x.GetString(), x => x.Address.ColumnNumber, StringComparer.OrdinalIgnoreCase);

        var missingRequired = map.Columns
            .Where(x => x.Required && !x.AcceptedHeaders().Any(headers.ContainsKey))
            .Select(x => x.Header)
            .ToArray();

        if (missingRequired.Length > 0)
            throw new ExcelMappingException(_messages.RequiredColumnsNotFound(map.SheetName, missingRequired));

        return headers;
    }

    internal bool WorksheetExists<T>(XLWorkbook workbook) where T : new()
    {
        var map = GetMap<T>();
        return workbook.Worksheets.Any(x =>
            x.Name.Equals(map.SheetName, StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryGetColumnNumber(
        IReadOnlyDictionary<string, int> headers,
        ExcelColumnMap column,
        out int columnNumber)
    {
        foreach (var header in column.AcceptedHeaders())
        {
            if (headers.TryGetValue(header, out columnNumber))
                return true;
        }

        columnNumber = 0;
        return false;
    }

    private static string? GetValidationColumn(ExcelTypeMap map, string? propertyName)
        => propertyName is null
            ? null
            : map.Columns.FirstOrDefault(x => x.PropertyName == propertyName)?.Header ?? propertyName;

    private static string ResolveSheetName<T>(ExcelTypeMap map, ExcelExportOptions<T> options)
        => string.IsNullOrWhiteSpace(options.SheetName) ? map.SheetName : options.SheetName!;

    private static IEnumerable<IXLRow> GetDataRows(
        IXLWorksheet worksheet,
        ExcelEmptyRowBehavior emptyRowBehavior)
    {
        if (emptyRowBehavior == ExcelEmptyRowBehavior.Ignore)
            return worksheet.RowsUsed().Skip(1);

        var lastRow = worksheet.LastRowUsed(XLCellsUsedOptions.All)?.RowNumber() ?? 1;
        return lastRow <= 1
            ? Enumerable.Empty<IXLRow>()
            : worksheet.Rows(2, lastRow);
    }

    private static bool IsMappedRowEmpty(
        IXLRow row,
        IReadOnlyDictionary<string, int> headers,
        ExcelTypeMap map)
        => map.Columns
            .Select(column => TryGetColumnNumber(headers, column, out var number) ? number : 0)
            .Where(number => number > 0)
            .All(number => row.Cell(number).IsEmpty());

    private void ReadCell<T>(IXLRow row, int columnNumber, ExcelColumnMap column, T item)
    {
        var cell = row.Cell(columnNumber);
        var value = column.Converter is null
            ? ExcelCellConverter.Read(cell, column.ValueType, _messages)
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
            var mappedColumn = map.Columns[i];
            cell.Value = mappedColumn.Header;
            ClosedXmlStyleApplier.Apply(cell.Style, options.HeaderStyle);
            ApplyDefaultColumnFormat(worksheet.Column(i + 1), mappedColumn.ValueType);

            if (options.Columns.TryGetValue(map.Columns[i].PropertyName, out var columnConfig))
            {
                if (columnConfig.Width.HasValue)
                    worksheet.Column(i + 1).Width = columnConfig.Width.Value;
                else if (options.AutoFitHeaders)
                    worksheet.Column(i + 1).AdjustToContents(1, 1);

                ClosedXmlStyleApplier.Apply(worksheet.Column(i + 1).Style, columnConfig.Style);
            }
            else if (options.AutoFitHeaders)
            {
                worksheet.Column(i + 1).AdjustToContents(1, 1);
            }
        }
    }

    private static void WriteRows<T>(
        IXLWorksheet worksheet,
        ExcelTypeMap map,
        ExcelExportOptions<T> options,
        IEnumerable<T> items)
    {
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
    }

    private void ApplyTemplateValidations<T>(
        XLWorkbook workbook,
        IXLWorksheet worksheet,
        ExcelTypeMap map,
        ExcelExportOptions<T> options,
        int templateRows)
    {
        const string validationSheetName = "__IndtecValidation";
        IXLWorksheet? validationSheet = null;

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

            validationSheet ??= workbook.Worksheets.FirstOrDefault(x =>
                x.Name.Equals(validationSheetName, StringComparison.OrdinalIgnoreCase))
                ?? workbook.AddWorksheet(validationSheetName);

            var validationColumn = (validationSheet.LastColumnUsed()?.ColumnNumber() ?? 0) + 1;
            for (var valueIndex = 0; valueIndex < values.Count; valueIndex++)
                validationSheet.Cell(valueIndex + 1, validationColumn).Value = values[valueIndex];

            var valuesRange = validationSheet.Range(1, validationColumn, values.Count, validationColumn);
            var rangeName = $"__IndtecValidation_{Guid.NewGuid():N}";
            workbook.DefinedNames.Add(rangeName, valuesRange).Visible = false;

            var targetRange = worksheet.Range(2, i + 1, templateRows + 1, i + 1);
            var validation = targetRange.CreateDataValidation();
            validation.List($"={rangeName}");
            validation.ErrorTitle = _messages.InvalidValueTitle();
            validation.ErrorMessage = _messages.InvalidAllowedValue(column.Header);
            validation.ShowErrorMessage = true;
        }

        if (validationSheet is not null)
            validationSheet.Visibility = XLWorksheetVisibility.VeryHidden;
    }

    private static void ApplyDefaultColumnFormat(IXLColumn column, Type valueType)
    {
        var type = Nullable.GetUnderlyingType(valueType) ?? valueType;

        if (type == typeof(DateTime))
            column.Style.NumberFormat.Format = "yyyy-mm-dd hh:mm:ss";
        else if (type == typeof(TimeSpan))
            column.Style.NumberFormat.Format = "[h]:mm:ss";
        else if (type == typeof(decimal) || type == typeof(double) || type == typeof(float))
            column.Style.NumberFormat.Format = "#,##0.########";
        else if (type == typeof(byte) || type == typeof(short) || type == typeof(int) || type == typeof(long))
            column.Style.NumberFormat.Format = "0";
        else if (type == typeof(string) || type == typeof(Guid) || type.IsEnum)
            column.Style.NumberFormat.Format = "@";
    }

    private static void NormalizeItem<T>(T item, ExcelTypeMap map, ExcelImportOptions<T> options)
    {
        foreach (var pair in options.Normalizers)
        {
            var column = map.Columns.First(x => x.PropertyName == pair.Key);
            if (column.Setter is null)
                continue;

            var value = column.Getter(item!);
            if (value is null || value is string text && string.IsNullOrWhiteSpace(text))
                continue;

            column.Setter(item!, pair.Value(value));
        }
    }

    private static bool HasReachedInvalidRowLimit(
        IReadOnlyCollection<ExcelImportError> errors,
        int? maxInvalidRows)
        => maxInvalidRows.HasValue &&
           errors.Where(error => error.Row > 0).Select(error => error.Row).Distinct().Take(maxInvalidRows.Value).Count() >= maxInvalidRows.Value;

    private static void FinishWorksheet<T>(IXLWorksheet worksheet, ExcelTypeMap map, ExcelExportOptions<T> options)
    {
        if (options.AutoFitColumns)
        {
            foreach (var column in worksheet.ColumnsUsed())
            {
                var columnIndex = column.ColumnNumber();
                var propertyName = map.Columns[columnIndex - 1].PropertyName;
                var hasExplicitWidth = options.Columns.TryGetValue(propertyName, out var config) && config.Width.HasValue;

                if (!hasExplicitWidth)
                    column.AdjustToContents();
            }
        }

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
