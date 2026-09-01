using ClosedXML.Excel;
using Indtec.ExcelMapper.Importing;
using Indtec.ExcelMapper.Styling;

namespace Indtec.ExcelMapper.Workbooks;

public sealed class ExcelWorkbookImportOptions
{
    internal List<IExcelWorkbookImportRegistration> Sheets { get; } = new();
    internal List<IExcelWorkbookValidator> Validators { get; } = new();

    public ExcelWorkbookImportOptions Sheet<T>(Action<ExcelImportOptions<T>>? configure = null) where T : new()
    {
        var options = new ExcelImportOptions<T>();
        configure?.Invoke(options);
        Sheets.Add(new ExcelWorkbookImportRegistration<T>(options));
        return this;
    }

    public ExcelWorkbookImportOptions AddValidator(IExcelWorkbookValidator validator)
    {
        if (validator is null) throw new ArgumentNullException(nameof(validator));
        Validators.Add(validator);
        return this;
    }
}

public sealed class ExcelWorkbookTemplateOptions
{
    internal List<IExcelWorkbookTemplateRegistration> Sheets { get; } = new();

    public ExcelWorkbookTemplateOptions Sheet<T>(Action<ExcelExportOptions<T>>? configure = null) where T : new()
    {
        var options = new ExcelExportOptions<T>();
        configure?.Invoke(options);
        Sheets.Add(new ExcelWorkbookTemplateRegistration<T>(options));
        return this;
    }
}

public sealed class ExcelWorkbookImportResult
{
    private readonly Dictionary<Type, object> _results;

    internal ExcelWorkbookImportResult(Dictionary<Type, object> results)
        => _results = results;

    public ExcelImportResult<T> Sheet<T>()
    {
        if (_results.TryGetValue(typeof(T), out var result) && result is ExcelImportResult<T> typed)
            return typed;

        throw new ExcelMappingException($"Sheet result for '{typeof(T).Name}' was not registered.");
    }
}

internal interface IExcelWorkbookImportRegistration
{
    Type ModelType { get; }
    Task<object> ImportAsync(ExcelMapper mapper, XLWorkbook workbook, CancellationToken cancellationToken);
}

internal sealed class ExcelWorkbookImportRegistration<T> : IExcelWorkbookImportRegistration where T : new()
{
    private readonly ExcelImportOptions<T> _options;

    public ExcelWorkbookImportRegistration(ExcelImportOptions<T> options) => _options = options;

    public Type ModelType => typeof(T);

    public async Task<object> ImportAsync(
        ExcelMapper mapper,
        XLWorkbook workbook,
        CancellationToken cancellationToken)
        => await mapper.ImportSheetAsync<T>(workbook, _options, cancellationToken).ConfigureAwait(false);
}

internal interface IExcelWorkbookTemplateRegistration
{
    void AddSheet(ExcelMapper mapper, XLWorkbook workbook);
}

internal sealed class ExcelWorkbookTemplateRegistration<T> : IExcelWorkbookTemplateRegistration where T : new()
{
    private readonly ExcelExportOptions<T> _options;

    public ExcelWorkbookTemplateRegistration(ExcelExportOptions<T> options) => _options = options;

    public void AddSheet(ExcelMapper mapper, XLWorkbook workbook)
        => mapper.AddTemplateSheet<T>(workbook, _options);
}
