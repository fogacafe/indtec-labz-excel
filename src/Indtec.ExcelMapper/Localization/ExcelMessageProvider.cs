namespace Indtec.ExcelMapper.Localization;

/// <summary>
/// Base class for customizing mapper-generated messages without implementing the entire
/// <see cref="IExcelMessageProvider"/> contract. Override only the messages you want to change;
/// every other member falls back to the built-in English provider.
/// </summary>
public abstract class ExcelMessageProvider : IExcelMessageProvider
{
    /// <summary>Gets the provider used for messages that are not overridden. Defaults to the built-in English provider.</summary>
    protected virtual IExcelMessageProvider Fallback => ExcelMessages.English;

    /// <summary>Used when synchronous import is attempted with async batch validators. Default: &quot;Batch validators require ImportAsync because they may perform asynchronous work.&quot;</summary>
    public virtual string BatchValidatorsRequireAsync() => Fallback.BatchValidatorsRequireAsync();

    /// <summary>Used when full-sheet batch validators are registered for streaming import. The default explains that validation must happen in the chunk callback or through ImportAsync.</summary>
    public virtual string StreamingBatchValidatorsNotSupported() => Fallback.StreamingBatchValidatorsNotSupported();

    /// <summary>Used when the expected worksheet is missing. Default example: &quot;Worksheet 'Products' was not found.&quot;</summary>
    public virtual string WorksheetNotFound(string sheetName) => Fallback.WorksheetNotFound(sheetName);

    /// <summary>Used when required mapped headers are missing. Default example: &quot;Required columns were not found in worksheet 'Products': Id, Name.&quot;</summary>
    public virtual string RequiredColumnsNotFound(string sheetName, IReadOnlyList<string> columns) => Fallback.RequiredColumnsNotFound(sheetName, columns);

    /// <summary>Used when import targets a mapped read-only property. Default example: &quot;Property mapped to 'Id' is read-only and cannot be imported.&quot;</summary>
    public virtual string ReadOnlyProperty(string header) => Fallback.ReadOnlyProperty(header);

    /// <summary>Used when an empty cell maps to a non-nullable type. Default example: &quot;Cell A2 is empty but 'Int32' is not nullable.&quot;</summary>
    public virtual string EmptyCellNotNullable(string address, string typeName) => Fallback.EmptyCellNotNullable(address, typeName);

    /// <summary>Used when cell conversion fails. Default example: &quot;Could not convert cell D3 to 'Decimal'.&quot;</summary>
    public virtual string CouldNotConvertCell(string address, string typeName) => Fallback.CouldNotConvertCell(address, typeName);

    /// <summary>Used when streaming import finds a cell without an Excel reference. Default: &quot;A streamed cell is missing its Excel reference.&quot;</summary>
    public virtual string MissingCellReference() => Fallback.MissingCellReference();

    /// <summary>Used when streaming import finds an invalid Excel cell reference. Default example: &quot;Invalid Excel cell reference '???'.&quot;</summary>
    public virtual string InvalidCellReference(string reference) => Fallback.InvalidCellReference(reference);

    /// <summary>Used when workbook import has no registered sheets. Default: &quot;At least one sheet must be registered for workbook import.&quot;</summary>
    public virtual string AtLeastOneSheetForImport() => Fallback.AtLeastOneSheetForImport();

    /// <summary>Used when workbook export has no registered sheets. Default: &quot;At least one sheet must be registered for workbook export.&quot;</summary>
    public virtual string AtLeastOneSheetForExport() => Fallback.AtLeastOneSheetForExport();

    /// <summary>Used when an empty row is configured as an error. Default example: &quot;Row 3 is empty.&quot;</summary>
    public virtual string EmptyRow(int row) => Fallback.EmptyRow(row);

    /// <summary>Used when the same model is registered more than once for workbook import. Default example: &quot;Model 'Product' was registered more than once in the workbook import.&quot;</summary>
    public virtual string DuplicateWorkbookModel(string modelName) => Fallback.DuplicateWorkbookModel(modelName);

    /// <summary>Used when a workbook validator targets an unregistered model. Default example: &quot;Workbook validator returned an error for unregistered model 'Product'.&quot;</summary>
    public virtual string UnregisteredWorkbookModel(string modelName) => Fallback.UnregisteredWorkbookModel(modelName);

    /// <summary>Used when workbook template generation has no registered sheets. Default: &quot;At least one sheet must be registered for workbook template generation.&quot;</summary>
    public virtual string AtLeastOneSheetForTemplate() => Fallback.AtLeastOneSheetForTemplate();

    /// <summary>Used when workbook registrations resolve to the same worksheet name. Default example: &quot;Worksheet 'Products' was registered more than once.&quot;</summary>
    public virtual string DuplicateWorksheet(string sheetName) => Fallback.DuplicateWorksheet(sheetName);

    /// <summary>Used when TemplateRows is zero or negative. Default: &quot;TemplateRows must be greater than zero.&quot;</summary>
    public virtual string InvalidTemplateRows() => Fallback.InvalidTemplateRows();

    /// <summary>Used when an inline validation list exceeds Excel's supported length. Default example: &quot;Allowed values for 'Status' exceed Excel's 255-character inline validation limit.&quot;</summary>
    public virtual string AllowedValuesTooLong(string header) => Fallback.AllowedValuesTooLong(header);

    /// <summary>Title for generated Excel validation prompts. Default: &quot;Invalid value&quot;.</summary>
    public virtual string InvalidValueTitle() => Fallback.InvalidValueTitle();

    /// <summary>Message for generated Excel validation prompts. Default example: &quot;Choose one of the allowed values for Status.&quot;</summary>
    public virtual string InvalidAllowedValue(string header) => Fallback.InvalidAllowedValue(header);
}
