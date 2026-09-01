namespace Indtec.ExcelMapper.Localization;

public abstract class ExcelMessageProvider : IExcelMessageProvider
{
    protected virtual IExcelMessageProvider Fallback => ExcelMessages.English;

    public virtual string BatchValidatorsRequireAsync() => Fallback.BatchValidatorsRequireAsync();
    public virtual string StreamingBatchValidatorsNotSupported() => Fallback.StreamingBatchValidatorsNotSupported();
    public virtual string WorksheetNotFound(string sheetName) => Fallback.WorksheetNotFound(sheetName);
    public virtual string RequiredColumnsNotFound(string sheetName, IReadOnlyList<string> columns) => Fallback.RequiredColumnsNotFound(sheetName, columns);
    public virtual string ReadOnlyProperty(string header) => Fallback.ReadOnlyProperty(header);
    public virtual string EmptyCellNotNullable(string address, string typeName) => Fallback.EmptyCellNotNullable(address, typeName);
    public virtual string CouldNotConvertCell(string address, string typeName) => Fallback.CouldNotConvertCell(address, typeName);
    public virtual string MissingCellReference() => Fallback.MissingCellReference();
    public virtual string InvalidCellReference(string reference) => Fallback.InvalidCellReference(reference);
    public virtual string AtLeastOneSheetForImport() => Fallback.AtLeastOneSheetForImport();
    public virtual string DuplicateWorkbookModel(string modelName) => Fallback.DuplicateWorkbookModel(modelName);
    public virtual string UnregisteredWorkbookModel(string modelName) => Fallback.UnregisteredWorkbookModel(modelName);
    public virtual string AtLeastOneSheetForTemplate() => Fallback.AtLeastOneSheetForTemplate();
    public virtual string DuplicateWorksheet(string sheetName) => Fallback.DuplicateWorksheet(sheetName);
    public virtual string InvalidTemplateRows() => Fallback.InvalidTemplateRows();
    public virtual string AllowedValuesTooLong(string header) => Fallback.AllowedValuesTooLong(header);
    public virtual string InvalidValueTitle() => Fallback.InvalidValueTitle();
    public virtual string InvalidAllowedValue(string header) => Fallback.InvalidAllowedValue(header);
}
