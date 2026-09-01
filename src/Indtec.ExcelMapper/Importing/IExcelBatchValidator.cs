namespace Indtec.ExcelMapper.Importing;

public interface IExcelBatchValidator<T>
{
    ValueTask<IReadOnlyList<ExcelImportError>> ValidateAsync(
        ExcelBatchValidationContext<T> context,
        CancellationToken cancellationToken = default);
}

public sealed class ExcelBatchValidationContext<T>
{
    internal ExcelBatchValidationContext(
        string sheetName,
        IReadOnlyList<ExcelImportRow<T>> rows)
    {
        SheetName = sheetName;
        Rows = rows;
    }

    public string SheetName { get; }
    public IReadOnlyList<ExcelImportRow<T>> Rows { get; }
}

public sealed class ExcelImportRow<T>
{
    internal ExcelImportRow(
        int rowNumber,
        T value,
        IReadOnlyList<ExcelImportError> errors)
    {
        RowNumber = rowNumber;
        Value = value;
        Errors = errors;
    }

    public int RowNumber { get; }
    public T Value { get; }
    public IReadOnlyList<ExcelImportError> Errors { get; }
    public bool HasErrors => Errors.Count > 0;
}
