namespace Indtec.ExcelMapper.Importing;

/// <summary>
/// Validates a complete mapped worksheet asynchronously and returns additional immutable import errors.
/// </summary>
/// <typeparam name="T">The mapped row model type.</typeparam>
public interface IExcelBatchValidator<T>
{
    /// <summary>Validates the complete set of parsed rows for a worksheet.</summary>
    /// <param name="context">The worksheet name and all parsed rows, including rows with existing errors.</param>
    /// <param name="cancellationToken">Token used to cancel asynchronous validation.</param>
    /// <returns>Additional errors to be aggregated into the import result.</returns>
    Task<IReadOnlyList<ExcelImportError>> ValidateAsync(
        ExcelBatchValidationContext<T> context,
        CancellationToken cancellationToken = default);
}

/// <summary>
/// Provides immutable worksheet context to an <see cref="IExcelBatchValidator{T}"/>.
/// </summary>
/// <typeparam name="T">The mapped row model type.</typeparam>
public sealed class ExcelBatchValidationContext<T>
{
    internal ExcelBatchValidationContext(
        string sheetName,
        IReadOnlyList<ExcelImportRow<T>> rows)
    {
        SheetName = sheetName;
        Rows = rows;
    }

    /// <summary>Gets the worksheet name being validated.</summary>
    public string SheetName { get; }

    /// <summary>Gets all parsed rows, including rows that already contain errors.</summary>
    public IReadOnlyList<ExcelImportRow<T>> Rows { get; }
}

/// <summary>
/// Represents a parsed Excel row together with its physical position and current validation state.
/// </summary>
/// <typeparam name="T">The mapped row model type.</typeparam>
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

    /// <summary>Gets the physical worksheet row number.</summary>
    public int RowNumber { get; }

    /// <summary>Gets the mapped row value.</summary>
    public T Value { get; }

    /// <summary>Gets the mapping and validation errors currently associated with the row.</summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }

    /// <summary>Gets whether the row currently contains one or more errors.</summary>
    public bool HasErrors => Errors.Count > 0;
}
