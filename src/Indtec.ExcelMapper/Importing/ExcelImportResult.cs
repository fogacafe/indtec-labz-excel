namespace Indtec.ExcelMapper.Importing;

/// <summary>
/// Contains the successfully imported items, parsed row state and structured errors for an Excel import.
/// </summary>
/// <typeparam name="T">The mapped row model type.</typeparam>
public sealed class ExcelImportResult<T>
{
    internal ExcelImportResult(
        IReadOnlyList<T> items,
        IReadOnlyList<ExcelImportError> errors,
        IReadOnlyList<ExcelImportRow<T>>? rows = null)
    {
        Items = items;
        Errors = errors;
        Rows = rows ?? Array.Empty<ExcelImportRow<T>>();
    }

    /// <summary>Gets successfully mapped and validated items.</summary>
    public IReadOnlyList<T> Items { get; }

    /// <summary>Gets all mapping and validation errors produced by the import.</summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }

    /// <summary>Gets every parsed row together with its physical row number and current errors.</summary>
    public IReadOnlyList<ExcelImportRow<T>> Rows { get; }

    /// <summary>Gets whether the import completed without mapping or validation errors.</summary>
    public bool IsValid => Errors.Count == 0;
}

/// <summary>
/// Describes a mapping or validation problem associated with an Excel row and optional column.
/// </summary>
public sealed class ExcelImportError
{
    /// <summary>Creates a structured Excel import error.</summary>
    /// <param name="row">The physical worksheet row number.</param>
    /// <param name="column">The mapped property or column associated with the error, when available.</param>
    /// <param name="message">The human-readable error message.</param>
    public ExcelImportError(int row, string? column, string message)
    {
        Row = row;
        Column = column;
        Message = message;
    }

    /// <summary>Gets the physical worksheet row number.</summary>
    public int Row { get; }

    /// <summary>Gets the mapped property or column associated with the error, when available.</summary>
    public string? Column { get; }

    /// <summary>Gets the human-readable error message.</summary>
    public string Message { get; }

    /// <summary>Returns a compact row/column representation of the error.</summary>
    public override string ToString()
        => Column is null
            ? $"Row {Row}: {Message}"
            : $"Row {Row}, column '{Column}': {Message}";
}
