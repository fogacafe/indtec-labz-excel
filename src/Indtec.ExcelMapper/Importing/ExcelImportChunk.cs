namespace Indtec.ExcelMapper.Importing;

/// <summary>
/// Represents one bounded group of rows produced by a streaming worksheet import.
/// </summary>
/// <typeparam name="T">The mapped row model type.</typeparam>
public sealed class ExcelImportChunk<T>
{
    internal ExcelImportChunk(
        int index,
        int startRow,
        int endRow,
        IReadOnlyList<T> items,
        IReadOnlyList<ExcelImportError> errors,
        IReadOnlyList<ExcelImportRow<T>> rows)
    {
        Index = index;
        StartRow = startRow;
        EndRow = endRow;
        Items = items;
        Errors = errors;
        Rows = rows;
    }

    /// <summary>Gets the zero-based chunk index.</summary>
    public int Index { get; }

    /// <summary>Gets the first physical worksheet row represented by this chunk.</summary>
    public int StartRow { get; }

    /// <summary>Gets the last physical worksheet row represented by this chunk.</summary>
    public int EndRow { get; }

    /// <summary>Gets successfully mapped and validated items in this chunk.</summary>
    public IReadOnlyList<T> Items { get; }

    /// <summary>Gets all mapping and validation errors produced for this chunk.</summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }

    /// <summary>Gets every parsed row together with its physical row number and current errors.</summary>
    public IReadOnlyList<ExcelImportRow<T>> Rows { get; }

    /// <summary>Gets whether this chunk contains no mapping or validation errors.</summary>
    public bool IsValid => Errors.Count == 0;
}
