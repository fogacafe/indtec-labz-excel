namespace Indtec.ExcelMapper.Importing;

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

    public int Index { get; }
    public int StartRow { get; }
    public int EndRow { get; }
    public IReadOnlyList<T> Items { get; }
    public IReadOnlyList<ExcelImportError> Errors { get; }
    public IReadOnlyList<ExcelImportRow<T>> Rows { get; }
    public bool IsValid => Errors.Count == 0;
}
