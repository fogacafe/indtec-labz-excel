namespace Indtec.ExcelMapper.Importing;

public sealed class ExcelImportResult<T>
{
    internal ExcelImportResult(IReadOnlyList<T> items, IReadOnlyList<ExcelImportError> errors)
    {
        Items = items;
        Errors = errors;
    }

    public IReadOnlyList<T> Items { get; }
    public IReadOnlyList<ExcelImportError> Errors { get; }
    public bool IsValid => Errors.Count == 0;
}

public sealed class ExcelImportError
{
    public ExcelImportError(int row, string? column, string message)
    {
        Row = row;
        Column = column;
        Message = message;
    }

    public int Row { get; }
    public string? Column { get; }
    public string Message { get; }

    public override string ToString()
        => Column is null
            ? $"Row {Row}: {Message}"
            : $"Row {Row}, column '{Column}': {Message}";
}
