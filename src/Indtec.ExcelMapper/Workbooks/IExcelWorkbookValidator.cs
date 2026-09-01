using Indtec.ExcelMapper.Importing;

namespace Indtec.ExcelMapper.Workbooks;

public interface IExcelWorkbookValidator
{
    Task<IReadOnlyList<ExcelWorkbookValidationError>> ValidateAsync(
        ExcelWorkbookValidationContext context,
        CancellationToken cancellationToken = default);
}

public sealed class ExcelWorkbookValidationContext
{
    private readonly IReadOnlyDictionary<Type, object> _results;

    internal ExcelWorkbookValidationContext(IReadOnlyDictionary<Type, object> results)
        => _results = results;

    public ExcelImportResult<T> Sheet<T>()
    {
        if (_results.TryGetValue(typeof(T), out var result) && result is ExcelImportResult<T> typed)
            return typed;

        throw new ExcelMappingException($"Sheet result for '{typeof(T).Name}' was not registered.");
    }
}

public sealed class ExcelWorkbookValidationError
{
    private ExcelWorkbookValidationError(Type modelType, int row, string? column, string message)
    {
        ModelType = modelType;
        Row = row;
        Column = column;
        Message = message;
    }

    public Type ModelType { get; }
    public int Row { get; }
    public string? Column { get; }
    public string Message { get; }

    public static ExcelWorkbookValidationError For<T>(int row, string? column, string message)
        => new(typeof(T), row, column, message);
}
