using Indtec.ExcelMapper.Importing;

namespace Indtec.ExcelMapper.Workbooks;

/// <summary>Validates relationships or rules that span multiple imported worksheets.</summary>
public interface IExcelWorkbookValidator
{
    /// <summary>Validates the imported workbook and returns additional sheet-specific errors.</summary>
    Task<IReadOnlyList<ExcelWorkbookValidationError>> ValidateAsync(
        ExcelWorkbookValidationContext context,
        CancellationToken cancellationToken = default);
}

/// <summary>Provides typed access to imported worksheet results during workbook validation.</summary>
public sealed class ExcelWorkbookValidationContext
{
    private readonly IReadOnlyDictionary<Type, object> _results;

    internal ExcelWorkbookValidationContext(IReadOnlyDictionary<Type, object> results)
        => _results = results;

    /// <summary>Gets the import result registered for the specified mapped model type.</summary>
    public ExcelImportResult<T> Sheet<T>()
    {
        if (_results.TryGetValue(typeof(T), out var result) && result is ExcelImportResult<T> typed)
            return typed;

        throw new ExcelMappingException($"Sheet result for '{typeof(T).Name}' was not registered.");
    }
}

/// <summary>Represents a validation error produced by a workbook-level validator.</summary>
public sealed class ExcelWorkbookValidationError
{
    private ExcelWorkbookValidationError(Type modelType, int row, string? column, string message)
    {
        ModelType = modelType;
        Row = row;
        Column = column;
        Message = message;
    }

    /// <summary>Gets the mapped model type whose worksheet contains the error.</summary>
    public Type ModelType { get; }
    /// <summary>Gets the physical worksheet row number.</summary>
    public int Row { get; }
    /// <summary>Gets the column associated with the error, when available.</summary>
    public string? Column { get; }
    /// <summary>Gets the human-readable validation message.</summary>
    public string Message { get; }

    /// <summary>Creates a workbook validation error for the worksheet mapped to <typeparamref name="T"/>.</summary>
    public static ExcelWorkbookValidationError For<T>(int row, string? column, string message)
        => new(typeof(T), row, column, message);
}
