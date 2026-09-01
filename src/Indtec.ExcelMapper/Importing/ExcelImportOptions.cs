namespace Indtec.ExcelMapper.Importing;

public enum ExcelImportErrorBehavior
{
    Throw,
    Collect
}

public sealed class ExcelImportOptions<T>
{
    internal List<ExcelRowValidationRule<T>> Validators { get; } = new();
    internal List<IExcelBatchValidator<T>> BatchValidators { get; } = new();

    public ExcelImportErrorBehavior ErrorBehavior { get; set; } = ExcelImportErrorBehavior.Throw;

    public ExcelImportOptions<T> Validate(Func<T, bool> predicate, string message)
    {
        if (predicate is null) throw new ArgumentNullException(nameof(predicate));
        if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException("Validation message cannot be empty.", nameof(message));

        Validators.Add(new ExcelRowValidationRule<T>(predicate, message));
        return this;
    }

    public ExcelImportOptions<T> AddBatchValidator(IExcelBatchValidator<T> validator)
    {
        if (validator is null) throw new ArgumentNullException(nameof(validator));
        BatchValidators.Add(validator);
        return this;
    }
}

internal sealed class ExcelRowValidationRule<T>
{
    public ExcelRowValidationRule(Func<T, bool> predicate, string message)
    {
        Predicate = predicate;
        Message = message;
    }

    public Func<T, bool> Predicate { get; }
    public string Message { get; }
}
