using System.Linq.Expressions;

namespace Indtec.ExcelMapper.Importing;

public enum ExcelImportErrorBehavior
{
    Throw,
    Collect
}

public enum ExcelEmptyRowBehavior
{
    Ignore,
    Include,
    Error
}

public sealed class ExcelImportOptions<T>
{
    internal List<ExcelRowValidationRule<T>> Validators { get; } = new();
    internal List<IExcelBatchValidator<T>> BatchValidators { get; } = new();

    public ExcelImportErrorBehavior ErrorBehavior { get; set; } = ExcelImportErrorBehavior.Throw;
    public ExcelEmptyRowBehavior EmptyRowBehavior { get; set; } = ExcelEmptyRowBehavior.Ignore;
    public bool OptionalSheet { get; set; }

    public ExcelImportOptions<T> Validate(Func<T, bool> predicate, string message)
        => AddValidation(predicate, null, message);

    public ExcelImportColumnBuilder<T, TProperty> Column<TProperty>(Expression<Func<T, TProperty>> selector)
    {
        if (selector is null) throw new ArgumentNullException(nameof(selector));
        if (selector.Body is not MemberExpression member)
            throw new ArgumentException("Column selector must point directly to a property.", nameof(selector));

        return new ExcelImportColumnBuilder<T, TProperty>(this, member.Member.Name, selector.Compile());
    }

    internal ExcelImportOptions<T> AddValidation(Func<T, bool> predicate, string? propertyName, string message)
    {
        if (predicate is null) throw new ArgumentNullException(nameof(predicate));
        if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException("Validation message cannot be empty.", nameof(message));

        Validators.Add(new ExcelRowValidationRule<T>(predicate, propertyName, message));
        return this;
    }

    public ExcelImportOptions<T> AddBatchValidator(IExcelBatchValidator<T> validator)
    {
        if (validator is null) throw new ArgumentNullException(nameof(validator));
        BatchValidators.Add(validator);
        return this;
    }
}

public sealed class ExcelImportColumnBuilder<T, TProperty>
{
    private readonly ExcelImportOptions<T> _options;
    private readonly string _propertyName;
    private readonly Func<T, TProperty> _selector;

    internal ExcelImportColumnBuilder(
        ExcelImportOptions<T> options,
        string propertyName,
        Func<T, TProperty> selector)
    {
        _options = options;
        _propertyName = propertyName;
        _selector = selector;
    }

    public ExcelImportOptions<T> Validate(Func<TProperty, bool> predicate, string message)
    {
        if (predicate is null) throw new ArgumentNullException(nameof(predicate));
        return _options.AddValidation(row => predicate(_selector(row)), _propertyName, message);
    }
}

internal sealed class ExcelRowValidationRule<T>
{
    public ExcelRowValidationRule(Func<T, bool> predicate, string? propertyName, string message)
    {
        Predicate = predicate;
        PropertyName = propertyName;
        Message = message;
    }

    public Func<T, bool> Predicate { get; }
    public string? PropertyName { get; }
    public string Message { get; }
}
