using System.Linq.Expressions;

namespace Indtec.ExcelMapper.Importing;

/// <summary>Controls whether import errors are thrown immediately or collected in the result.</summary>
public enum ExcelImportErrorBehavior
{
    /// <summary>Throws when an import error is encountered.</summary>
    Throw,
    /// <summary>Collects import errors and continues when possible.</summary>
    Collect
}

/// <summary>Controls how empty worksheet rows are handled during import.</summary>
public enum ExcelEmptyRowBehavior
{
    /// <summary>Skips empty rows.</summary>
    Ignore,
    /// <summary>Includes empty rows as default model instances.</summary>
    Include,
    /// <summary>Reports empty rows as import errors.</summary>
    Error
}

/// <summary>Configures mapping, normalization and validation for an import operation.</summary>
public sealed class ExcelImportOptions<T>
{
    internal List<ExcelRowValidationRule<T>> Validators { get; } = new();
    internal List<IExcelBatchValidator<T>> BatchValidators { get; } = new();
    internal Dictionary<string, Func<object?, object?>> Normalizers { get; } = new(StringComparer.Ordinal);

    /// <summary>Gets or sets how mapping and validation errors are handled.</summary>
    public ExcelImportErrorBehavior ErrorBehavior { get; set; } = ExcelImportErrorBehavior.Throw;
    /// <summary>Gets or sets how empty worksheet rows are handled.</summary>
    public ExcelEmptyRowBehavior EmptyRowBehavior { get; set; } = ExcelEmptyRowBehavior.Ignore;
    /// <summary>Gets or sets whether a missing worksheet is allowed during workbook import.</summary>
    public bool OptionalSheet { get; set; }

    /// <summary>
    /// Gets or sets the maximum number of invalid worksheet rows collected before parsing stops.
    /// Null means no limit. The limit applies to invalid rows, not individual error messages.
    /// </summary>
    public int? MaxInvalidRows { get; set; }

    /// <summary>Adds a row validation rule. The row is valid when the predicate returns true.</summary>
    /// <summary>Adds a column validation rule. The value is valid when the predicate returns true.</summary>
    public ExcelImportOptions<T> Validate(Func<T, bool> predicate, string message)
        => AddValidation(predicate, null, message);

    /// <summary>Adds a row error when the predicate evaluates to true.</summary>
    public ExcelImportOptions<T> ErrorWhen(Func<T, bool> predicate, string message)
    {
        if (predicate is null) throw new ArgumentNullException(nameof(predicate));
        return AddValidation(row => !predicate(row), null, message);
    }

    /// <summary>Selects a mapped property for column-level normalization or validation.</summary>
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

    internal void AddNormalizer(string propertyName, Func<object?, object?> normalizer)
        => Normalizers[propertyName] = normalizer;

    internal void ValidateConfiguration()
    {
        if (MaxInvalidRows is <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxInvalidRows), "MaxInvalidRows must be greater than zero.");
    }

    /// <summary>Adds an asynchronous validator that receives the complete parsed worksheet. Requires asynchronous import.</summary>
    public ExcelImportOptions<T> AddBatchValidator(IExcelBatchValidator<T> validator)
    {
        if (validator is null) throw new ArgumentNullException(nameof(validator));
        BatchValidators.Add(validator);
        return this;
    }
}

/// <summary>Configures normalization and validation for a mapped property.</summary>
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

    /// <summary>
    /// Normalizes a non-null, non-empty column value after conversion and before validation.
    /// Null and blank string values are left unchanged.
    /// </summary>
    public ExcelImportColumnBuilder<T, TProperty> Normalize(Func<TProperty, TProperty> normalizer)
    {
        if (normalizer is null) throw new ArgumentNullException(nameof(normalizer));
        _options.AddNormalizer(_propertyName, value => normalizer((TProperty)value!));
        return this;
    }

    /// <summary>Adds a column error when the predicate evaluates to true.</summary>
    public ExcelImportOptions<T> ErrorWhen(Func<TProperty, bool> predicate, string message)
    {
        if (predicate is null) throw new ArgumentNullException(nameof(predicate));
        return _options.AddValidation(row => !predicate(_selector(row)), _propertyName, message);
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
