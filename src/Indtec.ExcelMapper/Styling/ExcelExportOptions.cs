using System.Linq.Expressions;

namespace Indtec.ExcelMapper.Styling;

/// <summary>Configures worksheet export and template generation for a mapped model.</summary>
/// <typeparam name="T">The mapped model type.</typeparam>
public sealed class ExcelExportOptions<T>
{
    private readonly Dictionary<string, ExcelColumnStyleConfig<T>> _columns =
        new(StringComparer.Ordinal);
    private readonly ExcelStyle _headerStyle = new();

    internal IReadOnlyDictionary<string, ExcelColumnStyleConfig<T>> Columns => _columns;
    internal List<ExcelConditionalStyleRule<T>> RowRules { get; } = new();
    internal ExcelStyle HeaderStyle => _headerStyle;

    /// <summary>Gets the fluent style builder for the worksheet header row.</summary>
    public ExcelHeaderStyleBuilder Header => new(_headerStyle);
    /// <summary>Gets or sets an optional worksheet name override for this operation.</summary>
    public string? SheetName { get; set; }
    /// <summary>Gets or sets whether the header row is frozen. Defaults to true.</summary>
    public bool FreezeHeader { get; set; } = true;
    /// <summary>Gets or sets whether an Excel auto-filter is applied to the used range. Defaults to true.</summary>
    public bool AutoFilter { get; set; } = true;
    /// <summary>Auto-sizes columns using the header row only.</summary>
    public bool AutoFitHeaders { get; set; }

    /// <summary>
    /// Auto-sizes columns using headers and exported/template data. Explicit Width(...) always wins.
    /// </summary>
    public bool AutoFitColumns { get; set; }
    /// <summary>Gets or sets the number of data rows prepared in generated templates. This determines the range where features such as dropdown validation are applied. Defaults to 1000.</summary>
    public int TemplateRows { get; set; } = 1000;

    /// <summary>Applies a reusable theme to these export options.</summary>
    public ExcelExportOptions<T> UseTheme(IExcelTheme<T> theme)
    {
        if (theme is null) throw new ArgumentNullException(nameof(theme));
        theme.Configure(this);
        return this;
    }

    /// <summary>Configures styling, validation and conditional formatting for a mapped column.</summary>
    public ExcelColumnStyleBuilder<T, TProperty> Column<TProperty>(Expression<Func<T, TProperty>> selector)
    {
        if (selector.Body is not MemberExpression member)
            throw new ArgumentException("Column selector must point directly to a property.", nameof(selector));

        var propertyName = member.Member.Name;
        if (!_columns.TryGetValue(propertyName, out var config))
        {
            config = new ExcelColumnStyleConfig<T>(propertyName);
            _columns[propertyName] = config;
        }

        return new ExcelColumnStyleBuilder<T, TProperty>(config);
    }

    /// <summary>Starts configuration of conditional styles that apply to entire exported rows.</summary>
    public ExcelRowStyleBuilder<T> Row() => new(RowRules);
}

internal sealed class ExcelColumnStyleConfig<T>
{
    public ExcelColumnStyleConfig(string propertyName) => PropertyName = propertyName;

    public string PropertyName { get; }
    public double? Width { get; set; }
    public IReadOnlyList<string>? AllowedValues { get; set; }
    public ExcelStyle Style { get; } = new();
    public List<ExcelConditionalStyleRule<T>> Rules { get; } = new();
}

internal sealed class ExcelConditionalStyleRule<T>
{
    public ExcelConditionalStyleRule(Func<T, bool> predicate) => Predicate = predicate;

    public Func<T, bool> Predicate { get; }
    public ExcelStyle Style { get; } = new();
}

/// <summary>Provides fluent style configuration shared by headers, columns and conditional rules.</summary>
public class ExcelStyleBuilder<TBuilder> where TBuilder : ExcelStyleBuilder<TBuilder>
{
    private readonly ExcelStyle _style;

    internal ExcelStyleBuilder(ExcelStyle style) => _style = style;

    protected TBuilder Self => (TBuilder)this;

    /// <summary>Sets whether text is bold.</summary>
    public TBuilder Bold(bool value = true) { _style.Bold = value; return Self; }
    /// <summary>Sets whether text is italic.</summary>
    public TBuilder Italic(bool value = true) { _style.Italic = value; return Self; }
    /// <summary>Sets the font size in points.</summary>
    public TBuilder FontSize(double value) { _style.FontSize = value; return Self; }
    /// <summary>Sets the font color using a hexadecimal color value.</summary>
    public TBuilder FontColor(string hex) { _style.FontColor = hex; return Self; }
    /// <summary>Sets the cell background using a hexadecimal color value.</summary>
    public TBuilder Background(string hex) { _style.Background = hex; return Self; }
    /// <summary>Sets the Excel number format string.</summary>
    public TBuilder NumberFormat(string format) { _style.NumberFormat = format; return Self; }
    /// <summary>Sets the Excel date/time format string.</summary>
    public TBuilder DateFormat(string format) { _style.NumberFormat = format; return Self; }
    /// <summary>Sets horizontal cell alignment.</summary>
    public TBuilder Align(ExcelHorizontalAlignment alignment) { _style.HorizontalAlignment = alignment; return Self; }
    /// <summary>Sets whether cell text wraps.</summary>
    public TBuilder Wrap(bool value = true) { _style.WrapText = value; return Self; }
    /// <summary>Sets whether a thin border is applied around cells.</summary>
    public TBuilder Border(bool value = true) { _style.Border = value; return Self; }
}

/// <summary>Configures styles for the worksheet header row.</summary>
public sealed class ExcelHeaderStyleBuilder : ExcelStyleBuilder<ExcelHeaderStyleBuilder>
{
    internal ExcelHeaderStyleBuilder(ExcelStyle style) : base(style) { }
}

/// <summary>Configures a mapped column and its conditional styles.</summary>
public sealed class ExcelColumnStyleBuilder<T, TProperty> : ExcelStyleBuilder<ExcelColumnStyleBuilder<T, TProperty>>
{
    private readonly ExcelColumnStyleConfig<T> _config;

    internal ExcelColumnStyleBuilder(ExcelColumnStyleConfig<T> config) : base(config.Style)
        => _config = config;

    /// <summary>Sets the Excel column width. Explicit width takes precedence over auto-fit.</summary>
    public ExcelColumnStyleBuilder<T, TProperty> Width(double width)
    {
        _config.Width = width;
        return this;
    }

    /// <summary>Defines values offered by Excel data validation when generating a template.</summary>
    public ExcelColumnStyleBuilder<T, TProperty> AllowedValues(params string[] values)
    {
        if (values is null) throw new ArgumentNullException(nameof(values));
        if (values.Length == 0) throw new ArgumentException("At least one allowed value is required.", nameof(values));
        if (values.Any(string.IsNullOrWhiteSpace)) throw new ArgumentException("Allowed values cannot contain empty values.", nameof(values));

        _config.AllowedValues = values.Distinct(StringComparer.Ordinal).ToArray();
        return this;
    }

    /// <summary>Adds a conditional style rule applied when the row predicate is true.</summary>
    public ExcelConditionalColumnStyleBuilder<T, TProperty> When(Func<T, bool> predicate)
    {
        var rule = new ExcelConditionalStyleRule<T>(predicate);
        _config.Rules.Add(rule);
        return new ExcelConditionalColumnStyleBuilder<T, TProperty>(_config, rule);
    }
}

/// <summary>Configures a conditional style for a mapped column.</summary>
public sealed class ExcelConditionalColumnStyleBuilder<T, TProperty> : ExcelStyleBuilder<ExcelConditionalColumnStyleBuilder<T, TProperty>>
{
    private readonly ExcelColumnStyleConfig<T> _config;

    internal ExcelConditionalColumnStyleBuilder(
        ExcelColumnStyleConfig<T> config,
        ExcelConditionalStyleRule<T> rule) : base(rule.Style)
        => _config = config;

    /// <summary>Starts another conditional style rule for the same column.</summary>
    public ExcelConditionalColumnStyleBuilder<T, TProperty> When(Func<T, bool> predicate)
    {
        var rule = new ExcelConditionalStyleRule<T>(predicate);
        _config.Rules.Add(rule);
        return new ExcelConditionalColumnStyleBuilder<T, TProperty>(_config, rule);
    }
}

/// <summary>Configures conditional styles that apply to entire rows.</summary>
public sealed class ExcelRowStyleBuilder<T>
{
    private readonly List<ExcelConditionalStyleRule<T>> _rules;

    internal ExcelRowStyleBuilder(List<ExcelConditionalStyleRule<T>> rules) => _rules = rules;

    /// <summary>Adds a row style rule applied when the predicate is true.</summary>
    public ExcelConditionalRowStyleBuilder<T> When(Func<T, bool> predicate)
    {
        var rule = new ExcelConditionalStyleRule<T>(predicate);
        _rules.Add(rule);
        return new ExcelConditionalRowStyleBuilder<T>(_rules, rule);
    }
}

/// <summary>Configures a conditional style applied to an entire row.</summary>
public sealed class ExcelConditionalRowStyleBuilder<T> : ExcelStyleBuilder<ExcelConditionalRowStyleBuilder<T>>
{
    private readonly List<ExcelConditionalStyleRule<T>> _rules;

    internal ExcelConditionalRowStyleBuilder(
        List<ExcelConditionalStyleRule<T>> rules,
        ExcelConditionalStyleRule<T> rule) : base(rule.Style)
        => _rules = rules;

    /// <summary>Starts another conditional style rule for rows.</summary>
    public ExcelConditionalRowStyleBuilder<T> When(Func<T, bool> predicate)
    {
        var rule = new ExcelConditionalStyleRule<T>(predicate);
        _rules.Add(rule);
        return new ExcelConditionalRowStyleBuilder<T>(_rules, rule);
    }
}
