using System.Linq.Expressions;

namespace Indtec.ExcelMapper.Styling;

public sealed class ExcelExportOptions<T>
{
    private readonly Dictionary<string, ExcelColumnStyleConfig<T>> _columns =
        new(StringComparer.Ordinal);
    private readonly ExcelStyle _headerStyle = new();

    internal IReadOnlyDictionary<string, ExcelColumnStyleConfig<T>> Columns => _columns;
    internal List<ExcelConditionalStyleRule<T>> RowRules { get; } = new();
    internal ExcelStyle HeaderStyle => _headerStyle;

    public ExcelHeaderStyleBuilder Header => new(_headerStyle);
    public bool FreezeHeader { get; set; } = true;
    public bool AutoFilter { get; set; } = true;
    public int TemplateRows { get; set; } = 1000;

    public ExcelExportOptions<T> UseTheme(IExcelTheme<T> theme)
    {
        if (theme is null) throw new ArgumentNullException(nameof(theme));
        theme.Configure(this);
        return this;
    }

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

public class ExcelStyleBuilder<TBuilder> where TBuilder : ExcelStyleBuilder<TBuilder>
{
    private readonly ExcelStyle _style;

    internal ExcelStyleBuilder(ExcelStyle style) => _style = style;

    protected TBuilder Self => (TBuilder)this;

    public TBuilder Bold(bool value = true) { _style.Bold = value; return Self; }
    public TBuilder Italic(bool value = true) { _style.Italic = value; return Self; }
    public TBuilder FontSize(double value) { _style.FontSize = value; return Self; }
    public TBuilder FontColor(string hex) { _style.FontColor = hex; return Self; }
    public TBuilder Background(string hex) { _style.Background = hex; return Self; }
    public TBuilder NumberFormat(string format) { _style.NumberFormat = format; return Self; }
    public TBuilder Align(ExcelHorizontalAlignment alignment) { _style.HorizontalAlignment = alignment; return Self; }
    public TBuilder Wrap(bool value = true) { _style.WrapText = value; return Self; }
    public TBuilder Border(bool value = true) { _style.Border = value; return Self; }
}

public sealed class ExcelHeaderStyleBuilder : ExcelStyleBuilder<ExcelHeaderStyleBuilder>
{
    internal ExcelHeaderStyleBuilder(ExcelStyle style) : base(style) { }
}

public sealed class ExcelColumnStyleBuilder<T, TProperty> : ExcelStyleBuilder<ExcelColumnStyleBuilder<T, TProperty>>
{
    private readonly ExcelColumnStyleConfig<T> _config;

    internal ExcelColumnStyleBuilder(ExcelColumnStyleConfig<T> config) : base(config.Style)
        => _config = config;

    public ExcelColumnStyleBuilder<T, TProperty> Width(double width)
    {
        _config.Width = width;
        return this;
    }

    public ExcelColumnStyleBuilder<T, TProperty> AllowedValues(params string[] values)
    {
        if (values is null) throw new ArgumentNullException(nameof(values));
        if (values.Length == 0) throw new ArgumentException("At least one allowed value is required.", nameof(values));
        if (values.Any(string.IsNullOrWhiteSpace)) throw new ArgumentException("Allowed values cannot contain empty values.", nameof(values));

        _config.AllowedValues = values.Distinct(StringComparer.Ordinal).ToArray();
        return this;
    }

    public ExcelConditionalColumnStyleBuilder<T, TProperty> When(Func<T, bool> predicate)
    {
        var rule = new ExcelConditionalStyleRule<T>(predicate);
        _config.Rules.Add(rule);
        return new ExcelConditionalColumnStyleBuilder<T, TProperty>(_config, rule);
    }
}

public sealed class ExcelConditionalColumnStyleBuilder<T, TProperty> : ExcelStyleBuilder<ExcelConditionalColumnStyleBuilder<T, TProperty>>
{
    private readonly ExcelColumnStyleConfig<T> _config;

    internal ExcelConditionalColumnStyleBuilder(
        ExcelColumnStyleConfig<T> config,
        ExcelConditionalStyleRule<T> rule) : base(rule.Style)
        => _config = config;

    public ExcelConditionalColumnStyleBuilder<T, TProperty> When(Func<T, bool> predicate)
    {
        var rule = new ExcelConditionalStyleRule<T>(predicate);
        _config.Rules.Add(rule);
        return new ExcelConditionalColumnStyleBuilder<T, TProperty>(_config, rule);
    }
}

public sealed class ExcelRowStyleBuilder<T>
{
    private readonly List<ExcelConditionalStyleRule<T>> _rules;

    internal ExcelRowStyleBuilder(List<ExcelConditionalStyleRule<T>> rules) => _rules = rules;

    public ExcelConditionalRowStyleBuilder<T> When(Func<T, bool> predicate)
    {
        var rule = new ExcelConditionalStyleRule<T>(predicate);
        _rules.Add(rule);
        return new ExcelConditionalRowStyleBuilder<T>(_rules, rule);
    }
}

public sealed class ExcelConditionalRowStyleBuilder<T> : ExcelStyleBuilder<ExcelConditionalRowStyleBuilder<T>>
{
    private readonly List<ExcelConditionalStyleRule<T>> _rules;

    internal ExcelConditionalRowStyleBuilder(
        List<ExcelConditionalStyleRule<T>> rules,
        ExcelConditionalStyleRule<T> rule) : base(rule.Style)
        => _rules = rules;

    public ExcelConditionalRowStyleBuilder<T> When(Func<T, bool> predicate)
    {
        var rule = new ExcelConditionalStyleRule<T>(predicate);
        _rules.Add(rule);
        return new ExcelConditionalRowStyleBuilder<T>(_rules, rule);
    }
}
