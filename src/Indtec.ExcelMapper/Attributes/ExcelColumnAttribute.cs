namespace Indtec.ExcelMapper;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ExcelColumnAttribute : Attribute
{
    public ExcelColumnAttribute(string header)
    {
        Header = string.IsNullOrWhiteSpace(header)
            ? throw new ArgumentException("Column header cannot be empty.", nameof(header))
            : header;
    }

    public string Header { get; }
    public int Order { get; set; } = int.MaxValue;
    public bool Required { get; set; }
    public Type? Converter { get; set; }
}
