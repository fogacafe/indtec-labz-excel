namespace Indtec.ExcelMapper;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public sealed class ExcelSheetAttribute : Attribute
{
    public ExcelSheetAttribute(string name)
    {
        Name = string.IsNullOrWhiteSpace(name)
            ? throw new ArgumentException("Sheet name cannot be empty.", nameof(name))
            : name;
    }

    public string Name { get; }
}
