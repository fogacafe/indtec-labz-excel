namespace Indtec.ExcelMapper;

/// <summary>
/// Associates a mapped model with an Excel worksheet name.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public sealed class ExcelSheetAttribute : Attribute
{
    /// <summary>
    /// Creates a worksheet mapping with the specified name.
    /// </summary>
    /// <param name="name">The worksheet name used for import, export and template generation.</param>
    public ExcelSheetAttribute(string name)
    {
        Name = string.IsNullOrWhiteSpace(name)
            ? throw new ArgumentException("Sheet name cannot be empty.", nameof(name))
            : name;
    }

    /// <summary>Gets the mapped worksheet name.</summary>
    public string Name { get; }
}
