namespace Indtec.ExcelMapper;

/// <summary>
/// Maps a model property to a column in an Excel worksheet.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ExcelColumnAttribute : Attribute
{
    /// <summary>
    /// Creates a column mapping for the specified worksheet header.
    /// </summary>
    /// <param name="header">The exact header text used to identify the column.</param>
    public ExcelColumnAttribute(string header)
    {
        Header = string.IsNullOrWhiteSpace(header)
            ? throw new ArgumentException("Column header cannot be empty.", nameof(header))
            : header;
    }

    /// <summary>Gets the worksheet header associated with the property.</summary>
    public string Header { get; }

    /// <summary>Gets or sets the column order used when exporting or generating templates.</summary>
    public int Order { get; set; } = int.MaxValue;

    /// <summary>Gets or sets whether the column must exist when importing a worksheet.</summary>
    public bool Required { get; set; }

    /// <summary>Gets or sets an optional <see cref="Conversion.IExcelValueConverter"/> implementation used for this property.</summary>
    public Type? Converter { get; set; }
}
