namespace Indtec.ExcelMapper.Styling;

/// <summary>Horizontal alignment options supported by exported cells.</summary>
public enum ExcelHorizontalAlignment
{
    /// <summary>Uses Excel's default alignment for the cell value.</summary>
    General,
    /// <summary>Aligns content to the left.</summary>
    Left,
    /// <summary>Centers content horizontally.</summary>
    Center,
    /// <summary>Aligns content to the right.</summary>
    Right
}

/// <summary>Describes cell style values that can be applied during export or template generation.</summary>
public sealed class ExcelStyle
{
    /// <summary>Gets or sets bold text.</summary>
    public bool? Bold { get; set; }
    /// <summary>Gets or sets italic text.</summary>
    public bool? Italic { get; set; }
    /// <summary>Gets or sets font size in points.</summary>
    public double? FontSize { get; set; }
    /// <summary>Gets or sets font color as a hexadecimal value.</summary>
    public string? FontColor { get; set; }
    /// <summary>Gets or sets background color as a hexadecimal value.</summary>
    public string? Background { get; set; }
    /// <summary>Gets or sets Excel number format string.</summary>
    public string? NumberFormat { get; set; }
    /// <summary>Gets or sets horizontal alignment.</summary>
    public ExcelHorizontalAlignment? HorizontalAlignment { get; set; }
    /// <summary>Gets or sets whether text wraps.</summary>
    public bool? WrapText { get; set; }
    /// <summary>Gets or sets whether a thin border is applied.</summary>
    public bool? Border { get; set; }

    internal void MergeFrom(ExcelStyle other)
    {
        Bold = other.Bold ?? Bold;
        Italic = other.Italic ?? Italic;
        FontSize = other.FontSize ?? FontSize;
        FontColor = other.FontColor ?? FontColor;
        Background = other.Background ?? Background;
        NumberFormat = other.NumberFormat ?? NumberFormat;
        HorizontalAlignment = other.HorizontalAlignment ?? HorizontalAlignment;
        WrapText = other.WrapText ?? WrapText;
        Border = other.Border ?? Border;
    }
}
