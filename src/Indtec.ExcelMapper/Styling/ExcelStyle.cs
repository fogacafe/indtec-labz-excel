namespace Indtec.ExcelMapper.Styling;

public enum ExcelHorizontalAlignment
{
    General,
    Left,
    Center,
    Right
}

public sealed class ExcelStyle
{
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public double? FontSize { get; set; }
    public string? FontColor { get; set; }
    public string? Background { get; set; }
    public string? NumberFormat { get; set; }
    public ExcelHorizontalAlignment? HorizontalAlignment { get; set; }
    public bool? WrapText { get; set; }
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
