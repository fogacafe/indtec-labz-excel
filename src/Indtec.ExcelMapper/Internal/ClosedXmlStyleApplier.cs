using ClosedXML.Excel;
using Indtec.ExcelMapper.Styling;

namespace Indtec.ExcelMapper.Internal;

internal static class ClosedXmlStyleApplier
{
    public static void Apply(IXLStyle target, ExcelStyle style)
    {
        if (style.Bold.HasValue) target.Font.Bold = style.Bold.Value;
        if (style.Italic.HasValue) target.Font.Italic = style.Italic.Value;
        if (style.FontSize.HasValue) target.Font.FontSize = style.FontSize.Value;

        var fontColor = style.FontColor;
        if (!string.IsNullOrWhiteSpace(fontColor))
            target.Font.FontColor = XLColor.FromHtml(fontColor!);

        var background = style.Background;
        if (!string.IsNullOrWhiteSpace(background))
            target.Fill.BackgroundColor = XLColor.FromHtml(background!);

        if (!string.IsNullOrWhiteSpace(style.NumberFormat)) target.NumberFormat.Format = style.NumberFormat;
        if (style.WrapText.HasValue) target.Alignment.WrapText = style.WrapText.Value;

        if (style.HorizontalAlignment.HasValue)
        {
            target.Alignment.Horizontal = style.HorizontalAlignment.Value switch
            {
                ExcelHorizontalAlignment.Left => XLAlignmentHorizontalValues.Left,
                ExcelHorizontalAlignment.Center => XLAlignmentHorizontalValues.Center,
                ExcelHorizontalAlignment.Right => XLAlignmentHorizontalValues.Right,
                _ => XLAlignmentHorizontalValues.General
            };
        }

        if (style.Border == true)
        {
            target.Border.TopBorder = XLBorderStyleValues.Thin;
            target.Border.BottomBorder = XLBorderStyleValues.Thin;
            target.Border.LeftBorder = XLBorderStyleValues.Thin;
            target.Border.RightBorder = XLBorderStyleValues.Thin;
        }
    }
}
