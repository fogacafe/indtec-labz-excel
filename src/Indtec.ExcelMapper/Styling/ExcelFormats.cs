namespace Indtec.ExcelMapper.Styling;

/// <summary>Common Excel number-format strings for dates, numbers, percentages and currencies.</summary>
public static class ExcelFormats
{
    /// <summary>Brazilian date format.</summary>
    public const string DateBrazil = "dd/mm/yyyy";
    /// <summary>ISO-style date format.</summary>
    public const string DateIso = "yyyy-mm-dd";
    /// <summary>Brazilian date and time format.</summary>
    public const string DateTimeBrazil = "dd/mm/yyyy hh:mm:ss";
    /// <summary>ISO-style date and time format.</summary>
    public const string DateTimeIso = "yyyy-mm-dd hh:mm:ss";
    /// <summary>integer number format.</summary>
    public const string Integer = "0";
    /// <summary>decimal format with two fractional digits.</summary>
    public const string Decimal2 = "#,##0.00";
    /// <summary>decimal format with four fractional digits.</summary>
    public const string Decimal4 = "#,##0.0000";
    /// <summary>percentage format with two fractional digits.</summary>
    public const string Percentage2 = "0.00%";
    /// <summary>Brazilian real currency format.</summary>
    public const string CurrencyBrazil = "R$ #,##0.00";
    /// <summary>US dollar currency format.</summary>
    public const string CurrencyUs = "$#,##0.00";
}
