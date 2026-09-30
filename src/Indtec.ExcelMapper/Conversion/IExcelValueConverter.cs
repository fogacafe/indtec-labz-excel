namespace Indtec.ExcelMapper.Conversion;

/// <summary>Converts values between Excel cells and mapped CLR properties.</summary>
public interface IExcelValueConverter
{
    /// <summary>Converts an Excel value to the requested destination type during import.</summary>
    object? Read(ExcelValue value, Type destinationType);

    /// <summary>Converts a CLR value to an Excel-compatible value during export.</summary>
    ExcelValue Write(object? value);
}

/// <summary>Represents a raw value exchanged with a custom Excel value converter.</summary>
public readonly struct ExcelValue
{
    /// <summary>Creates an Excel value wrapping the supplied CLR value.</summary>
    public ExcelValue(object? value) => Value = value;

    /// <summary>Gets the wrapped CLR value.</summary>
    public object? Value { get; }

    /// <summary>Gets whether the value is empty.</summary>
    public bool IsEmpty => Value is null;

    /// <summary>Returns the value as text using its standard string representation.</summary>
    public string? AsString() => Value?.ToString();

    /// <summary>Creates an Excel value from a string.</summary>
    public static implicit operator ExcelValue(string? value) => new(value);
    /// <summary>Creates an Excel value from a Boolean.</summary>
    public static implicit operator ExcelValue(bool value) => new(value);
    /// <summary>Creates an Excel value from a 32-bit integer.</summary>
    public static implicit operator ExcelValue(int value) => new(value);
    /// <summary>Creates an Excel value from a 64-bit integer.</summary>
    public static implicit operator ExcelValue(long value) => new(value);
    /// <summary>Creates an Excel value from a double-precision number.</summary>
    public static implicit operator ExcelValue(double value) => new(value);
    /// <summary>Creates an Excel value from a decimal number.</summary>
    public static implicit operator ExcelValue(decimal value) => new(value);
    /// <summary>Creates an Excel value from a date and time.</summary>
    public static implicit operator ExcelValue(DateTime value) => new(value);
    /// <summary>Creates an Excel value from a time interval.</summary>
    public static implicit operator ExcelValue(TimeSpan value) => new(value);
}
