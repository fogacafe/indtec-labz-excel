namespace Indtec.ExcelMapper.Conversion;

public interface IExcelValueConverter
{
    object? Read(ExcelValue value, Type destinationType);
    ExcelValue Write(object? value);
}

public readonly struct ExcelValue
{
    public ExcelValue(object? value) => Value = value;

    public object? Value { get; }
    public bool IsEmpty => Value is null;

    public string? AsString() => Value?.ToString();

    public static implicit operator ExcelValue(string? value) => new(value);
    public static implicit operator ExcelValue(bool value) => new(value);
    public static implicit operator ExcelValue(int value) => new(value);
    public static implicit operator ExcelValue(long value) => new(value);
    public static implicit operator ExcelValue(double value) => new(value);
    public static implicit operator ExcelValue(decimal value) => new(value);
    public static implicit operator ExcelValue(DateTime value) => new(value);
    public static implicit operator ExcelValue(TimeSpan value) => new(value);
}
