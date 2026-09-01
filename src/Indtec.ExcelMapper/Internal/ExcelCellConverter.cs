using System.Globalization;
using ClosedXML.Excel;
using Indtec.ExcelMapper.Conversion;

namespace Indtec.ExcelMapper.Internal;

internal static class ExcelCellConverter
{
    public static ExcelValue ToExcelValue(IXLCell cell)
    {
        if (cell.IsEmpty()) return new ExcelValue(null);

        return cell.DataType switch
        {
            XLDataType.Boolean => new ExcelValue(cell.GetBoolean()),
            XLDataType.Number => new ExcelValue(cell.GetDouble()),
            XLDataType.DateTime => new ExcelValue(cell.GetDateTime()),
            XLDataType.TimeSpan => new ExcelValue(cell.GetTimeSpan()),
            _ => new ExcelValue(cell.GetString())
        };
    }

    public static object? Read(IXLCell cell, Type destinationType)
    {
        var targetType = Nullable.GetUnderlyingType(destinationType) ?? destinationType;

        if (cell.IsEmpty())
        {
            if (Nullable.GetUnderlyingType(destinationType) is not null || !destinationType.IsValueType)
                return null;

            throw new ExcelMappingException($"Cell {cell.Address} is empty but '{destinationType.Name}' is not nullable.");
        }

        try
        {
            if (targetType == typeof(string)) return cell.GetString();
            if (targetType == typeof(bool)) return cell.GetValue<bool>();
            if (targetType == typeof(byte)) return cell.GetValue<byte>();
            if (targetType == typeof(short)) return cell.GetValue<short>();
            if (targetType == typeof(int)) return cell.GetValue<int>();
            if (targetType == typeof(long)) return cell.GetValue<long>();
            if (targetType == typeof(float)) return cell.GetValue<float>();
            if (targetType == typeof(double)) return cell.GetValue<double>();
            if (targetType == typeof(decimal)) return cell.GetValue<decimal>();
            if (targetType == typeof(DateTime)) return cell.GetValue<DateTime>();
            if (targetType == typeof(TimeSpan)) return cell.GetValue<TimeSpan>();
            if (targetType == typeof(Guid)) return Guid.Parse(cell.GetString());
            if (targetType.IsEnum) return Enum.Parse(targetType, cell.GetString(), ignoreCase: true);

            return Convert.ChangeType(cell.GetString(), targetType, CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (ex is not ExcelMappingException)
        {
            throw new ExcelMappingException(
                $"Could not convert cell {cell.Address} to '{destinationType.Name}'.", ex);
        }
    }

    public static void Write(IXLCell cell, object? value)
    {
        Write(cell, new ExcelValue(value));
    }

    public static void Write(IXLCell cell, ExcelValue value)
    {
        switch (value.Value)
        {
            case null: cell.Clear(); break;
            case string x: cell.Value = x; break;
            case bool x: cell.Value = x; break;
            case byte x: cell.Value = x; break;
            case short x: cell.Value = x; break;
            case int x: cell.Value = x; break;
            case long x: cell.Value = x; break;
            case float x: cell.Value = x; break;
            case double x: cell.Value = x; break;
            case decimal x: cell.Value = x; break;
            case DateTime x: cell.Value = x; break;
            case TimeSpan x: cell.Value = x; break;
            case Guid x: cell.Value = x.ToString(); break;
            case Enum x: cell.Value = x.ToString(); break;
            default: cell.Value = value.Value.ToString() ?? string.Empty; break;
        }
    }
}
