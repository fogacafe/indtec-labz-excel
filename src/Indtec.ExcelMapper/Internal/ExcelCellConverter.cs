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
        => Read(ToExcelValue(cell), destinationType, cell.Address.ToString());

    public static object? Read(ExcelValue value, Type destinationType, string? address = null)
    {
        var targetType = Nullable.GetUnderlyingType(destinationType) ?? destinationType;
        var source = value.Value;
        var location = string.IsNullOrWhiteSpace(address) ? "cell" : $"cell {address}";

        if (source is null || source is string text && string.IsNullOrEmpty(text))
        {
            if (Nullable.GetUnderlyingType(destinationType) is not null || !destinationType.IsValueType)
                return null;

            throw new ExcelMappingException($"{location} is empty but '{destinationType.Name}' is not nullable.");
        }

        try
        {
            if (targetType == typeof(string))
                return Convert.ToString(source, CultureInfo.InvariantCulture) ?? string.Empty;

            if (targetType == typeof(DateTime))
            {
                if (source is DateTime dateTime) return dateTime;
                if (TryToDouble(source, out var serialDate)) return DateTime.FromOADate(serialDate);
                return DateTime.Parse(Convert.ToString(source, CultureInfo.InvariantCulture)!, CultureInfo.InvariantCulture);
            }

            if (targetType == typeof(TimeSpan))
            {
                if (source is TimeSpan timeSpan) return timeSpan;
                if (TryToDouble(source, out var serialTime)) return TimeSpan.FromDays(serialTime);
                return TimeSpan.Parse(Convert.ToString(source, CultureInfo.InvariantCulture)!, CultureInfo.InvariantCulture);
            }

            if (targetType == typeof(Guid))
                return Guid.Parse(Convert.ToString(source, CultureInfo.InvariantCulture)!);

            if (targetType.IsEnum)
                return Enum.Parse(targetType, Convert.ToString(source, CultureInfo.InvariantCulture)!, ignoreCase: true);

            return Convert.ChangeType(source, targetType, CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (ex is not ExcelMappingException)
        {
            throw new ExcelMappingException(
                $"Could not convert {location} to '{destinationType.Name}'.", ex);
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

    private static bool TryToDouble(object source, out double value)
    {
        switch (source)
        {
            case double x: value = x; return true;
            case float x: value = x; return true;
            case decimal x: value = (double)x; return true;
            case byte x: value = x; return true;
            case short x: value = x; return true;
            case int x: value = x; return true;
            case long x: value = x; return true;
            default:
                return double.TryParse(
                    Convert.ToString(source, CultureInfo.InvariantCulture),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out value);
        }
    }
}
