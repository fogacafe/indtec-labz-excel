using System.ComponentModel;

namespace Indtec.ExcelMapper.Mapping;

public sealed class ExcelColumnMap
{
    public ExcelColumnMap(
        string propertyName,
        string header,
        int order,
        Type valueType,
        Func<object, object?> getter,
        Action<object, object?>? setter)
    {
        PropertyName = propertyName;
        Header = header;
        Order = order;
        ValueType = valueType;
        Getter = getter;
        Setter = setter;
    }

    public string PropertyName { get; }
    public string Header { get; }
    public int Order { get; }
    public Type ValueType { get; }
    public Func<object, object?> Getter { get; }
    public Action<object, object?>? Setter { get; }
}

public sealed class ExcelTypeMap
{
    public ExcelTypeMap(string sheetName, IReadOnlyList<ExcelColumnMap> columns)
    {
        SheetName = sheetName;
        Columns = columns.OrderBy(x => x.Order).ToArray();
    }

    public string SheetName { get; }
    public IReadOnlyList<ExcelColumnMap> Columns { get; }
}

[EditorBrowsable(EditorBrowsableState.Never)]
public interface IGeneratedExcelMapProvider
{
    ExcelTypeMap ExcelMap { get; }
}
