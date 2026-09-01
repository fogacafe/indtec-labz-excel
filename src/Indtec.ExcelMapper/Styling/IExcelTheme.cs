namespace Indtec.ExcelMapper.Styling;

public interface IExcelTheme<T>
{
    void Configure(ExcelExportOptions<T> options);
}

public abstract class ExcelTheme<T> : IExcelTheme<T>
{
    public abstract void Configure(ExcelExportOptions<T> options);
}
