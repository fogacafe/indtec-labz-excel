namespace Indtec.ExcelMapper.Styling;

/// <summary>Defines a reusable export or template style configuration for a mapped model.</summary>
public interface IExcelTheme<T>
{
    /// <summary>Applies the theme to the supplied export options.</summary>
    void Configure(ExcelExportOptions<T> options);
}

/// <summary>Base class for reusable strongly typed Excel themes.</summary>
public abstract class ExcelTheme<T> : IExcelTheme<T>
{
    /// <summary>Applies the theme to the supplied export options.</summary>
    public abstract void Configure(ExcelExportOptions<T> options);
}
