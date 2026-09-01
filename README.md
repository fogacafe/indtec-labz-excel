# Indtec.ExcelMapper

Source-generated Excel mapper for .NET. Map strongly typed models to `.xlsx` files with attributes, without runtime property discovery or `PropertyInfo.SetValue` in the mapping path.

## Why

Most Excel import/export code eventually turns into repetitive column lookup, conversion, styling and property assignment. `Indtec.ExcelMapper` keeps the model declaration small while generating the mapping code at compile time.

ClosedXML is an internal workbook engine. Models, generated mappings, converters and styling rules do not expose ClosedXML types.

## Targets

- `netstandard2.0` for broad compatibility.
- `net8.0` as a modern target.
- Source generator targeting `netstandard2.0`.

## Installation

```bash
dotnet add package Indtec.ExcelMapper
```

## Mapping

```csharp
using Indtec.ExcelMapper;

[ExcelSheet("Products")]
public partial class Product
{
    [ExcelColumn("Id", Order = 1, Required = true)]
    public int Id { get; set; }

    [ExcelColumn("Product Name", Order = 2, Required = true)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Cost", Order = 3)]
    public decimal Cost { get; set; }

    [ExcelColumn("Price", Order = 4)]
    public decimal Price { get; set; }

    [ExcelColumn("Active", Order = 5, Converter = typeof(YesNoBoolConverter))]
    public bool Active { get; set; }
}
```

Mapped classes must be `partial`. The source generator emits strongly typed getters, setters and mapping metadata during compilation.

## Import and export

```csharp
var mapper = new ExcelMapper();
mapper.Export(products, "products.xlsx");

using var stream = File.OpenRead("products.xlsx");
var imported = mapper.Import<Product>(stream);
```

Columns marked `Required = true` are validated before rows are mapped.

## Typed row-aware styling

Conditional styling receives the complete typed row, allowing one cell to react to values from other properties.

```csharp
mapper.Export(products, "products.xlsx", options =>
{
    options.Column(x => x.Price)
        .NumberFormat("#,##0.00")
        .Width(18)
        .When(row => row.Price < row.Cost)
        .Background("#FFCCCC")
        .Bold();

    options.Row()
        .When(row => !row.Active)
        .FontColor("#999999");
});
```

Headers are frozen and auto-filtered by default.

## Reusable themes

```csharp
public sealed class ProductTheme : ExcelTheme<Product>
{
    public override void Configure(ExcelExportOptions<Product> options)
    {
        options.Header
            .Bold()
            .Background("#1F2937")
            .FontColor("#FFFFFF");

        options.Column(x => x.Price)
            .NumberFormat("#,##0.00")
            .Width(18);
    }
}
```

Use it anywhere:

```csharp
mapper.Export(products, stream, options =>
    options.UseTheme(new ProductTheme()));
```

## Custom converters

Converters work with `ExcelValue`, not ClosedXML types.

```csharp
using Indtec.ExcelMapper.Conversion;

public sealed class YesNoBoolConverter : IExcelValueConverter
{
    public object? Read(ExcelValue value, Type destinationType)
        => string.Equals(value.AsString(), "Yes", StringComparison.OrdinalIgnoreCase);

    public ExcelValue Write(object? value)
        => new(value is true ? "Yes" : "No");
}
```

Attach it at compile time:

```csharp
[ExcelColumn("Active", Converter = typeof(YesNoBoolConverter))]
public bool Active { get; set; }
```

## Architecture

```text
Attributes
   ↓
Roslyn incremental source generator
   ↓
Strongly typed generated map
   ↓
ExcelMapper runtime
   ↓
ClosedXML internal adapter
```

The core mapping path does not scan properties with reflection and does not call `PropertyInfo.GetValue` / `PropertyInfo.SetValue` per cell.

## 1.0 scope

- Source-generated sheet and column mapping.
- Required-column validation.
- Custom value converters.
- Strongly typed import/export.
- Reusable typed themes.
- Header, column and row styling.
- Cross-column `When(row => ...)` rules.
- Number formats, widths, fonts, fills, borders, alignment and wrap text.
- Freeze header and auto-filter.
- `netstandard2.0` and `net8.0` targets.
- CI build, test and NuGet packing.

## Repository structure

```text
src/
  Indtec.ExcelMapper/
  Indtec.ExcelMapper.Generators/
tests/
  Indtec.ExcelMapper.Tests/
samples/
  Indtec.ExcelMapper.Sample/
```

## License

MIT
