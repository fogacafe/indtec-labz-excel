# Indtec.ExcelMapper

[![NuGet](https://img.shields.io/nuget/v/Indtec.ExcelMapper.svg)](https://www.nuget.org/packages/Indtec.ExcelMapper)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Indtec.ExcelMapper.svg)](https://www.nuget.org/packages/Indtec.ExcelMapper)
[![CI](https://github.com/fogacafe/indtec-labz-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/fogacafe/indtec-labz-excel/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A source-generated Excel mapper for .NET: map strongly typed models to `.xlsx` files with attributes, import/export, custom converters, reusable themes and row-aware conditional styling.

```bash
dotnet add package Indtec.ExcelMapper
```

## Why

Excel integrations tend to accumulate repetitive header lookup, conversion, styling and property-assignment code. `Indtec.ExcelMapper` keeps that mapping declarative while generating strongly typed accessors at compile time.

The runtime uses ClosedXML internally, but public models, converters and styling APIs do not expose ClosedXML types. The core mapping path does not scan properties with reflection or call `PropertyInfo.GetValue` / `PropertyInfo.SetValue` for every cell.

## Quick start

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
}
```

Mapped classes are `partial` because the Roslyn incremental source generator emits the strongly typed mapping code during compilation.

### Export

```csharp
var mapper = new ExcelMapper();
mapper.Export(products, "products.xlsx");
```

### Import

```csharp
using var stream = File.OpenRead("products.xlsx");
var products = mapper.Import<Product>(stream);
```

Columns marked `Required = true` are validated before row mapping begins.

## Row-aware conditional styling

Rules receive the complete typed row, so the style of one cell can depend on another property without column indexes or string-based lookups.

```csharp
mapper.Export(products, "products.xlsx", options =>
{
    options.Header
        .Bold()
        .Background("#1F2937")
        .FontColor("#FFFFFF");

    options.Column(x => x.Price)
        .NumberFormat("#,##0.00")
        .Width(18)
        .When(row => row.Price < row.Cost)
        .Background("#FFCCCC")
        .Bold();

    options.Row()
        .When(row => row.Price <= 0)
        .FontColor("#999999");
});
```

Headers are frozen and auto-filtered by default. Styling also supports fonts, fills, borders, alignment, wrapping, number formats and widths.

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

```csharp
mapper.Export(products, stream, options =>
    options.UseTheme(new ProductTheme()));
```

## Custom converters

Converters use the library-owned `ExcelValue` abstraction rather than ClosedXML types.

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

```csharp
[ExcelColumn("Active", Converter = typeof(YesNoBoolConverter))]
public bool Active { get; set; }
```

## How it works

```text
Attributes
   ↓
Roslyn incremental source generator
   ↓
Strongly typed generated mapping
   ↓
ExcelMapper runtime
   ↓
ClosedXML internal adapter
```

This keeps reflection out of the per-cell mapping path while preserving a small attribute-based API.

## Compatibility

The package targets:

- `netstandard2.0` for broad .NET compatibility.
- `net8.0` as a modern optimized target.
- The source generator itself targets `netstandard2.0`.

## Features

- Source-generated sheet and column mapping.
- Strongly typed import and export.
- Required-column validation.
- Custom value converters.
- Reusable typed themes.
- Header, column and row styling.
- Cross-column `When(row => ...)` conditional rules.
- Number formats, widths, fonts, fills, borders, alignment and wrap text.
- Freeze header and auto-filter.
- CI-tested NuGet releases using GitHub OIDC Trusted Publishing.

## Project history

`Indtec.ExcelMapper` is the redesigned successor to [`Codebrew.ExcelAnnotations`](https://github.com/fogacafe/Codebrew.ExcelAnnotations). The original package validated the annotation-based Excel mapping idea; this implementation rebuilds it around source generation, stronger typing and a cleaner public API.

## Contributing

Issues and pull requests are welcome. For behavior changes, adding or updating tests alongside the change is encouraged.

## License

MIT
