# Indtec.ExcelMapper

Source-generated Excel mapper for .NET. Map strongly typed models to `.xlsx` files with attributes, without runtime property discovery or `PropertyInfo.SetValue` in the mapping path.

> Status: early development (`0.1.0-alpha`). The public API may still change before `1.0.0`.

## Why

Most Excel import/export code eventually turns into repetitive column lookup, conversion and property assignment. `Indtec.ExcelMapper` keeps the model declaration small while generating the mapping code at compile time.

The runtime uses ClosedXML as an internal workbook engine. Models and generated mappings do not expose ClosedXML types.

## Targets

- `netstandard2.0` for broad compatibility, including modern SDK-style .NET Framework projects.
- `net8.0` as a modern target.
- Source generator targeting `netstandard2.0`.

## Example

```csharp
using Indtec.ExcelMapper;

[ExcelSheet("Products")]
public partial class Product
{
    [ExcelColumn("Id", Order = 1)]
    public int Id { get; set; }

    [ExcelColumn("Product Name", Order = 2)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Price", Order = 3)]
    public decimal Price { get; set; }
}
```

The mapped class must be `partial`. The generator adds a strongly typed map during compilation.

### Export

```csharp
var mapper = new ExcelMapper();

mapper.Export(products, "products.xlsx");
```

Or write to a stream:

```csharp
using var stream = new MemoryStream();
mapper.Export(products, stream);
```

### Import

```csharp
using var stream = File.OpenRead("products.xlsx");
var products = mapper.Import<Product>(stream);
```

## How mapping works

At compile time the source generator reads `ExcelSheet` and `ExcelColumn` metadata and emits strongly typed getters and setters for each mapped property. At runtime `ExcelMapper` consumes that generated map and uses ClosedXML only for workbook I/O.

That means the core mapping path does not scan properties with reflection and does not call `PropertyInfo.GetValue` / `PropertyInfo.SetValue` for every cell.

## Current scope

The foundation currently covers:

- Attribute-based sheet and column mapping.
- Column ordering.
- Strongly typed generated getters/setters.
- Import from streams.
- Export to streams and paths.
- Primitive, nullable, enum, date/time and GUID conversion groundwork.
- Multi-targeting (`netstandard2.0;net8.0`).
- Tests, sample project and CI packing a NuGet artifact.

## Roadmap

Next iterations will add the styling layer as metadata rather than ClosedXML behavior leaking into the public API, including themes, headers, number formats, widths, borders, alignment, freeze panes and conditional styling (`When(...)`).

Further roadmap items include custom converters, diagnostics, required-column policies, richer import errors, source-generator tests, AOT/trimming validation and NuGet release automation.

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
