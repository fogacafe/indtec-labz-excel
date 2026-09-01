# Indtec.ExcelMapper

[![NuGet](https://img.shields.io/nuget/v/Indtec.ExcelMapper.svg)](https://www.nuget.org/packages/Indtec.ExcelMapper)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Indtec.ExcelMapper.svg)](https://www.nuget.org/packages/Indtec.ExcelMapper)
[![CI](https://github.com/fogacafe/indtec-labz-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/fogacafe/indtec-labz-excel/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A source-generated Excel mapper for .NET: strongly typed `.xlsx` import/export, validation, multi-sheet workbooks, templates, localized messages, styling and memory-bounded streaming imports for large worksheets.

```bash
dotnet add package Indtec.ExcelMapper
```

## Why

Excel integrations tend to accumulate repetitive header lookup, conversion, validation, styling and property-assignment code. `Indtec.ExcelMapper` keeps that mapping declarative while generating strongly typed accessors at compile time.

The regular import/export path uses ClosedXML internally. Public models, converters, validators and styling APIs do not expose ClosedXML types, and the core mapping path does not scan properties with reflection or call `PropertyInfo.GetValue` / `PropertyInfo.SetValue` for every cell.

For large worksheet imports, a separate forward-only Open XML path reads rows incrementally and delivers bounded chunks without materializing the full worksheet in ClosedXML.

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

    [ExcelColumn("Status", Order = 5)]
    public ProductStatus Status { get; set; }
}

public enum ProductStatus
{
    Active,
    Inactive,
    Discontinued
}
```

Mapped classes are `partial` because the Roslyn incremental source generator emits the strongly typed mapping code during compilation.

### Export

```csharp
var mapper = new ExcelMapper();
mapper.Export(products, "products.xlsx");
```

All export APIs also accept a `Stream`, so `MemoryStream`, HTTP responses and cloud-storage workflows do not require temporary files:

```csharp
using var stream = new MemoryStream();
mapper.Export(products, stream);
stream.Position = 0;
```

### Import

```csharp
using var stream = File.OpenRead("products.xlsx");
var products = mapper.Import<Product>(stream);
```

Columns marked `Required = true` are validated before row mapping begins.

## Import validation and error collection

Use typed row rules when business validation needs values from more than one cell.

```csharp
using Indtec.ExcelMapper.Importing;

var result = mapper.Import<Product>(stream, options =>
{
    options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
    options.Validate(
        row => row.Price >= row.Cost,
        "Price cannot be lower than cost.");
});
```

`Collect` mode preserves valid rows and returns structured errors:

```csharp
foreach (var error in result.Errors)
    Console.WriteLine($"Row {error.Row} | {error.Column} | {error.Message}");

var validProducts = result.Items;
```

Async imports also expose every parsed row together with its current validation state:

```csharp
foreach (var row in result.Rows)
{
    Console.WriteLine($"Excel row {row.RowNumber}: {row.HasErrors}");

    foreach (var error in row.Errors)
        Console.WriteLine(error.Message);
}
```

Use `ExcelImportErrorBehavior.Throw` when the first invalid row should fail the import immediately.

## Async batch validation

Use `IExcelBatchValidator<T>` when a validation needs the complete sheet, a batched service call, duplicate detection or another asynchronous operation.

```csharp
public sealed class CustomerBatchValidator : IExcelBatchValidator<Trade>
{
    private readonly ICustomerService _service;

    public CustomerBatchValidator(ICustomerService service)
        => _service = service;

    public async Task<IReadOnlyList<ExcelImportError>> ValidateAsync(
        ExcelBatchValidationContext<Trade> context,
        CancellationToken cancellationToken = default)
    {
        var ids = context.Rows
            .Where(row => !row.HasErrors)
            .Select(row => row.Value.CustomerId)
            .Distinct()
            .ToArray();

        var validIds = await _service.GetValidIdsAsync(ids, cancellationToken);

        return context.Rows
            .Where(row => !row.HasErrors && !validIds.Contains(row.Value.CustomerId))
            .Select(row => new ExcelImportError(
                row.RowNumber,
                nameof(Trade.CustomerId),
                "Customer was not found."))
            .ToArray();
    }
}
```

Register the validator explicitly; no dependency-injection integration is required by the library:

```csharp
var result = await mapper.ImportAsync<Trade>(stream, options =>
{
    options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
    options.AddBatchValidator(new CustomerBatchValidator(customerService));
});
```

Batch validators receive all parsed rows, including rows that already contain mapping or local-validation errors. Validators return new errors rather than mutating row state.

## Large worksheets: streaming chunk import

When a worksheet is too large to comfortably materialize as a full ClosedXML workbook, use `ImportChunksAsync<T>`. The streaming path uses `OpenXmlReader` and only keeps the current chunk of worksheet rows in memory.

```csharp
await mapper.ImportChunksAsync<Trade>(
    stream,
    async (chunk, cancellationToken) =>
    {
        await repository.SaveAsync(chunk.Items, cancellationToken);

        foreach (var error in chunk.Errors)
            logger.LogWarning("{Error}", error);
    },
    chunkSize: 1000,
    configure: options =>
    {
        options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
        options.Validate(row => row.Price > 0, "Price must be positive.");
    },
    cancellationToken);
```

Each `ExcelImportChunk<T>` exposes:

```csharp
chunk.Index
chunk.StartRow
chunk.EndRow
chunk.Items
chunk.Rows
chunk.Errors
chunk.IsValid
```

After the callback returns, the mapper can release that chunk and continue reading the worksheet. This makes flows such as database persistence, queue publishing or external-service enrichment possible without accumulating all imported rows.

The streaming API requires a readable, seekable `.xlsx` stream. It keeps the workbook shared-string table in memory because `.xlsx` files may reference it from any row, but it does not materialize the full worksheet DOM or retain previous chunks.

`IExcelBatchValidator<T>` intentionally remains a full-sheet validator and is therefore rejected by `ImportChunksAsync<T>`. For chunk-scoped external validation, perform the batch operation inside the chunk callback. This avoids silently changing validator semantics.

## Multi-sheet workbook import

Import several mapped models while opening the workbook only once. Each sheet keeps its own import configuration and validators.

```csharp
var result = await mapper.ImportWorkbookAsync(stream, workbook =>
{
    workbook.Sheet<Trade>(options =>
    {
        options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
        options.Validate(row => row.Price > 0, "Price must be positive.");
        options.AddBatchValidator(new TradeBatchValidator(service));
    });

    workbook.Sheet<Customer>(options =>
        options.ErrorBehavior = ExcelImportErrorBehavior.Collect);
});

var trades = result.Sheet<Trade>();
var customers = result.Sheet<Customer>();
```

The sheet name still comes from `[ExcelSheet]`; the workbook layer only orchestrates independent typed mappings.

## Workbook-level validation

Use `IExcelWorkbookValidator` when a rule depends on more than one sheet, such as validating references from `Trades.CustomerId` against `Customers.Id`.

```csharp
public sealed class TradeCustomerValidator : IExcelWorkbookValidator
{
    public Task<IReadOnlyList<ExcelWorkbookValidationError>> ValidateAsync(
        ExcelWorkbookValidationContext context,
        CancellationToken cancellationToken = default)
    {
        var customerIds = context
            .Sheet<Customer>()
            .Items
            .Select(customer => customer.Id)
            .ToHashSet();

        IReadOnlyList<ExcelWorkbookValidationError> errors = context
            .Sheet<Trade>()
            .Rows
            .Where(row => !row.HasErrors && !customerIds.Contains(row.Value.CustomerId))
            .Select(row => ExcelWorkbookValidationError.For<Trade>(
                row.RowNumber,
                nameof(Trade.CustomerId),
                "Customer was not found in Customers."))
            .ToArray();

        return Task.FromResult(errors);
    }
}
```

Register it at workbook level:

```csharp
var result = await mapper.ImportWorkbookAsync(stream, workbook =>
{
    workbook.Sheet<Trade>(options =>
        options.ErrorBehavior = ExcelImportErrorBehavior.Collect);

    workbook.Sheet<Customer>(options =>
        options.ErrorBehavior = ExcelImportErrorBehavior.Collect);

    workbook.AddValidator(new TradeCustomerValidator());
});
```

Workbook validators see typed sheet results and existing row errors, return immutable validation errors, and those errors are routed back to the target sheet. The target sheet's `Collect` / `Throw` behavior is respected.

The validation pipeline is intentionally layered:

```text
cell mapping
    ↓
local row validation
    ↓
async batch validation
    ↓
workbook / cross-sheet validation
```

## Localization

Processing messages are available in English (default) and Brazilian Portuguese.

```csharp
using Indtec.ExcelMapper.Localization;

var mapper = new ExcelMapper(options =>
    options.Language = ExcelLanguage.PortugueseBrazil);
```

This covers mapper-generated messages such as missing worksheets, required columns, cell conversions, workbook configuration and generated Excel validation prompts. Validation messages supplied by your own application remain exactly as you provide them.

For custom wording or another language, override only the messages you need:

```csharp
public sealed class MyExcelMessages : ExcelMessageProvider
{
    public override string WorksheetNotFound(string sheetName)
        => $"Não achei a aba {sheetName}.";
}

var mapper = new ExcelMapper(options =>
    options.Messages = new MyExcelMessages());
```

You can also implement `IExcelMessageProvider` directly when you want complete control.

## Excel display formats

Excel stores numbers and dates as typed values, so formatting stays a presentation concern. Use custom masks with `NumberFormat(...)` / `DateFormat(...)` or one of the built-in common masks:

```csharp
using Indtec.ExcelMapper.Styling;

mapper.Export(trades, stream, options =>
{
    options.Column(x => x.Price)
        .NumberFormat(ExcelFormats.CurrencyBrazil);

    options.Column(x => x.TradeDate)
        .DateFormat(ExcelFormats.DateBrazil);
});
```

Available helpers include Brazilian and ISO date/date-time masks, integer/decimal masks, percentage and BRL/USD currency masks.

## Excel templates

Generate an empty workbook directly from the mapped model:

```csharp
mapper.CreateTemplate<Product>("products-template.xlsx", options =>
{
    options.UseTheme(new ProductTheme());
    options.TemplateRows = 1000;

    options.Column(x => x.Name)
        .AllowedValues("Coffee", "Tea", "Milk");
});
```

Templates reuse headers, themes, widths, number formats, freeze-header and auto-filter settings. Enum columns automatically receive a dropdown containing their enum values.

`AllowedValues(...)` can be used for custom dropdowns such as status codes, currencies or business-domain options.

### Multi-sheet templates

Generate a complete import workbook with independent configuration per sheet:

```csharp
mapper.CreateWorkbookTemplate("import-template.xlsx", workbook =>
{
    workbook.Sheet<Trade>(options =>
    {
        options.UseTheme(new TradeTheme());
        options.Column(x => x.Currency)
            .AllowedValues("BRL", "USD", "EUR");
    });

    workbook.Sheet<Customer>();
    workbook.Sheet<Product>();
});
```

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
   ├── ClosedXML adapter (regular import/export/templates)
   └── Open XML forward reader (streaming chunks)
```

This keeps reflection out of the per-cell mapping path while preserving a small attribute-based API.

## Compatibility

The package targets:

- `netstandard2.0` for broad .NET compatibility.
- `net8.0` as a modern optimized target.
- The source generator itself targets `netstandard2.0`.

## Features

- Source-generated sheet and column mapping.
- Strongly typed import and export to paths or streams.
- Structured import results with collect/throw behavior.
- Typed cross-column row validation.
- Async batch validators with cancellation support.
- Memory-bounded streaming chunk imports for large worksheets.
- Multi-sheet workbook import with independent sheet configuration.
- Workbook-level cross-sheet validation.
- English and Brazilian Portuguese processing messages with custom overrides.
- Common Excel display-format masks.
- Excel template and multi-sheet workbook-template generation.
- Custom dropdown values and automatic enum dropdowns.
- Required-column validation.
- Custom value converters.
- Reusable typed themes.
- Header, column and row styling.
- Cross-column `When(row => ...)` conditional rules.
- Number formats, widths, fonts, fills, borders, alignment and wrap text.
- Freeze header and auto-filter.
- `netstandard2.0` + `net8.0` targets.
- CI-tested NuGet releases using GitHub OIDC Trusted Publishing.

## Project history

`Indtec.ExcelMapper` is the redesigned successor to [`Codebrew.ExcelAnnotations`](https://github.com/fogacafe/Codebrew.ExcelAnnotations). The original package validated the annotation-based Excel mapping idea; this implementation rebuilds it around source generation, stronger typing and a cleaner public API.

## Contributing

Issues and pull requests are welcome. For behavior changes, adding or updating tests alongside the change is encouraged.

## License

MIT
