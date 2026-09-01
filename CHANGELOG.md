# Changelog

All notable changes to `Indtec.ExcelMapper` are documented here.

## 1.3.0

- Built-in English and Brazilian Portuguese processing messages.
- Configurable `ExcelLanguage` through `ExcelMapperOptions`.
- Overridable messages through `IExcelMessageProvider` and `ExcelMessageProvider`.
- Localized worksheet, required-column, conversion, workbook and template-validation messages.
- Common display masks through `ExcelFormats` and a `DateFormat(...)` styling helper.
- Memory-bounded `ImportChunksAsync<T>` for large worksheets using forward-only Open XML reading.
- `ExcelImportChunk<T>` with chunk index, physical row range, valid items, parsed rows and structured errors.
- Source-generated mappings, custom converters, local validation and collect/throw behavior preserved in streaming mode.
- Full-sheet batch validators are explicitly rejected in streaming mode to avoid changing their semantics.

## 1.2.0

- Async batch validation through `IExcelBatchValidator<T>`.
- `ExcelImportResult<T>.Rows` with row number, value, errors and `HasErrors`.
- Multi-sheet workbook import through `ImportWorkbookAsync`.
- Independent import configuration and validators per sheet.
- Multi-sheet template generation through `CreateWorkbookTemplate`.
- Workbook-level cross-sheet validation through `IExcelWorkbookValidator`.
- Immutable workbook validation errors routed back to the target typed sheet.
- Per-sheet `Collect` / `Throw` behavior preserved for workbook validation.

## 1.1.0

- Typed row validation with structured import errors.
- `Collect` and `Throw` import error behaviors.
- Excel template generation.
- Custom allowed-value dropdowns.
- Automatic enum dropdowns.

## 1.0.0

- Source-generated attribute mapping.
- Strongly typed import and export.
- Custom value converters.
- Reusable themes.
- Row-aware conditional styling.
- `netstandard2.0` and `net8.0` targets.
- Automated NuGet publishing through GitHub OIDC Trusted Publishing.
