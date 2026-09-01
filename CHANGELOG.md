# Changelog

All notable changes to `Indtec.ExcelMapper` are documented here.

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
