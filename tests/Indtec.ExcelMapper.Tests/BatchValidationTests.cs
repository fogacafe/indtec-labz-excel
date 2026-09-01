using ClosedXML.Excel;
using Indtec.ExcelMapper.Importing;
using Xunit;

namespace Indtec.ExcelMapper.Tests;

public sealed class BatchValidationTests
{
    [Fact]
    public async Task ImportAsync_ShouldExposeAllParsedRowsAndExistingErrorsToBatchValidator()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "Cost";
            sheet.Cell(1, 4).Value = "Price";
            sheet.Cell(1, 5).Value = "Active";
            sheet.Cell(1, 6).Value = "Status";

            sheet.Cell(2, 1).Value = 1;
            sheet.Cell(2, 2).Value = "Coffee";
            sheet.Cell(2, 3).Value = 10;
            sheet.Cell(2, 4).Value = 12;
            sheet.Cell(2, 5).Value = "Yes";
            sheet.Cell(2, 6).Value = "Active";

            sheet.Cell(3, 1).Value = 2;
            sheet.Cell(3, 2).Value = "Tea";
            sheet.Cell(3, 3).Value = 10;
            sheet.Cell(3, 4).Value = "not-a-price";
            sheet.Cell(3, 5).Value = "No";
            sheet.Cell(3, 6).Value = "Inactive";
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var validator = new RecordingBatchValidator();
        var mapper = new ExcelMapper();

        var result = await mapper.ImportAsync<ProductRow>(stream, options =>
        {
            options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
            options.AddBatchValidator(validator);
        });

        Assert.NotNull(validator.Context);
        Assert.Equal("Products", validator.Context!.SheetName);
        Assert.Equal(2, validator.Context.Rows.Count);
        Assert.False(validator.Context.Rows.Single(x => x.RowNumber == 2).HasErrors);
        Assert.True(validator.Context.Rows.Single(x => x.RowNumber == 3).HasErrors);

        Assert.Empty(result.Items);
        Assert.Equal(2, result.Errors.Count);
        Assert.Contains(result.Errors, x => x.Row == 2 && x.Message == "Rejected by external batch validation.");
        Assert.Contains(result.Errors, x => x.Row == 3 && x.Column == "Price");
    }

    [Fact]
    public void Import_ShouldRejectBatchValidatorOnSynchronousApi()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "Cost";
            sheet.Cell(1, 4).Value = "Price";
            sheet.Cell(1, 5).Value = "Active";
            sheet.Cell(1, 6).Value = "Status";
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper();

        var error = Assert.Throws<ExcelMappingException>(() => mapper.Import<ProductRow>(stream, options =>
            options.AddBatchValidator(new RecordingBatchValidator())));

        Assert.Contains("ImportAsync", error.Message);
    }
}

internal sealed class RecordingBatchValidator : IExcelBatchValidator<ProductRow>
{
    public ExcelBatchValidationContext<ProductRow>? Context { get; private set; }

    public ValueTask<IReadOnlyList<ExcelImportError>> ValidateAsync(
        ExcelBatchValidationContext<ProductRow> context,
        CancellationToken cancellationToken = default)
    {
        Context = context;

        IReadOnlyList<ExcelImportError> errors = new[]
        {
            new ExcelImportError(2, nameof(ProductRow.Id), "Rejected by external batch validation.")
        };

        return new ValueTask<IReadOnlyList<ExcelImportError>>(errors);
    }
}
