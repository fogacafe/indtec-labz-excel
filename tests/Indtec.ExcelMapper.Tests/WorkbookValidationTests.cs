using ClosedXML.Excel;
using Indtec.ExcelMapper.Importing;
using Indtec.ExcelMapper.Workbooks;
using Xunit;

namespace Indtec.ExcelMapper.Tests;

public sealed class WorkbookValidationTests
{
    [Fact]
    public async Task ImportWorkbookAsync_ShouldAggregateCrossSheetErrorsOnTargetRows()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var orderSheet = workbook.AddWorksheet("Orders");
            orderSheet.Cell(1, 1).Value = "Id";
            orderSheet.Cell(1, 2).Value = "CustomerId";
            orderSheet.Cell(2, 1).Value = 1;
            orderSheet.Cell(2, 2).Value = 999;

            var customerSheet = workbook.AddWorksheet("CustomerRefs");
            customerSheet.Cell(1, 1).Value = "Id";
            customerSheet.Cell(2, 1).Value = 42;

            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper();

        var result = await mapper.ImportWorkbookAsync(stream, workbook =>
        {
            workbook.Sheet<OrderRow>(options =>
                options.ErrorBehavior = ExcelImportErrorBehavior.Collect);

            workbook.Sheet<CustomerRefRow>(options =>
                options.ErrorBehavior = ExcelImportErrorBehavior.Collect);

            workbook.AddValidator(new OrderCustomerValidator());
        });

        var orders = result.Sheet<OrderRow>();

        Assert.Empty(orders.Items);
        Assert.Single(orders.Errors);
        Assert.Single(orders.Rows);
        Assert.True(orders.Rows[0].HasErrors);
        Assert.Equal(2, orders.Errors[0].Row);
        Assert.Equal(nameof(OrderRow.CustomerId), orders.Errors[0].Column);
        Assert.Equal("Customer was not found in CustomerRefs.", orders.Errors[0].Message);
    }

    [Fact]
    public async Task ImportWorkbookAsync_WhenTargetSheetUsesThrow_ShouldThrowWorkbookValidationError()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var orderSheet = workbook.AddWorksheet("Orders");
            orderSheet.Cell(1, 1).Value = "Id";
            orderSheet.Cell(1, 2).Value = "CustomerId";
            orderSheet.Cell(2, 1).Value = 1;
            orderSheet.Cell(2, 2).Value = 999;

            var customerSheet = workbook.AddWorksheet("CustomerRefs");
            customerSheet.Cell(1, 1).Value = "Id";
            customerSheet.Cell(2, 1).Value = 42;

            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper();

        var error = await Assert.ThrowsAsync<ExcelMappingException>(() =>
            mapper.ImportWorkbookAsync(stream, workbook =>
            {
                workbook.Sheet<OrderRow>();
                workbook.Sheet<CustomerRefRow>();
                workbook.AddValidator(new OrderCustomerValidator());
            }));

        Assert.Contains(nameof(OrderRow.CustomerId), error.Message);
        Assert.Contains("Customer was not found", error.Message);
    }
}

internal sealed class OrderCustomerValidator : IExcelWorkbookValidator
{
    public Task<IReadOnlyList<ExcelWorkbookValidationError>> ValidateAsync(
        ExcelWorkbookValidationContext context,
        CancellationToken cancellationToken = default)
    {
        var customerIds = new HashSet<int>(
            context.Sheet<CustomerRefRow>().Items.Select(customer => customer.Id));

        IReadOnlyList<ExcelWorkbookValidationError> errors = context
            .Sheet<OrderRow>()
            .Rows
            .Where(row => !row.HasErrors && !customerIds.Contains(row.Value.CustomerId))
            .Select(row => ExcelWorkbookValidationError.For<OrderRow>(
                row.RowNumber,
                nameof(OrderRow.CustomerId),
                "Customer was not found in CustomerRefs."))
            .ToArray();

        return Task.FromResult(errors);
    }
}

[ExcelSheet("Orders")]
public partial class OrderRow
{
    [ExcelColumn("Id", Order = 1, Required = true)]
    public int Id { get; set; }

    [ExcelColumn("CustomerId", Order = 2, Required = true)]
    public int CustomerId { get; set; }
}

[ExcelSheet("CustomerRefs")]
public partial class CustomerRefRow
{
    [ExcelColumn("Id", Order = 1, Required = true)]
    public int Id { get; set; }
}
