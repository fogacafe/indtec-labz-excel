using ClosedXML.Excel;
using Indtec.ExcelMapper.Importing;
using Xunit;

namespace Indtec.ExcelMapper.Tests;

public sealed class MultiSheetWorkbookTests
{
    [Fact]
    public async Task ImportWorkbookAsync_ShouldMapRegisteredSheetsWithIndependentConfiguration()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var products = workbook.AddWorksheet("Products");
            products.Cell(1, 1).Value = "Id";
            products.Cell(1, 2).Value = "Name";
            products.Cell(1, 3).Value = "Cost";
            products.Cell(1, 4).Value = "Price";
            products.Cell(1, 5).Value = "Active";
            products.Cell(1, 6).Value = "Status";
            products.Cell(2, 1).Value = 1;
            products.Cell(2, 2).Value = "Coffee";
            products.Cell(2, 3).Value = 10;
            products.Cell(2, 4).Value = 12;
            products.Cell(2, 5).Value = "Yes";
            products.Cell(2, 6).Value = "Active";

            var customers = workbook.AddWorksheet("Customers");
            customers.Cell(1, 1).Value = "Id";
            customers.Cell(1, 2).Value = "Name";
            customers.Cell(2, 1).Value = 42;
            customers.Cell(2, 2).Value = "Ada";

            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper();

        var result = await mapper.ImportWorkbookAsync(stream, workbook =>
        {
            workbook.Sheet<ProductRow>(options =>
            {
                options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
                options.Validate(x => x.Price >= x.Cost, "Price cannot be lower than cost.");
            });

            workbook.Sheet<CustomerRow>();
        });

        var products = result.Sheet<ProductRow>();
        var customers = result.Sheet<CustomerRow>();

        Assert.Single(products.Items);
        Assert.Equal("Coffee", products.Items[0].Name);
        Assert.Empty(products.Errors);

        Assert.Single(customers.Items);
        Assert.Equal(42, customers.Items[0].Id);
        Assert.Equal("Ada", customers.Items[0].Name);
    }

    [Fact]
    public void CreateWorkbookTemplate_ShouldGenerateEveryRegisteredSheet()
    {
        using var stream = new MemoryStream();
        var mapper = new ExcelMapper();

        mapper.CreateWorkbookTemplate(stream, workbook =>
        {
            workbook.Sheet<ProductRow>(options =>
            {
                options.TemplateRows = 25;
                options.Column(x => x.Name).AllowedValues("Coffee", "Tea");
            });

            workbook.Sheet<CustomerRow>(options =>
                options.Header.Bold());
        });

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);

        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.Equal("Id", workbook.Worksheet("Products").Cell(1, 1).GetString());
        Assert.True(workbook.Worksheet("Products").DataValidations.Any());
        Assert.Equal("Name", workbook.Worksheet("Customers").Cell(1, 2).GetString());
        Assert.True(workbook.Worksheet("Customers").Cell(1, 1).Style.Font.Bold);
    }
}

[ExcelSheet("Customers")]
public partial class CustomerRow
{
    [ExcelColumn("Id", Order = 1, Required = true)]
    public int Id { get; set; }

    [ExcelColumn("Name", Order = 2, Required = true)]
    public string Name { get; set; } = string.Empty;
}
