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
            var productsSheet = workbook.AddWorksheet("Products");
            productsSheet.Cell(1, 1).Value = "Id";
            productsSheet.Cell(1, 2).Value = "Name";
            productsSheet.Cell(1, 3).Value = "Cost";
            productsSheet.Cell(1, 4).Value = "Price";
            productsSheet.Cell(1, 5).Value = "Active";
            productsSheet.Cell(1, 6).Value = "Status";
            productsSheet.Cell(2, 1).Value = 1;
            productsSheet.Cell(2, 2).Value = "Coffee";
            productsSheet.Cell(2, 3).Value = 10;
            productsSheet.Cell(2, 4).Value = 12;
            productsSheet.Cell(2, 5).Value = "Yes";
            productsSheet.Cell(2, 6).Value = "Active";

            var customersSheet = workbook.AddWorksheet("Customers");
            customersSheet.Cell(1, 1).Value = "Id";
            customersSheet.Cell(1, 2).Value = "Name";
            customersSheet.Cell(2, 1).Value = 42;
            customersSheet.Cell(2, 2).Value = "Ada";

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
