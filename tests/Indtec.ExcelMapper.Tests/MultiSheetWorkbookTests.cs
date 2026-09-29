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

        Assert.Equal(2, workbook.Worksheets.Count(x => x.Visibility == XLWorksheetVisibility.Visible));
        Assert.Equal("Id", workbook.Worksheet("Products").Cell(1, 1).GetString());
        Assert.True(workbook.Worksheet("Products").DataValidations.Any());
        Assert.Equal("Name", workbook.Worksheet("Customers").Cell(1, 2).GetString());
        Assert.True(workbook.Worksheet("Customers").Cell(1, 1).Style.Font.Bold);
    }

    [Fact]
    public void CreateWorkbookTemplate_ShouldPopulateInitialRowsAndAutoFitHeaders()
    {
        using var stream = new MemoryStream();
        var mapper = new ExcelMapper();
        var products = new[]
        {
            new ProductRow
            {
                Id = 1,
                Name = "Coffee",
                Cost = 10m,
                Price = 12.5m,
                Active = true,
                Status = ProductStatus.Active
            }
        };

        mapper.CreateWorkbookTemplate(stream, workbook =>
        {
            workbook.Sheet(products, options =>
            {
                options.AutoFitHeaders = true;
                options.Column(x => x.Price).Width(30);
            });

            workbook.Sheet<CustomerRow>();
        });

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        var sheet = workbook.Worksheet("Products");

        Assert.Equal("Coffee", sheet.Cell(2, 2).GetString());
        Assert.Equal("Yes", sheet.Cell(2, 5).GetString());
        Assert.True(sheet.Column(2).Width > 0);
        Assert.Equal(30d, sheet.Column(4).Width);
    }

    [Fact]
    public void CreateWorkbookTemplate_ShouldUseHiddenRangeForLongDataValidationLists()
    {
        using var stream = new MemoryStream();
        var mapper = new ExcelMapper();
        var values = Enumerable.Range(1, 40)
            .Select(i => $"Allowed value {i:00} with enough text to exceed the inline Excel limit")
            .ToArray();

        Assert.True(string.Join(",", values).Length > 255);

        mapper.CreateWorkbookTemplate(stream, workbook =>
        {
            workbook.Sheet<ProductRow>(options =>
            {
                options.TemplateRows = 25;
                options.Column(x => x.Name).AllowedValues(values);
            });

            workbook.Sheet<CustomerRow>();
        });

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        var sheet = workbook.Worksheet("Products");
        var validationSheet = workbook.Worksheet("__IndtecValidation");

        Assert.Equal(XLWorksheetVisibility.VeryHidden, validationSheet.Visibility);
        Assert.True(sheet.DataValidations.Any());

        var validation = sheet.DataValidations.First();
        Assert.Equal(XLAllowedValues.List, validation.AllowedValues);
        Assert.StartsWith("=__IndtecValidation_", validation.Value);
        Assert.Contains(values[0], validationSheet.CellsUsed().Select(x => x.GetString()));
        Assert.Contains(values[^1], validationSheet.CellsUsed().Select(x => x.GetString()));
    }
    [Fact]
    public void ExportWorkbook_ShouldExportMultipleTypedSheets()
    {
        using var stream = new MemoryStream();
        var mapper = new ExcelMapper();

        mapper.ExportWorkbook(stream, workbook =>
        {
            workbook.Sheet(
                new[] { new ProductRow { Id = 1, Name = "Coffee", Price = 12.5m, Active = true } },
                options => options.AutoFitHeaders = true);

            workbook.Sheet(
                new[] { new CustomerRow { Id = 7, Name = "Alice" } });
        });

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);

        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.Equal("Coffee", workbook.Worksheet("Products").Cell(2, 2).GetString());
        Assert.Equal("Yes", workbook.Worksheet("Products").Cell(2, 5).GetString());
        Assert.Equal("Alice", workbook.Worksheet("Customers").Cell(2, 2).GetString());
    }

    [Fact]
    public async Task ImportWorkbookAsync_ShouldAllowOptionalMissingSheet()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(2, 1).Value = 1;
            sheet.Cell(2, 2).Value = "Coffee";
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var result = await new ExcelMapper().ImportWorkbookAsync(stream, workbook =>
        {
            workbook.Sheet<ProductRow>();
            workbook.Sheet<CustomerRow>(options => options.OptionalSheet = true);
        });

        Assert.Single(result.Sheet<ProductRow>().Items);
        Assert.Empty(result.Sheet<CustomerRow>().Items);
    }

    [Fact]
    public void ExportWorkbook_ShouldAllowSameModelWithDifferentSheetNames()
    {
        using var stream = new MemoryStream();
        var mapper = new ExcelMapper();

        mapper.ExportWorkbook(stream, workbook =>
        {
            workbook.Sheet(new[] { new CustomerRow { Id = 1, Name = "Alice" } },
                options => options.SheetName = "Active Customers");
            workbook.Sheet(new[] { new CustomerRow { Id = 2, Name = "Bob" } },
                options => options.SheetName = "Inactive Customers");
        });

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        Assert.Equal("Alice", workbook.Worksheet("Active Customers").Cell(2, 2).GetString());
        Assert.Equal("Bob", workbook.Worksheet("Inactive Customers").Cell(2, 2).GetString());
    }

    [Fact]
    public async Task ImportWorkbookAsync_MissingRequiredSheet_ShouldThrow()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("Products");
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        await Assert.ThrowsAsync<ExcelMappingException>(() =>
            new ExcelMapper().ImportWorkbookAsync(stream, workbook =>
                workbook.Sheet<CustomerRow>()));
    }

    [Fact]
    public void ExportWorkbook_DuplicateResolvedSheetNames_ShouldThrow()
    {
        using var stream = new MemoryStream();
        var mapper = new ExcelMapper();

        Assert.Throws<ExcelMappingException>(() =>
            mapper.ExportWorkbook(stream, workbook =>
            {
                workbook.Sheet(new[] { new CustomerRow { Id = 1, Name = "Alice" } },
                    options => options.SheetName = "People");
                workbook.Sheet(new[] { new ProductRow { Id = 2, Name = "Coffee" } },
                    options => options.SheetName = "people");
            }));
    }

    [Fact]
    public void CreateWorkbookTemplate_ShouldHonorSheetNameOverride()
    {
        using var stream = new MemoryStream();

        new ExcelMapper().CreateWorkbookTemplate(stream, workbook =>
            workbook.Sheet<CustomerRow>(options => options.SheetName = "People"));

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        Assert.True(workbook.TryGetWorksheet("People", out _));
        Assert.False(workbook.TryGetWorksheet("Customers", out _));
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
