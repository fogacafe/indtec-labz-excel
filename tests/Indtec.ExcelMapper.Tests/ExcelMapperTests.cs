namespace Indtec.ExcelMapper.Tests;

public sealed class ExcelMapperTests
{
    [Fact]
    public void ExportThenImport_ShouldPreserveMappedValues()
    {
        var mapper = new ExcelMapper();
        var source = new[]
        {
            new ProductRow { Id = 1, Name = "Coffee", Price = 12.50m, Active = true },
            new ProductRow { Id = 2, Name = "Tea", Price = 8.75m, Active = false }
        };

        using var stream = new MemoryStream();
        mapper.Export(source, stream);
        stream.Position = 0;

        var result = mapper.Import<ProductRow>(stream);

        Assert.Equal(2, result.Count);
        Assert.Equal(1, result[0].Id);
        Assert.Equal("Coffee", result[0].Name);
        Assert.Equal(12.50m, result[0].Price);
        Assert.True(result[0].Active);
        Assert.Equal(2, result[1].Id);
        Assert.Equal("Tea", result[1].Name);
        Assert.Equal(8.75m, result[1].Price);
        Assert.False(result[1].Active);
    }
}

[ExcelSheet("Products")]
public partial class ProductRow
{
    [ExcelColumn("Id", Order = 1)]
    public int Id { get; set; }

    [ExcelColumn("Name", Order = 2)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Price", Order = 3)]
    public decimal Price { get; set; }

    [ExcelColumn("Active", Order = 4)]
    public bool Active { get; set; }
}
