using Indtec.ExcelMapper;

var mapper = new ExcelMapper();
var products = new[]
{
    new Product { Id = 1, Name = "Coffee", Price = 12.50m },
    new Product { Id = 2, Name = "Tea", Price = 8.75m }
};

mapper.Export(products, "products.xlsx");

using var stream = File.OpenRead("products.xlsx");
var imported = mapper.Import<Product>(stream);

foreach (var product in imported)
    Console.WriteLine($"{product.Id} - {product.Name} - {product.Price:C}");

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
