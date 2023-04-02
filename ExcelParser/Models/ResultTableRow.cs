namespace ExcelParser.Models
{
    public record ResultTableRow : ProductItem
    {
    }

    public record ProductItem
    {
        /// <summary>
        /// Артикул 2
        /// </summary>
        public string VendorCode2 { get; init; } = string.Empty;

        /// <summary>
        /// Наименование
        /// </summary>
        public string Name { get; init; } = string.Empty;

        /// <summary>
        /// Цвет
        /// </summary>
        public string Color = string.Empty;

        /// <summary>
        /// Количество
        /// </summary>
        public int Count { get; init; }

        /// <summary>
        /// Штрих-код
        /// </summary>
        public string Barcode { get; init; } = string.Empty;

        /// <summary>
        /// Площадь
        /// </summary>
        public int Area { get; set; }
    }

    public class Product
    {
        public string VendorCode1 { get; init; } = string.Empty;
        public string Color { get; init; } = string.Empty;
        public List<ProductItem> Items { get; init; } = new();
    }
}