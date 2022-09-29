namespace ExcelParser.Models
{
    internal record ResultTableRow : ProductItem
    {
    }

    internal record ProductItem
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
        /// Количество
        /// </summary>
        public int Count { get; init; } = 0;
        /// <summary>
        /// Штрих-код
        /// </summary>
        public string Barcode { get; init; } = string.Empty;
    }
    
    internal class Product
    {
        public string VendorCode1 { get; init; }
        public string Color { get; init; }
        public List<ProductItem> Items { get; init; }
    }
}
