namespace ExcelParser.Models
{
    internal class ResultTableRow : ProductItem
    {
        /// <summary>
        /// Штрих-код
        /// </summary>
        public string Barcode { get; set; } = string.Empty;
    }

    internal class ProductItem
    {
        /// <summary>
        /// Артикул 2
        /// </summary>
        public string VendorCode2 { get; set; } = string.Empty;
        /// <summary>
        /// Наименование
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// Количество
        /// </summary>
        public int Count { get; set; } = 0;
    }
}
