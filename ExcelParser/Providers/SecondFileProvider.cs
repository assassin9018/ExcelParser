using ExcelParser.Models;
using OfficeOpenXml;

namespace ExcelParser.Providers
{
    internal class SecondFileProvider : ExcelProviderBase
    {
        private readonly SecondDocument _cfs;

        public SecondFileProvider(Settings settings)
        {
            _cfs = settings.SecondDocument;
        }

        internal void AppendItems(List<Product> products)
        {
            using var package = new ExcelPackage(_cfs.FileName);

            var cells = GetCells(package, _cfs.WorksheetName);

            int column = _cfs.VendorCode1ColumnNumber;
            var valuesWithRows = LoadAllCellsFromColumn(cells, column);

            var values = valuesWithRows
                .Where(x => products.Any(p=>p.VendorCode1.Equals(x.value, StringComparison.OrdinalIgnoreCase)))
                .Select(x => new
                {
                    VendorCode1 = x.value,
                    VendorCode2 = GetStringFromCell(cells[x.row, _cfs.VendorCode2ColumnNumber]),
                    Name = GetStringFromCell(cells[x.row, _cfs.NameColumnNumber]),
                    Count = TryGetIntFromCell(cells[x.row, _cfs.CountColumnNumber]),
                })
                .GroupBy(k => k.VendorCode1, v => v)
                .ToDictionary(kvp => kvp.Key, kvp => kvp.ToList());
            
            foreach (var product in products)
                if (values.TryGetValue(product.VendorCode1, out var items))
                    product.Items.AddRange(items.Select(x=>new ProductItem
                    {
                        VendorCode2 = string.IsNullOrWhiteSpace(x.VendorCode2) ? x.Name : x.VendorCode2,
                        Count = x.Count,
                        Name = x.Name,
                    }));
                else
                    product.Items.Add(new ProductItem
                    {
                        VendorCode2 = $"Не найден {product.VendorCode1}",
                        Count = 0,
                        Name = $"Для Артикула 1 \"{product.VendorCode1}\" не был найден список деталей",
                    });
        }
    }
}
