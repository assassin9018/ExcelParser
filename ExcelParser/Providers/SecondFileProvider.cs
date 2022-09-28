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

        internal List<ProductItem> LoadRowsWith(List<string> vendorcodes)
        {
            using var package = new ExcelPackage(_cfs.FileName);

            var cells = _cfs.WorksheetName is null ?
                package.Workbook.Worksheets.First().Cells :
                package.Workbook.Worksheets.First(x => x.Name.Equals(_cfs.WorksheetName, StringComparison.CurrentCultureIgnoreCase)).Cells;

            int column = _cfs.VendorCode1ColumnNumber;
            var valuesWithRows = LoadAllCellsFromColumn(cells, column);

            var values = valuesWithRows
                .Where(x => vendorcodes.Contains(x.value))
                .Select(x => new
                {
                    VendorCode1 = x.value,
                    VendorCode2 = GetStringFromCell(cells[x.row, _cfs.VendorCode2ColumnNumber]),
                    Name = GetStringFromCell(cells[x.row, _cfs.NameColumnNumber]),
                    Count = TryGetIntFromCell(cells[x.row, _cfs.CountColumnNumber]),
                })
                .GroupBy(k => k.VendorCode1, v => v)
                .ToDictionary(kvp => kvp.Key, kvp => kvp.ToList());


            List<ProductItem> result = new();
            foreach (var vendorCode1 in vendorcodes)
            {
                if (values.TryGetValue(vendorCode1, out var items))
                    foreach (var item in items)
                    {
                        result.Add(new ProductItem()
                        {
                            VendorCode2 = string.IsNullOrWhiteSpace(item.VendorCode2) ? item.Name : item.VendorCode2,
                            Count = item.Count,
                            Name = item.Name,
                        });
                    }
                else
                    result.Add(new ProductItem()
                    {
                        VendorCode2 = $"Не найден {vendorCode1}",
                        Count = 0,
                        Name = $"Для Артикула 1 \"{vendorCode1}\" не был найден список деталей",
                    });

            }

            return result;
        }
    }
}
