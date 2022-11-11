using ExcelParser.Models;
using OfficeOpenXml;

namespace ExcelParser.Providers
{
    internal class ThirdFileProvider : ExcelProviderBase
    {
        private readonly ThirdDocument _cfs;
        private readonly string _defaultBarcode;
        private readonly int _barcodeLength;

        public ThirdFileProvider(Settings settings)
        {
            _cfs = settings.ThirdDocument;
            _barcodeLength = settings.SolutionDocument.BarcodeLength;
            _defaultBarcode = "".PadLeft(_barcodeLength, '0');
        }

        internal void AppendBarcodesByColors(List<Product> products)
        {
            using var package = new ExcelPackage(_cfs.FileName);

            var barcodeByVendor = LoadBarcodesDictionary(package);
            var colorSuffixes = LoadColorSuffixes(package);

            var cips = _cfs.ColoredItemsPrefixes;
            foreach (var product in products)
            {
                Color? color = colorSuffixes.FirstOrDefault(x => x.Name.Equals(product.Color, StringComparison.OrdinalIgnoreCase));
                if (color is null)
                    continue;

                var items = product.Items;
                for (int i = 0; i < items.Count; i++)
                {
                    var current = items[i];
                    if (cips.Any(x => current.VendorCode2.StartsWith(x, StringComparison.OrdinalIgnoreCase)))
                        items[i] = current with { VendorCode2 = current.VendorCode2 + color.Suffix };
                }
            }

            foreach (var product in products)
            {
                var items = product.Items;
                for (int i = 0; i < items.Count; i++)
                    items[i] = items[i] with { Barcode = GetBarcode(barcodeByVendor, items[i]) };

            }
        }

        private Dictionary<string, string> LoadBarcodesDictionary(ExcelPackage package)
        {
            var cells = GetCells(package, _cfs.BarcodeWorksheetName);

            int column = _cfs.VendorCode2ColumnNumber;
            List<(int row, string value)> vendorcodes = LoadAllCellsFromColumn(cells, column);

            //ZSKHS40 артикул совпадает
            Dictionary<string, string> barcodeByVendor = new();
            foreach (var (row, value) in vendorcodes)
            {
                string barcode = GetStringFromCell(cells[row, _cfs.BarcodeColumnNumber]);
                if (barcodeByVendor.TryGetValue(value, out string? existed))
                    Console.WriteLine($"Для артикула {value} штрих-код уже сохранён. Использован - {existed}, пропущен - {barcode}.");
                else
                    barcodeByVendor.Add(value, barcode);
            }
            //vendorcodes.ToDictionary(k => k.value, v => GetStringFromCell(cells[v.row, _cfs.BarcodeColumnNumber]));
            return barcodeByVendor;
        }

        private List<Color> LoadColorSuffixes(ExcelPackage package)
        {
            var cells = GetCells(package, _cfs.ColorWorksheetName);
            List<(int row, string value)> colorWithRow = LoadAllCellsFromColumn(cells, _cfs.ColorColumnNumber);
            int sfxNum = _cfs.ColorSuffixColumnNumber;
            return colorWithRow
                .Select(x => new Color(x.value, GetStringFromCell(cells[x.row, sfxNum])))
                .ToList();
        }

        private string GetBarcode(Dictionary<string, string> barByVendorCodes, ProductItem x)
        {
            if (barByVendorCodes.TryGetValue(x.VendorCode2, out string? barcode))
                return barcode.PadLeft(_barcodeLength, '0');
            return _defaultBarcode;
        }

        private record Color(string Name, string Suffix);
    }
}
