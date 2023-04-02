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

            var barcodeByVendor = LoadBarcodesAndSizesDictionary(package);
            var colorSuffixes = LoadColorSuffixes(package);

            var cips = _cfs.ColoredItemsPrefixes;
            foreach (var product in products)
            {
                Color? color = colorSuffixes.FirstOrDefault(x =>
                    x.Name.Equals(product.Color, StringComparison.OrdinalIgnoreCase));
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
                {
                    var value = ExtractValue(barcodeByVendor, items[i]);
                    items[i] = items[i] with { Barcode = value.barcode, Area = value.area };
                }
            }
        }

        private Dictionary<string, (string barcode, int area)> LoadBarcodesAndSizesDictionary(ExcelPackage package)
        {
            var cells = GetCells(package, _cfs.BarcodeWorksheetName);

            int column = _cfs.VendorCode2ColumnNumber;
            var vendorcodes = LoadAllCellsFromColumn(cells, column);

            Dictionary<string, (string barcode, int area)> barcodeByVendor = new();
            foreach (var (row, value) in vendorcodes)
            {
                string barcode = GetStringFromCell(cells[row, _cfs.BarcodeColumnNumber]);
                int width = TryGetIntFromCell(cells[row, _cfs.WidthColumnNumber]);
                int height = TryGetIntFromCell(cells[row, _cfs.HeightColumnNumber]);
                if (barcodeByVendor.TryGetValue(value, out var existed))
                    Console.WriteLine(
                        $"Для артикула {value} штрих-код уже сохранён. Использован - {existed}, пропущен - {barcode}.");
                else
                    barcodeByVendor.Add(value, (barcode, width * height));
            }

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

        private (string barcode, int area) ExtractValue(
            Dictionary<string, (string barcode, int area)> barAndAreaByVendorCodes, ProductItem x)
        {
            if (barAndAreaByVendorCodes.TryGetValue(x.VendorCode2, out var value))
                return (value.barcode.PadLeft(_barcodeLength, '0'), value.area);
            return (_defaultBarcode, 0);
        }

        private record Color(string Name, string Suffix);
    }
}