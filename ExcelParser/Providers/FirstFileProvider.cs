using OfficeOpenXml;
using ExcelParser.Models;

namespace ExcelParser.Providers
{
    internal class FirstFileProvider : ExcelProviderBase
    {
        private readonly FirstDocument _cfs;

        public FirstFileProvider(Settings settings)
{
            _cfs = settings.FirstDocument;
        }

        public List<Product> LoadVendorCodes(string fileName)
        {
            using var package = new ExcelPackage(fileName);

            var cells = GetCells(package, _cfs.WorksheetName);

            int column = _cfs.VendorCodeColumnNumber;
            var valuesWithRows = LoadAllCellsFromColumn(cells, column);
            var filter = _cfs.WordsFilter;

            return valuesWithRows
                .Where(x => !filter.Any(sw => sw.Equals(x.value, StringComparison.CurrentCultureIgnoreCase)))
                .Select(x => new Product
                {
                    VendorCode1 = x.value,
                    Color = GetStringFromCell(cells[x.row + 1, _cfs.ColorColumnNumber]),
                    Items = new()
                })
                .ToList();
        }
    }
}
