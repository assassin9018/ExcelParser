using OfficeOpenXml;
using System.Runtime;

namespace ExcelParser.Providers
{
    internal class FirstFileProvider : ExcelProviderBase
    {
        private readonly FirstDocument _cfs;

        public FirstFileProvider(Settings settings)
{
            _cfs = settings.FirstDocument;
        }

        public List<string> LoadVendorCodes(string fileName)
        {
            using var package = new ExcelPackage(fileName);

            var cells = _cfs.WorksheetName is null ?
                package.Workbook.Worksheets.First().Cells :
                package.Workbook.Worksheets.First(x=>x.Name.Equals(_cfs.WorksheetName, StringComparison.CurrentCultureIgnoreCase)).Cells;

            int column = _cfs.VendorCodeColumnNumber;
            var valuesWithRows = LoadAllCellsFromColumn(cells, column);
            var filter = _cfs.WordsFilter;

            return valuesWithRows
                .Select(x => x.value)
                .Where(x => !filter.Any(sw => sw.Equals(x, StringComparison.CurrentCultureIgnoreCase)))
                .ToList();
        }
    }
}
