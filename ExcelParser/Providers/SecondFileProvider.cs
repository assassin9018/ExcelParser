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

        internal List<ResultTableRow> LoadRowsWith(List<string> vendorcodes)
        {
            using var package = new ExcelPackage(_cfs.FileName);

            var cells = _cfs.WorksheetName is null ?
                package.Workbook.Worksheets.First().Cells :
                package.Workbook.Worksheets.First(x => x.Name.Equals(_cfs.WorksheetName, StringComparison.CurrentCultureIgnoreCase)).Cells;

            int column = _cfs.VendorCode1ColumnNumber;
            var valuesWithRows = LoadAllCellsFromColumn(cells, column);

            HashSet<string> codesForSearch = new(vendorcodes);

            return valuesWithRows
                .Where(x => codesForSearch.Contains(x.value))
                .Select(x => new ResultTableRow()
                {
                    VendorCode1 = x.value,
                    VendorCode2 = GetStringFromCell(cells[x.row, _cfs.VendorCode2ColumnNumber]),
                    Name = GetStringFromCell(cells[x.row, _cfs.NameColumnNumber]),
                    Count = TryGetIntFromCell(cells[x.row, _cfs.CountColumnNumber]),
                })
                .ToList();
        }
    }
}
