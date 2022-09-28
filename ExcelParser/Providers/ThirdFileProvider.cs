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

        internal List<ResultTableRow> AppendBarcodes(List<ProductItem> rows)
        {
            using var package = new ExcelPackage(_cfs.FileName);

            var cells = _cfs.WorksheetName is null ?
                package.Workbook.Worksheets.First().Cells :
                package.Workbook.Worksheets.First(x => x.Name.Equals(_cfs.WorksheetName, StringComparison.CurrentCultureIgnoreCase)).Cells;

            int column = _cfs.VendorCode2ColumnNumber;
            List<(int row, string value)> vendorcodes = LoadAllCellsFromColumn(cells, column);

            //ZSKHS40 артикул совпадает
            Dictionary<string, string> barByVendorCodes = new();
            foreach (var codeWithRow in vendorcodes)
            {
                string barcode = GetStringFromCell(cells[codeWithRow.row, _cfs.BarcodeColumnNumber]);
                if (barByVendorCodes.TryGetValue(codeWithRow.value, out string existed))
                    Console.WriteLine($"Для артикула {codeWithRow.value} штрих-код уже сохранён. Использован - {existed}, пропущен - {barcode}.");
                else
                    barByVendorCodes.Add(codeWithRow.value, barcode);
            }
            //vendorcodes.ToDictionary(k => k.value, v => GetStringFromCell(cells[v.row, _cfs.BarcodeColumnNumber]));


            return rows.Select(x => new ResultTableRow()
            {
                VendorCode2 = x.VendorCode2,
                Count = x.Count,
                Name = x.Name,
                Barcode = GetBarcode(barByVendorCodes, x),
            }).ToList();
        }

        private string GetBarcode(Dictionary<string, string> barByVendorCodes, ProductItem x)
        {
            if (barByVendorCodes.TryGetValue(x.VendorCode2, out string barcode))
                return barcode.PadLeft(_barcodeLength, '0');
            return _defaultBarcode;
        }
    }
}
