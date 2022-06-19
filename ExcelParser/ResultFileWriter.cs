using OfficeOpenXml;

namespace ExcelParser
{
    internal class ResultFileWriter
    {
        private SolutionDocument _settings;

        public ResultFileWriter(Settings settings)
        {
            _settings = settings.SolutionDocument;
        }

        public void Write(IEnumerable<ResultTableRow> rows, string sourceFileName)
        {
            using var package = new ExcelPackage();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(_settings.WorksheetName);

            var cells = worksheet.Cells;
            cells[1, 1].Value = _settings.VendorCode1Header;
            cells[1, 2].Value = _settings.VendorCode2Header;
            cells[1, 3].Value = _settings.NameHeader;
            cells[1, 4].Value = _settings.CountHeader;
            cells[1, 5].Value = _settings.BarcodeHeader;

            int iterator = 2;
            foreach (var row in rows)
            {
                cells[iterator, 1].Value = row.VendorCode1;
                cells[iterator, 2].Value = row.VendorCode2;
                cells[iterator, 3].Value = row.Name;
                cells[iterator, 4].Value = row.Count;
                cells[iterator, 5].Value = row.Barcode;
                iterator++;
            }

            string dir = Path.GetDirectoryName(_settings.SolutionFileName) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            string fileName = Path.Combine(dir, sourceFileName + " - " + Path.GetFileName(_settings.SolutionFileName));
            if (File.Exists(fileName))
                File.Delete(fileName);

            package.SaveAs(fileName);
        }
    }
}
