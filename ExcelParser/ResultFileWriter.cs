using ExcelParser.Models;
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
            cells[1, 1].Value = _settings.VendorCode2Header;
            cells[1, 2].Value = _settings.NameHeader;
            cells[1, 3].Value = _settings.CountHeader;
            cells[1, 4].Value = _settings.BarcodeHeader;

            for (int i = 1; i <= 4; i++)
                worksheet.Column(i).StyleName = "Text";

            int iterator = 2;
            foreach (var row in rows)
            {
                cells[iterator, 1].Value = row.VendorCode2;
                cells[iterator, 2].Value = row.Name;
                cells[iterator, 3].Value = row.Count;
                cells[iterator, 4].Value = row.Barcode.PadLeft(11, '0');
                iterator++;
            }

            for (int i = 1; i <= 5; i++)
                worksheet.Column(i).AutoFit();

            string dir = Path.GetDirectoryName(_settings.SolutionFolder) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            string fileName = GetFileName(sourceFileName, dir);
            if (File.Exists(fileName))
                File.Delete(fileName);

            package.SaveAs(fileName);
        }

        private string GetFileName(string sourceFileName, string dir)
        {
            string prefix = _settings.SolutionFileNamePrefix;
            string timeStamp = $" {DateTime.Now:dd.MM.yyyy HH-mm-ss}";
            string suffix = _settings.SolutionFileNameSuffix;
            return Path.Combine(dir, prefix + sourceFileName + timeStamp + suffix) + ".xlsx";
        }
    }
}
