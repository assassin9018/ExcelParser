using ExcelParser.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelParser.Writers
{
    internal class ResultExcelWriter
    {
        protected readonly SolutionDocument Settings;

        public ResultExcelWriter(Settings settings)
        {
            Settings = settings.SolutionDocument;
        }

        public void WriteDelimited(IEnumerable<ResultTableRow> rows, string sourceFileName)
        {
            Func<ResultTableRow, bool> NonKFilter = x =>
                x.VendorCode2.StartsWith("k", StringComparison.OrdinalIgnoreCase)
                || x.Name.Equals(x.VendorCode2);
            Write(rows.Where(NonKFilter), sourceFileName + "_nonK");
            Write(rows.Where(x => !NonKFilter(x)), sourceFileName + "_other");
        }

        public void Write(IEnumerable<ResultTableRow> rows, string sourceFileName)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(Settings.WorksheetName);

            int iterator = BeforeWrite(worksheet, 1);
            iterator = WriteBody(worksheet, rows, iterator);
            AfterWrite(worksheet, iterator);

            for (int i = 1; i <= 5; i++)
                worksheet.Column(i).AutoFit();

            string dir = Path.GetDirectoryName(GetResultDir()) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            string fileName = GetFileName(sourceFileName, dir);
            if (File.Exists(fileName))
                File.Delete(fileName);

            package.SaveAs(fileName);
        }

        protected virtual int WriteBody(ExcelWorksheet worksheet, IEnumerable<ResultTableRow> rows, int iterator)
        {
            int start = iterator;
            var cells = worksheet.Cells;

            foreach (var row in rows)
            {
                cells[iterator, 1].Value = row.VendorCode2;
                cells[iterator, 2].Value = row.Name;
                cells[iterator, 3].Value = row.Count;
                cells[iterator, 4].Value = row.Barcode;
                iterator++;
            }

            int end = iterator;
            var border = cells[start, 1, end, 5].Style.Border;
            border.Bottom.Style = ExcelBorderStyle.Thick;
            border.Top.Style = ExcelBorderStyle.Thick;
            border.Left.Style = ExcelBorderStyle.Thick;
            border.Right.Style = ExcelBorderStyle.Thick;
            return iterator;
        }

        protected virtual int BeforeWrite(ExcelWorksheet worksheet, int iterator)
        {
            var cells = worksheet.Cells;

            cells[iterator, 1].Value = Settings.VendorCode2Header;
            cells[iterator, 2].Value = Settings.NameHeader;
            cells[iterator, 3].Value = Settings.CountHeader;
            cells[iterator, 4].Value = Settings.BarcodeHeader;

            for (int i = 1; i <= 5; i++)
                worksheet.Column(i).StyleName = "Text";
            return ++iterator;
        }

        protected virtual int AfterWrite(ExcelWorksheet worksheet, int iterator)
            => iterator;

        protected virtual string GetResultDir()
            => Settings.ExcelFolder;

        private string GetFileName(string sourceFileName, string dir)
        {
            string prefix = Settings.SolutionFileNamePrefix;
            string timeStamp = $" {DateTime.Now:dd.MM.yyyy HH-mm-ss}";
            string suffix = Settings.SolutionFileNameSuffix;
            return Path.Combine(dir, prefix + sourceFileName + timeStamp + suffix) + ".xlsx";
        }
    }
}