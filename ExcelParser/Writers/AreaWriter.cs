using ExcelParser.Models;
using OfficeOpenXml;

namespace ExcelParser.Writers;

public class AreaWriter
{
    private readonly SolutionDocument _settings;

    public AreaWriter(Settings settings)
    {
        _settings = settings.SolutionDocument;
    }

    public void Write(IEnumerable<ResultTableRow> rows, string sourceFileName)
    {
        using var package = new ExcelPackage();
        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(_settings.WorksheetName);
        var cells = worksheet.Cells;
        cells[1, 1].Value = "Цвет";
        cells[1, 2].Value = "Площадь общая";
        for (int i = 1; i <= 2; i++)
            worksheet.Column(i).StyleName = "Text";
        int iterator = 2;

        var group = rows.GroupBy(x => x.Color)
            .Select(grp => (color: grp.Key, total: grp.Sum(x => x.Area * x.Count)));

        foreach (var row in group)
        {
            cells[iterator, 1].Value = row.color;
            cells[iterator, 2].Value = row.total / 10_000_000;
            iterator++;
        }

        for (int i = 1; i <= 2; i++)
            worksheet.Column(i).AutoFit();
        string dir = Path.GetDirectoryName(_settings.ExcelFolder) ?? string.Empty;
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
        return Path.Combine(dir, "Area_" + sourceFileName + timeStamp) + ".xlsx";
    }
}