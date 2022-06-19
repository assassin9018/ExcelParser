// See https://aka.ms/new-console-template for more information
using ExcelParser;
using ExcelParser.Providers;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
string[] supportedExtentions = new[] { ".xlsx", ".xls" };

try
{
    var settings = Settings.Load();
    FirstFileProvider firstProvider = new(settings);
    SecondFileProvider secondProvider = new(settings);
    ThirdFileProvider thirdProvider = new(settings);
    ResultFileWriter resultWriter = new(settings);

    var filesForParsing = Directory.EnumerateFiles(settings.FirstDocument.FodlerName)
        .Where(x => supportedExtentions.Any(s => s.Equals(Path.GetExtension(x), StringComparison.CurrentCultureIgnoreCase)));

    foreach (var fileName in filesForParsing)
    {
        Console.ForegroundColor = ConsoleColor.White;
        try
        {
            List<string> vendorcodes = firstProvider.LoadVendorCodes(fileName);
            List<ResultTableRow> rows = secondProvider.LoadRowsWith(vendorcodes);
            thirdProvider.AppendBarcodes(rows);

            resultWriter.Write(rows, Path.GetFileNameWithoutExtension(fileName));
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Файл - {fileName} обработан.");
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Не удалось обработать файл - {fileName}.");
            Console.WriteLine(ex.Message);
        }
    }
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("Работа завершена!");
}
catch (Exception ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Работа завершена с ошибкой!");
    Console.WriteLine(ex.Message);
}
Console.ReadKey();
