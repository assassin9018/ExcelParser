// See https://aka.ms/new-console-template for more information
using ExcelParser;
using ExcelParser.Models;
using ExcelParser.Providers;
using ExcelParser.Writers;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
string[] supportedExtentions = new[] { ".xlsx", ".xls" };

try
{
    var settings = Settings.Load();
    FirstFileProvider firstProvider = new(settings);
    SecondFileProvider secondProvider = new(settings);
    GroupedByVendorCode2Provider groupProvider = new();
    ThirdFileProvider thirdProvider = new(settings);
    ResultExcelWriter excelWriter = new(settings);
    ResultDmWriter csvWriter = new(settings);

    var filesForParsing = Directory.EnumerateFiles(settings.FirstDocument.FodlerName)
        .Where(x => supportedExtentions.Any(s => s.Equals(Path.GetExtension(x), StringComparison.CurrentCultureIgnoreCase)));

    foreach (var fileName in filesForParsing)
    {
        Console.ForegroundColor = ConsoleColor.White;
        try
        {
            List<string> vendorcodes = firstProvider.LoadVendorCodes(fileName);
            List<ProductItem> items = secondProvider.LoadRowsWith(vendorcodes);
            List<ProductItem> groupedItems = groupProvider.ApplyGrouping(items);
            List<ResultTableRow> resultRows = thirdProvider.AppendBarcodes(groupedItems);

            string fileNameWithoutExtention = Path.GetFileNameWithoutExtension(fileName);
            if (settings.SolutionDocument.OutExcel)
                excelWriter.Write(resultRows, fileNameWithoutExtention);
            if (settings.SolutionDocument.OutDm)
                csvWriter.Write(resultRows, fileNameWithoutExtention);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Файл - {fileName} обработан.");
            File.Delete(fileName);
            Console.WriteLine($"Файл - {fileName} удалён.");

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
