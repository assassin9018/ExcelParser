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
    ReportExcelWriter reportWriter = new(settings);
    ResultDmWriter csvWriter = new(settings);

    string ordersFolder = Path.Combine(Directory.GetCurrentDirectory(), settings.FirstDocument.FodlerName);
    var filesForParsing = Directory.EnumerateFiles(ordersFolder)
        .Where(x => supportedExtentions.Any(s => s.Equals(Path.GetExtension(x), StringComparison.CurrentCultureIgnoreCase)));

    foreach (var fileName in filesForParsing)
    {
        Console.ForegroundColor = ConsoleColor.White;
        try
        {
            List<Product> products = firstProvider.LoadVendorCodes(fileName);
            secondProvider.AppendItems(products);
            thirdProvider.AppendBarcodesByColors(products);
            List<ResultTableRow> resultRows = groupProvider.ApplyGrouping(products.SelectMany(x=>x.Items));
            resultRows = resultRows.OrderBy(x => x.Name).ToList()
                ;
            string fileNameWithoutExtention = Path.GetFileNameWithoutExtension(fileName);
            if (settings.SolutionDocument.OutExcel)
                excelWriter.Write(resultRows, fileNameWithoutExtention);
            if(settings.SolutionDocument.OutReport)
                reportWriter.Write(resultRows, fileNameWithoutExtention);
            if (settings.SolutionDocument.OutDm)
                csvWriter.Write(resultRows, fileNameWithoutExtention);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Файл - {fileName} обработан.");
            File.Delete(fileName);
            Console.WriteLine($"Файл - {fileName} удалён.");
            Increment(settings);
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


static void Increment(Settings settings)
{
    settings.SolutionDocument.Iterator++;
    settings.Save();
}
