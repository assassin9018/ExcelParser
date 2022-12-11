using ExcelParser.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelParser.Writers;

internal class ReportExcelWriter : ResultExcelWriter
{
    private const int Height = 3;
    private const int Width = 5;

    public ReportExcelWriter(Settings settings) : base(settings)
    {
    }

    protected override string GetResultDir()
        => Settings.ReportFolder;


    protected override int WriteBody(ExcelWorksheet worksheet, IEnumerable<ResultTableRow> rows, int iterator)
    {
        int start = iterator;
        var cells = worksheet.Cells;

        foreach (var row in rows)
        {
            cells[iterator, 1].Value = row.VendorCode2;
            cells[iterator, 2].Value = row.Name;
            cells[iterator, 3].Value = row.Count;
            cells[iterator, 4].Value = row.Count;
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

    protected override int BeforeWrite(ExcelWorksheet worksheet, int iterator)
    {
        var cells = worksheet.Cells;

        var act = cells[iterator, 1, iterator, Width];
        act.Merge = true;
        act.Style.Font.Bold = true;
        act.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        act.Value = "Акт";
        iterator++;

        var date = cells[iterator, 1, iterator, Width];
        date.Merge = true;
        date.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        date.Value = $"Дата: {DateTime.Now:dd-MM-yyyy}";
        iterator++;

        var act2 = cells[iterator, 1, iterator, Width];
        act2.Merge = true;
        act2.Style.Font.Bold = true;
        act2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        act2.Value = "Приема-передачи товара";
        iterator++;

        var header = cells[iterator, 1, iterator + Height, Width];
        header.Merge = true;
        header.Value =
            @"Брамф именуемый в дальнейшем Поставщик,клиент, именуемое в дальнейшем Заказчик __________________________________,
именуемое в дальнейшем Заказчик с другой стороны (в дальнейшем вместе именуемые «Стороны» и по отдельности «Сторона»), составили настоящий
Акт о нижеследующем:
В соответствии Поставщик передает Водителю, а Водитель передает и Заказчик принимает Товар следующего ассортимента и количества:";

        ApplyStyle(header);

        iterator += Height + 2;

        cells[iterator, 1].Value = "Название";
        cells[iterator, 2].Value = "Расшифровка";
        cells[iterator, 3].Value = "Количество\nотправленный\nсо склада";
        cells[iterator, 4].Value = "Количество\nпринято\nВодителем";
        cells[iterator, 5].Value = "Количество\nпринято\nЗаказчиком";

        for (int i = 1; i <= 5; i++)
            worksheet.Column(i).StyleName = "Text";
        worksheet.Row(iterator).Height = 30;

        return ++iterator;
    }

    protected override int AfterWrite(ExcelWorksheet worksheet, int iterator)
    {
        iterator += 2;
        var footer = worksheet.Cells[iterator, 1, iterator + Height, Width];
        footer.Merge = true;
        footer.Value =
            @"2. Принятый Заказчиком товар обладает качеством и ассортиментом. Заказчик не имеет никаких претензий к принятому им товару,и подтверждает полученное количество товара от Водителя.

Поставщик                   Водитель                    Заказщик
________________            ______________________      __________________";
        ApplyStyle(footer);

        iterator += Height;

        return base.AfterWrite(worksheet, ++iterator);
    }

    private static void ApplyStyle(ExcelRange range)
    {
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        range.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
        range.Style.WrapText = true;
    }
}