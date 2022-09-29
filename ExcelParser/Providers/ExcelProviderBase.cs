using OfficeOpenXml;

namespace ExcelParser.Providers
{
    internal class ExcelProviderBase
    {
        protected static ExcelRange GetCells(ExcelPackage package, string? worksheetName)
        {
            return worksheetName is null ?
                package.Workbook.Worksheets.First().Cells :
                package.Workbook.Worksheets.First(x
                        =>x.Name.Equals(worksheetName, StringComparison.CurrentCultureIgnoreCase))
                    .Cells;
        }
        
        protected static List<(int row, string value)> LoadAllCellsFromColumn(ExcelRange cells, int column)
        {
            return cells.Where(x => x.Start.Column == column)
                                .Select(x => (x.Start.Row, GetStringFromCell(cells[x.Start.Row, column])))
                                .Where(x => x.Item2.Length > 0)
                                .ToList();
        }

        protected static int TryGetIntFromCell(ExcelRange excelRange)
        {
            if (excelRange.Value is int val)
                return val;
            if (excelRange.Value is double dbl)
                return (int)dbl;
            if (excelRange.Value is string str && int.TryParse(str.Trim(), out val))
                return val;

            return - 1;
        }


        protected static string GetStringFromCell(ExcelRange cell)
        {
            if(cell.Value is string str)
                return str.Trim();

            if (cell.Value is double dbl)
                return dbl.ToString();

            return cell.Value?.ToString() ?? string.Empty;
        }
    }
}