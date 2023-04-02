using ExcelParser.Models;

namespace ExcelParser.Providers;

internal class GroupedByVendorCode2Provider
{
    public List<ResultTableRow> ApplyGrouping(IEnumerable<ProductItem> items)
    {
        return items
            .GroupBy(kvp => kvp.VendorCode2, kvp => kvp)
            .Select(grp => new ResultTableRow()
            {
                VendorCode2 = grp.Key,
                Name = grp.First().Name,
                Color = grp.First().Color,
                Area = grp.First().Area,
                Count = grp.Sum(x => x.Count),
                Barcode = grp.First().Barcode
            }).ToList();
    }
}