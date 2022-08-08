using ExcelParser.Models;

internal class GroupedByVendorCode2Provider
{
    internal List<ProductItem> ApplyGrouping(List<ProductItem> items)
    {
        return items
            .GroupBy(kvp => kvp.VendorCode2, kvp => kvp)
            .Select(x => new ProductItem()
            {
                VendorCode2 = x.Key,
                Name = x.First().Name,
                Count = x.Sum(x => x.Count)
            }).ToList();
    }
}