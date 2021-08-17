using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

class ProductCSAOwner
{
    public string Slug;
    public string CSA;
    public string M2;
}

static class ProductCSAOwnerTable
{
    // Maps product slug -> Product
    public static Dictionary<string, ProductCSAOwner> Mapping = new();

    public static void Load(string workbookPath)
    {
        var productMapping = Mapping;

        using (var workbook = new XLWorkbook(workbookPath))
        {
            var worksheet = workbook.Worksheets.Where(ws => ws.Name == "mapping").Single();
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                productMapping[row.Cell(1).Value.ToString()] = new ProductCSAOwner() {
                    Slug = row.Cell(1).Value.ToString(),
                    CSA =  row.Cell(2).Value.ToString(),
                    M2 = row.Cell(3).Value.ToString()
                };
            }
        }
    }

    public static void Dump()
    {
        foreach (var kvp in Mapping)
        {
            var v = kvp.Value;
            Console.WriteLine($"\"{v.Slug}\",\"{v.CSA}\",\"{v.M2}\"");
        }
    }

    public static IEnumerable<string> MapProductsToCSAs(IEnumerable<string> productNames)
    {
        List<string> list = new();
        foreach (var productName in productNames)
        {
            ProductCSAOwner product;
            if (Mapping.TryGetValue(productName, out product))
            {
                list.Add(product.CSA);
            }
        }
        return list;
    }
}