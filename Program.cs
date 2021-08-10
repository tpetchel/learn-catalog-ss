using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

// hardcoded; this project is a one-off solution
var repos = new string[] {
    "/Users/thpetche/Dev/MicrosoftDocs/learn-pr",
    "/Users/thpetche/Dev/MicrosoftDocs/learn-m365-pr",
    "/Users/thpetche/Dev/MicrosoftDocs/learn-dynamics-pr",
    "/Users/thpetche/Dev/MicrosoftDocs/learn-bizapps-pr",
};

using (var workbook = new XLWorkbook())
{
    var allMetadata = new List<(string, IEnumerable<ModuleMetadata>)>();
    foreach (var repo in repos)
    {
        string repoName = repo.Split('/').Last();
        var metadataList = CollectModuleMetadata(repo);
        allMetadata.Add((repoName, metadataList));
    }

    var overview = workbook.Worksheets.Add("Overview");
    WriteOverview(overview, allMetadata);

    var products = workbook.Worksheets.Add("By Product");
    WriteProducts(products, allMetadata);
    var productCount = workbook.Worksheets.Add("By Product Count");
    WriteProductCounts(productCount, allMetadata);
    var freshness = workbook.Worksheets.Add("Freshness");
    WriteFreshness(freshness, allMetadata);

    workbook.SaveAs("LearnCatalog2.xlsx");
}

static IEnumerable<ModuleMetadata> CollectModuleMetadata(string repo)
{
    var metadataList =  new List<ModuleMetadata>();
    var files = Directory.GetFiles(repo, "index.yml", SearchOption.AllDirectories);
    foreach (string file in files)
    {
        // Skip learning path files.
        if (file.Contains("learn-pr/paths/"))
        {
            continue;
        }
        var metadata = ModuleMetadata.Load(file);
        metadataList.Add(metadata);
    }
    return metadataList;
}

static void WriteOverview(IXLWorksheet worksheet, IEnumerable<(string, IEnumerable<ModuleMetadata>)> allMetadata)
{
    worksheet.Cell(1, 1).Value = "Title";
    worksheet.Cell(1, 2).Value = "Repo";
    worksheet.Cell(1, 3).Value = "Author";
    worksheet.Cell(1, 4).Value = "Date";
    worksheet.Cell(1, 5).Value = "Products";

    var firstRow = worksheet.Row(1);
    firstRow.Style.Fill.BackgroundColor = XLColor.Black;
    firstRow.Style.Font.Bold = true;
    firstRow.Style.Font.FontSize = 16;
    firstRow.Style.Font.FontColor = XLColor.Aqua;
    worksheet.SheetView.FreezeRows(1);

    int index = 2;
    foreach (var (repo, metadataList) in allMetadata)
    {
        foreach (var metadata in metadataList)
        {
            worksheet.Row(index).Style.Font.FontSize = 12;
            worksheet.Cell(index, 1).Value = metadata.Title;
            worksheet.Cell(index, 2).Value = repo;
            worksheet.Cell(index, 3).Value = metadata.Author;
            worksheet.Cell(index, 4).Value = metadata.Date;
            var products = new List<string>(metadata.Products);
            products.Sort();
            worksheet.Cell(index, 5).Value = string.Join(',', products);
            index++;
        }
    }

    foreach (var column in worksheet.ColumnsUsed())
    {
        column.AdjustToContents();
    }
}

static void WriteProducts(IXLWorksheet worksheet, IEnumerable<(string, IEnumerable<ModuleMetadata>)> allMetadata)
{
    worksheet.Cell(1, 1).Value = "Product";
    worksheet.Cell(1, 2).Value = "Title";
    worksheet.Cell(1, 3).Value = "Repo";
    worksheet.Cell(1, 4).Value = "Author";
    worksheet.Cell(1, 5).Value = "Date";

    var firstRow = worksheet.Row(1);
    firstRow.Style.Fill.BackgroundColor = XLColor.Black;
    firstRow.Style.Font.Bold = true;
    firstRow.Style.Font.FontSize = 16;
    firstRow.Style.Font.FontColor = XLColor.Aqua;
    worksheet.SheetView.FreezeRows(1);

    Dictionary<string, List<(string Title, string Repo, string Author, string Date)>> items = new();
    foreach (var (repo, metadataList) in allMetadata)
    {
        foreach (var metadata in metadataList)
        {
            foreach (var product in metadata.Products)
            {
                if (!items.ContainsKey(product))
                {
                    items.Add(product, new List<(string Title, string Repo, string Author, string Date)>());
                }
                items[product].Add((Title: metadata.Title, Repo: repo, Author: metadata.Author, Date: metadata.Date));
            }
        }
    }

    int index = 2;
    foreach (var kvp in items)
    {
        var product = kvp.Key;
        foreach (var item in kvp.Value)
        {
            worksheet.Row(index).Style.Font.FontSize = 12;
            worksheet.Cell(index, 1).Value = product;
            worksheet.Cell(index, 2).Value = item.Title;
            worksheet.Cell(index, 3).Value = item.Repo;
            worksheet.Cell(index, 4).Value = item.Author;
            worksheet.Cell(index, 5).Value = item.Date;
            index++;
        }
    }

    foreach (var column in worksheet.ColumnsUsed())
    {
        column.AdjustToContents();
    }
}

static void WriteProductCounts(IXLWorksheet worksheet, IEnumerable<(string, IEnumerable<ModuleMetadata>)> allMetadata)
{
    worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;

    worksheet.Cell(1, 1).Value = "Product";
    worksheet.Cell(1, 2).Value = "Repo";
    worksheet.Cell(1, 3).Value = "Count";
    worksheet.Cell(1, 4).Value = "Subtotal";

    var firstRow = worksheet.Row(1);
    firstRow.Style.Fill.BackgroundColor = XLColor.Black;
    firstRow.Style.Font.Bold = true;
    firstRow.Style.Font.FontSize = 16;
    firstRow.Style.Font.FontColor = XLColor.Aqua;
    worksheet.SheetView.FreezeRows(1);

    Dictionary<string, Dictionary<string, int>> items = new();
    foreach (var (repo, metadataList) in allMetadata)
    {
        foreach (var metadata in metadataList)
        {
            foreach (var product in metadata.Products)
            {
                if (!items.ContainsKey(product))
                {
                    items.Add(product, new());
                }
                if (!items[product].ContainsKey(repo))
                {
                    items[product].Add(repo, 0);
                }
                items[product][repo]++;
            }
        }
    }

    int index = 2;
    foreach (var kvp in items.OrderBy(p => p.Key))
    {
        var product = kvp.Key;
        var startIndex = index;
        int subtotal = 0;
        foreach (var repo in kvp.Value)
        {
            worksheet.Row(index).Style.Font.FontSize = 12;
            worksheet.Cell(index, 1).Value = product;
            worksheet.Cell(index, 2).Value = repo.Key;
            worksheet.Cell(index, 3).Value = repo.Value;
            subtotal += repo.Value;
            index++;
        }
        worksheet.Rows(startIndex, index - 1).Group();
        worksheet.Cell(index, 4).Value = subtotal;
        worksheet.Cell(index, 4).Style.Font.Bold = true;
        worksheet.Cell(index, 4).Style.Fill.BackgroundColor = XLColor.Aqua;
        index++;
        //worksheet.Cell(2, 6).SetFormulaA1("=SUM(D2:D4)");
    }

    foreach (var column in worksheet.ColumnsUsed())
    {
        column.AdjustToContents();
    }
}

static void WriteFreshness(IXLWorksheet worksheet, IEnumerable<(string, IEnumerable<ModuleMetadata>)> allMetadata)
{
    (string Label, DateTime StartDate)[] timeCategories = {
        ("3 Months",  DateTime.Today.AddMonths(-3)),
        ("6 Months",  DateTime.Today.AddMonths(-6)),
        ("12 Months", DateTime.Today.AddMonths(-12)),
        ("18 Months", DateTime.Today.AddMonths(-18)),
        ("24 Months", DateTime.Today.AddMonths(-24)),
        ("Older",     DateTime.MinValue)
    };

    worksheet.Cell(1, 1).Value = "Updated Within";
    worksheet.Cell(1, 2).Value = "Repo";
    worksheet.Cell(1, 3).Value = "Count";
    worksheet.Cell(1, 4).Value = "Subtotal";

    var firstRow = worksheet.Row(1);
    firstRow.Style.Fill.BackgroundColor = XLColor.Black;
    firstRow.Style.Font.Bold = true;
    firstRow.Style.Font.FontSize = 16;
    firstRow.Style.Font.FontColor = XLColor.Aqua;
    worksheet.SheetView.FreezeRows(1);

    var buckets = new Dictionary<string, int>[timeCategories.Length];
    for (int i = 0; i < buckets.Length; i++)
    {
        buckets[i] = new();
    }

    foreach (var (repo, metadataList) in allMetadata)
    {
        foreach (var metadata in metadataList)
        {
            DateTime date;
            try
            {
                date = DateTime.Parse(metadata.Date);
            }
            catch (System.FormatException)
            {
                continue;
            }

            int categoryIndex = -1;
            for (int i = 0; i < timeCategories.Length; i++)
            {
                //System.Console.WriteLine($"{date.ToShortDateString()} : {timeCategories[i].StartDate.ToShortDateString()}");
                if (date >= timeCategories[i].StartDate)
                {
                    categoryIndex = i;
                    break;
                }
            }
            System.Diagnostics.Debug.Assert(categoryIndex >= 0);

            var bucket = buckets[categoryIndex];
            if (!bucket.ContainsKey(repo))
            {
                bucket.Add(repo, 0);
            }
            bucket[repo]++;
        }
    }

    int index = 2;
    for (int i = 0; i < timeCategories.Length; i++)
    {
        var category = timeCategories[i];
        var bucket = buckets[i];

        worksheet.Row(index).Style.Font.FontSize = 12;
        worksheet.Cell(index, 1).Value = category.Label;
        index++;

        var startIndex = index;
        int subtotal = 0;
        foreach (var kvp in bucket)
        {
            worksheet.Row(index).Style.Font.FontSize = 12;
            worksheet.Cell(index, 2).Value = kvp.Key;
            worksheet.Cell(index, 3).Value = kvp.Value;
            subtotal += kvp.Value;
            index++;
        }
        worksheet.Rows(startIndex, index - 1).Group();
        worksheet.Cell(index, 4).Value = subtotal;
        worksheet.Cell(index, 4).Style.Font.Bold = true;
        worksheet.Cell(index, 4).Style.Fill.BackgroundColor = XLColor.Aqua;
        index++;
        //worksheet.Cell(2, 6).SetFormulaA1("=SUM(D2:D4)");
    }

    foreach (var column in worksheet.ColumnsUsed())
    {
        column.AdjustToContents();
    }
}
