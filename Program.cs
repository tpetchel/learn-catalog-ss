using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

const string repoBasePath = "/Users/thpetche/Dev/MicrosoftDocs/";

// hardcoded; this project is a one-off solution
var repos = new string[] {
    $"{repoBasePath}learn-pr",
    $"{repoBasePath}learn-m365-pr",
    $"{repoBasePath}learn-dynamics-pr",
    $"{repoBasePath}learn-bizapps-pr",
};

// Load CSA to Learn Product mapping.
ProductCSAOwnerTable.Load("CSA to Learn Product mapping.xlsx");
//ProductsTable.Dump();

string contentOwnersMarkdownFile = $"{repoBasePath}docs-help-pr/help-content/contribute/faq-service-content-owners.md";
CSAApproversTable.Load(contentOwnersMarkdownFile);
//CSAApproversTable.Dump();

ProductTaxonomyTable.Load("product-taxonomy.json");
//ProductTaxonomyTable.Dump();

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

    var recommendations = workbook.Worksheets.Add("Recommendations");
    WriteRecommendations(recommendations, Query.RecommendProducts(allMetadata));

    workbook.SaveAs("LearnCatalog.xlsx");
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
        PathTable.FilePaths[metadata.Uid] = file.Remove(0, repoBasePath.Length);
        metadataList.Add(metadata);
    }
    return metadataList;
}

static void WriteOverview(IXLWorksheet worksheet, IEnumerable<(string, IEnumerable<ModuleMetadata>)> allMetadata)
{
    worksheet.Cell(1, 1).Value = "Title";
    worksheet.Cell(1, 2).Value = "Repo";
    worksheet.Cell(1, 3).Value = "Author";
    worksheet.Cell(1, 4).Value = "ms.author";
    worksheet.Cell(1, 5).Value = "ms.date";
    worksheet.Cell(1, 6).Value = "Products";
    worksheet.Cell(1, 7).Value = "CSAs";
    worksheet.Cell(1, 8).Value = "Owner";

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
            worksheet.Cell(index, 4).Value = metadata.MsAuthor;
            worksheet.Cell(index, 5).Value = metadata.Date;
            var products = new List<string>(metadata.Products);
            products.Sort();
            worksheet.Cell(index, 6).Value = string.Join(',', products);
            var productCSAs = ProductCSAOwnerTable.MapProductsToCSAs(products);
            worksheet.Cell(index, 7).Value = string.Join(',', productCSAs);
            var approvers = Query.GetApprovers(products);
            worksheet.Cell(index, 8).Value = string.Join(',', approvers);
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
    worksheet.Cell(1, 2).Value = "CSA";
    worksheet.Cell(1, 3).Value = "Owner";
    worksheet.Cell(1, 4).Value = "Title";
    worksheet.Cell(1, 5).Value = "Repo";
    worksheet.Cell(1, 6).Value = "Author";
    worksheet.Cell(1, 7).Value = "ms.author";
    worksheet.Cell(1, 8).Value = "ms.date";

    var firstRow = worksheet.Row(1);
    firstRow.Style.Fill.BackgroundColor = XLColor.Black;
    firstRow.Style.Font.Bold = true;
    firstRow.Style.Font.FontSize = 16;
    firstRow.Style.Font.FontColor = XLColor.Aqua;
    worksheet.SheetView.FreezeRows(1);

    Dictionary<string, List<(string Title, string Repo, string Author, string MsAuthor, string Date)>> items = new();
    foreach (var (repo, metadataList) in allMetadata)
    {
        foreach (var metadata in metadataList)
        {
            foreach (var product in metadata.Products)
            {
                if (!items.ContainsKey(product))
                {
                    items.Add(product, new List<(string Title, string Repo, string Author, string MsAuthor, string Date)>());
                }
                items[product].Add((Title: metadata.Title, Repo: repo, Author: metadata.Author, MsAuthor: metadata.MsAuthor, Date: metadata.Date));
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
            worksheet.Cell(index, 2).Value = ProductCSAOwnerTable.MapProductsToCSAs(new[] {product});
            var approvers = Query.GetApprovers(new[] {product});
            worksheet.Cell(index, 3).Value = string.Join(',', approvers);
            worksheet.Cell(index, 4).Value = item.Title;
            worksheet.Cell(index, 5).Value = item.Repo;
            worksheet.Cell(index, 6).Value = item.Author;
            worksheet.Cell(index, 7).Value = item.MsAuthor;
            worksheet.Cell(index, 8).Value = item.Date;
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
    worksheet.Cell(1, 2).Value = "CSA";
    worksheet.Cell(1, 3).Value = "Owner";
    worksheet.Cell(1, 4).Value = "Repo";
    worksheet.Cell(1, 5).Value = "Count";
    worksheet.Cell(1, 6).Value = "Subtotal";

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
            worksheet.Cell(index, 2).Value = ProductCSAOwnerTable.MapProductsToCSAs(new[] {product});
            var approvers = Query.GetApprovers(new[] {product});
            worksheet.Cell(index, 3).Value = string.Join(',', approvers);
            worksheet.Cell(index, 4).Value = repo.Key;
            worksheet.Cell(index, 5).Value = repo.Value;
            subtotal += repo.Value;
            index++;
        }
        worksheet.Rows(startIndex, index - 1).Group();
        worksheet.Cell(index, 6).Value = subtotal;
        worksheet.Cell(index, 6).Style.Font.Bold = true;
        worksheet.Cell(index, 6).Style.Fill.BackgroundColor = XLColor.Aqua;
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

static void WriteRecommendations(IXLWorksheet worksheet, IEnumerable<ProductRecommendation> recommendations)
{
    worksheet.Cell(1, 1).Value = "Repo";
    worksheet.Cell(1, 2).Value = "Title";
    worksheet.Cell(1, 3).Value = "ms.author";
    worksheet.Cell(1, 4).Value = "Path";
    worksheet.Cell(1, 5).Value = "Uid";
    worksheet.Cell(1, 6).Value = "Recommendation";
    worksheet.Cell(1, 7).Value = "Label";

    var firstRow = worksheet.Row(1);
    firstRow.Style.Fill.BackgroundColor = XLColor.Black;
    firstRow.Style.Font.Bold = true;
    firstRow.Style.Font.FontSize = 16;
    firstRow.Style.Font.FontColor = XLColor.Aqua;
    worksheet.SheetView.FreezeRows(1);

    worksheet.Column(6).Style.Fill.BackgroundColor = XLColor.LightGray;

    int index = 2;
    foreach (var recommendation in recommendations)
    {
        worksheet.Row(index).Style.Font.FontSize = 12;
        worksheet.Cell(index, 1).Value = recommendation.Repo;
        worksheet.Cell(index, 2).Value = recommendation.Metadata.Title;
        worksheet.Cell(index, 3).Value = recommendation.Metadata.MsAuthor;
        worksheet.Cell(index, 4).Value = PathTable.FilePaths[recommendation.Metadata.Uid];
        worksheet.Cell(index, 5).Value = recommendation.Metadata.Uid;
        worksheet.Cell(index, 6).Value = recommendation.Slug;
        worksheet.Cell(index, 7).Value = recommendation.Label;
        index++;
    }

    foreach (var column in worksheet.ColumnsUsed())
    {
        column.AdjustToContents();
    }
}
