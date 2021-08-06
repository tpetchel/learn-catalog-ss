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
    //"/Users/thpetche/Dev/MicrosoftDocs/learn-certs-pr"
};

using (var workbook = new XLWorkbook())
{
    var allMetadata = new List<(string, IEnumerable<ModuleMetadata>)>();
    foreach (var repo in repos)
    {
        string repoName = repo.Split('/').Last();
        var worksheet = workbook.Worksheets.Add(repoName);
        var metadataList = CollectModuleMetadata(repo);
        allMetadata.Add((repoName, metadataList));
        WriteWorksheet(worksheet, metadataList);
    }

    var summaryWorksheet = workbook.Worksheets.Add("Summary");
    WriteSummary(summaryWorksheet, allMetadata);

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
        metadataList.Add(metadata);
    }
    return metadataList;
}

static void WriteWorksheet(IXLWorksheet worksheet, IEnumerable<ModuleMetadata> metadataList)
{
    worksheet.Cell(1, 1).Value = "Title";
    worksheet.Cell(1, 2).Value = "Author";
    worksheet.Cell(1, 3).Value = "Product";

    var firstRow = worksheet.Row(1);
    firstRow.Style.Fill.BackgroundColor = XLColor.Black;
    firstRow.Style.Font.Bold = true;
    firstRow.Style.Font.FontSize = 16;
    firstRow.Style.Font.FontColor = XLColor.Aqua;
    worksheet.SheetView.FreezeRows(1);

    int index = 2;
    foreach (var metadata in metadataList)
    {
        foreach (var product in metadata.Products)
        {
            worksheet.Row(index).Style.Font.FontSize = 12;
            worksheet.Cell(index, 1).Value = metadata.Title;
            worksheet.Cell(index, 2).Value = metadata.Author;
            worksheet.Cell(index, 3).Value = product;
            index++;
        }
    }

    foreach (var column in worksheet.ColumnsUsed())
    {
        column.AdjustToContents();
    }
}

static void WriteSummary(IXLWorksheet worksheet, IEnumerable<(string, IEnumerable<ModuleMetadata>)> allMetadata)
{
    worksheet.Cell(1, 1).Value = "Repo";
    worksheet.Cell(1, 2).Value = "Modules";

    var firstRow = worksheet.Row(1);
    firstRow.Style.Fill.BackgroundColor = XLColor.Black;
    firstRow.Style.Font.Bold = true;
    firstRow.Style.Font.FontSize = 16;
    firstRow.Style.Font.FontColor = XLColor.Aqua;
    worksheet.SheetView.FreezeRows(1);

    int index = 2;
    foreach (var (repo, metadataList) in allMetadata)
    {
        worksheet.Row(index).Style.Font.FontSize = 12;
        worksheet.Cell(index, 1).Value = repo;
        worksheet.Cell(index, 2).Value = metadataList.Count();
        index++;
    }

    foreach (var column in worksheet.ColumnsUsed())
    {
        column.AdjustToContents();
    }
}