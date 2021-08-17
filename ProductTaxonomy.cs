// https://github.com/MicrosoftDocs/docs-help-pr/blob/master/help-content/contribute/metadata-taxonomies.md
// To obtain the table data locally, `curl https://taxonomyservice.azurefd.net/taxonomies/product > product-taxonomy.json`
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

// Example:
/*{
    "uid":"https://authoring-docs-microsoft.poolparty.biz/devrel/006ab567-b18c-4cf1-9a25-c24daa46ede1",
    "level":2,
    "label":"App Service",
    "slug":"azure-app-service",
    "parentSlug":"azure",
    "createdAt":"2020-09-15T23:28:54.863Z","updatedAt":"2020-09-15T23:28:54.863Z",
    "updatedBy":"https://prod-docs-microsoft.poolparty.biz/user/swc_api_sync"
}*/
public class ProductTaxonomy
{
    [JsonPropertyName("level")]
    public int Level { get; set; }

    [JsonPropertyName("label")]
    public string Label { get; set; }

    [JsonPropertyName("slug")]
    public string Slug { get; set; }

    [JsonPropertyName("parentSlug")]
    public string ParentSlug { get; set; }
}

static class ProductTaxonomyTable
{
    public static List<ProductTaxonomy> Entries = new();

    public static void Dump()
    {
        var level1Entries = Entries.Where(pt => pt.Level == 1).ToList();

        level1Entries.Sort((pt1, pt2) => pt1.Slug.CompareTo(pt2.Slug));
        foreach (var level1Entry in level1Entries)
        {
            Console.WriteLine($"\"{level1Entry.Slug}\": \"{level1Entry.Label}\"");
            var level2Entries = Entries.Where(pt => pt.Level == 2 && pt.ParentSlug == level1Entry.Slug).ToList();
            level2Entries.Sort((pt1, pt2) => pt1.Slug.CompareTo(pt2.Slug));
            foreach (var level2Entry in level2Entries)
            {
                Console.WriteLine($"\t\"{level2Entry.Slug}\": \"{level2Entry.Label}\"");
            }
        }
    }

    public static void Load(string jsonFile)
    {
        Entries = JsonSerializer.Deserialize<List<ProductTaxonomy>>(File.ReadAllText(jsonFile));
    }

    // public static IEnumerable<string> MapProductsToCSAs(IEnumerable<string> productNames)
    // {
    //     List<string> list = new();
    //     foreach (var productName in productNames)
    //     {
    //         Product product;
    //         if (Mapping.TryGetValue(productName, out product))
    //         {
    //             list.Add(product.CSA);
    //         }
    //     }
    //     return list;
    // }
}