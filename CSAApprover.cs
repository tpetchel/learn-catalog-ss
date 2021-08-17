
//https://github.com/MicrosoftDocs/docs-help-pr/blob/master/help-content/contribute/faq-service-content-owners.md
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Toolkit.Parsers.Markdown;
using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using Microsoft.Toolkit.Parsers.Markdown.Inlines;

class CSAApprover
{
    public string CSA;
    public string ApproverName;
    public string ApproverAlias;
}

static class CSAApproversTable
{
    // Maps CSA -> Approver
    public static Dictionary<string, CSAApprover> Mapping = new();

    public static void Dump()
    {
        foreach (var kvp in Mapping)
        {
            var v = kvp.Value;
            Console.WriteLine($"\"{v.CSA}\",\"{v.ApproverName}\",\"{v.ApproverAlias}\"");
        }
    }

    public static void Load(string markdownFile)
    {
        var approverMapping = Mapping;

        var doc = new MarkdownDocument();
        doc.Parse(File.ReadAllText(markdownFile));
        var tableBlock = doc.Blocks.Where(block => block.Type == MarkdownBlockType.Table).Single() as TableBlock;

        foreach (var row in tableBlock.Rows)
        {
            string c0 = Concat(row.Cells[0].Inlines).Trim();
            string c1 = Concat(row.Cells[1].Inlines).Trim();
            string c2 = Concat(row.Cells[2].Inlines).Trim();

            // Skip rows labled "**A**", "**B**", "**C**", etc. used for categorization.
            if (c0.Length == "**A**".Length && c1.Length == 0 && c2.Length == 0)
            {
                continue;
            }

            approverMapping[c0] = new CSAApprover() {
                CSA = c0,
                ApproverName = c1,
                ApproverAlias = c2
            };
        }
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

    private static string Concat(IList<MarkdownInline> inlines)
    {
        System.Text.StringBuilder sb = new();
        foreach (var inline in inlines)
        {
            sb.Append(inline.ToString());
        }
        return sb.ToString();
    }
}