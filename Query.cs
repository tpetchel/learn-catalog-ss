
using System.Linq;
using System.Collections.Generic;

static class Query
{    public static IEnumerable<string> GetApprovers(IEnumerable<string> productSlugs)
    {
        List<string> approvers = new();
        foreach (var productSlug in productSlugs)
        {
            // Get the entry from the taxonomy table so that we can access the product's parent slug.
            var taxonomy = ProductTaxonomyTable.Entries.Where(pt => pt.Slug == productSlug).Single();

            //System.Console.WriteLine($"Slug: {productSlug}");

            // Get the product CSA owner and the CSA owner of the parent product.
            var productCSAOwner = ProductCSAOwnerTable.Mapping.Values.Where(csaOwner => csaOwner.Slug == productSlug).SingleOrDefault();
            var productCSAParentOwner = ProductCSAOwnerTable.Mapping.Values.Where(csaOwner => csaOwner.Slug == taxonomy.ParentSlug).SingleOrDefault();
            if (productCSAOwner != null)
            {
                approvers.Add(productCSAOwner.M2);
                //System.Console.WriteLine($"\tCSA1: {productCSAOwner.CSA}");

                var approver = CSAApproversTable.Mapping.Values.Where(a => a.CSA == productCSAOwner.CSA).SingleOrDefault();
                if (approver != null)
                {
                    approvers.Add(approver.ApproverName);
                }
            }
            if (productCSAParentOwner != null)
            {
                approvers.Add(productCSAParentOwner.M2);
                //System.Console.WriteLine($"\tCSA2: {productCSAParentOwner.CSA}");

                var approver = CSAApproversTable.Mapping.Values.Where(a => a.CSA == productCSAParentOwner.CSA).SingleOrDefault();
                if (approver != null)
                {
                    approvers.Add(approver.ApproverName);
                }
            }
        }
        approvers.RemoveAll(s => string.IsNullOrWhiteSpace(s));
        approvers.Sort();
        return approvers.Distinct();
    }

    public static IEnumerable<ProductRecommendation> RecommendProducts(IEnumerable<(string, IEnumerable<ModuleMetadata>)> allMetadata)
    {
        List<ProductRecommendation> recommendations = new();
        foreach (var (repo, metadataList) in allMetadata)
        {
            foreach (var metadata in metadataList)
            {
                foreach (var entry in ProductTaxonomyTable.Entries)
                {
                    if (metadata.Title.Contains($" {entry.Label} ") ||
                        metadata.Title.EndsWith($" {entry.Label}"))
                    {
                        if (! metadata.Products.Contains(entry.Slug))
                        {
                            recommendations.Add(new ProductRecommendation() {
                                Repo = repo,
                                Metadata = metadata,
                                Slug = entry.Slug,
                                Label = entry.Label
                            });
                            //System.Console.WriteLine($"{metadata.Title} => {entry.Slug} ({entry.Label})");
                        }
                    }
                }
            }
        }
        recommendations.Sort((r1, r2) => {
            int cmp = r1.Repo.CompareTo(r2.Repo);
            if (cmp == 0)
            {
                cmp = r1.Metadata.Title.CompareTo(r2.Metadata.Title);
            }
            return cmp;
        });
        return recommendations;
    }
}