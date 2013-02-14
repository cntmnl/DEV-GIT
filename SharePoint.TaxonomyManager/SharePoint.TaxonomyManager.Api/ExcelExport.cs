using CarlosAg.ExcelXmlWriter;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SharePoint.TaxonomyManager.Api
{
  public class ExcelExport
  {
    public static void ExportFullTaxonomy(string directoryPath, string managedServiceName, string siteUrl)
    {
      using (SPSite site = new SPSite(siteUrl))
      {

        TaxonomySession __taxonomySession = new TaxonomySession(site, true);
        TermStore __defaultSiteCollectionTermStore = __taxonomySession.TermStores[managedServiceName];

        foreach (Group group in __defaultSiteCollectionTermStore.Groups)
        {
          ExcelExport.ExportTermGroup(directoryPath, group);

        }
      }
    }

    public static void ExportTermGroup(string directoryPath, Guid groupId, string siteUrl)
    {

      using (SPSite site = new SPSite(siteUrl))
      {

        TaxonomySession __taxonomySession = new TaxonomySession(site, true);
        TermStore __defaultSiteCollectionTermStore = __taxonomySession.DefaultSiteCollectionTermStore;

        Group group = __defaultSiteCollectionTermStore.GetGroup(groupId);

        ExcelExport.ExportTermGroup(directoryPath, group);


      }
    }

    internal static void ExportTermGroup(string directoryPath, Group group)
    {
      Console.ForegroundColor = ConsoleColor.Cyan;
      Console.WriteLine("Exporting XLS Group [" + group.Name + "]");

      bool _termsFound = false;


      Workbook book = new Workbook();
      foreach (TermSet termset in group.TermSets)
      {
        if (termset.Terms.Count > 0)
        {
          Console.ForegroundColor = ConsoleColor.Magenta;
          Console.WriteLine("Exporting XLS TermSet [" + termset.Name + "]");
          Worksheet sheet = book.Worksheets.Add(termset.Name);
          foreach (Term term in termset.Terms)
          {
            _termsFound = true;
            WorksheetRow row = sheet.Table.Rows.Add();

            foreach (Label lbl in term.Labels)
            {
              row.Cells.Add(lbl.Value);              
              row.Cells.Add(new WorksheetCell(lbl.Language.ToString(), DataType.Number));
            }
            
            
            if (term.TermsCount > 0)
            {
              ExcelExport.RecursiveRowManagement(1, sheet, term.Terms);
            }
          }
        }
        else
        {
          Console.ForegroundColor = ConsoleColor.Yellow;
          Console.WriteLine("No Terms Found for TermSet[" + termset.Name + "]");
        }
      }

      if (_termsFound)
        book.Save(Path.Combine(directoryPath, group.Name + ".xls"));

      Console.ResetColor();

    }

    private static void RecursiveRowManagement(int level, Worksheet sheet, TermCollection collectionTerm)
    {
      foreach (Term term in collectionTerm)
      {
        WorksheetRow row = sheet.Table.Rows.Add();

        for (int i = 0; i < level; i++)
        {
          row.Cells.Add("");
        }

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Exporting XLS Term [" + term.Name + "]");

        foreach (Label lbl in term.Labels)
        {
          row.Cells.Add(lbl.Value);
          row.Cells.Add(new WorksheetCell(lbl.Language.ToString(), DataType.Number));
        }

        if (term.TermsCount > 0)
          ExcelExport.RecursiveRowManagement(level + 1, sheet, term.Terms);

      }



    }
  }
}
