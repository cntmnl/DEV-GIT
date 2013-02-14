using CarlosAg.ExcelXmlWriter;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Text;

namespace SharePoint.TaxonomyManager.Api
{
  public class XmlExport
  {
    public static void ExportFullTaxonomy(string directoryPath, string managedServiceName, string siteUrl)
    {
      using (SPSite site = new SPSite(siteUrl))
      {
        TaxonomySession __taxonomySession = new TaxonomySession(site, true);
        TermStore __defaultSiteCollectionTermStore = __taxonomySession.TermStores[managedServiceName];

        XDocument __xDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"));
        XElement __rootElt = new XElement("root");



        foreach (Group group in __defaultSiteCollectionTermStore.Groups)
        {
          XmlExport.ExportTermGroup(__rootElt, group);
        }

        __xDoc.Add(__rootElt);

        __xDoc.Save(Path.Combine(directoryPath, "Taxonomy_" + System.DateTime.Today.ToString("yyyyMMdd") + ".xml"));
      }
    }

    public static void ExportTermGroup(string directoryPath, Group group)
    {
      XDocument __xDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"));
      XElement __rootElt = new XElement("root");
      XmlExport.ExportTermGroup(__rootElt, group);
    }

    internal static void ExportTermGroup(XElement root, Group group)
    {
      Console.ForegroundColor = ConsoleColor.Cyan;
      Console.WriteLine("Exporting XML Group [" + group.Name + "]");

      XElement _eltGroup = new XElement("TermGroup", string.Empty);
      _eltGroup.Add(new XAttribute("name", group.Name));
      _eltGroup.Add(new XAttribute("id", group.Id));
      root.Add(_eltGroup);


      foreach (TermSet termset in group.TermSets)
      {
        if (termset.Terms.Count > 0)
        {
          Console.ForegroundColor = ConsoleColor.Magenta;
          Console.WriteLine("Exporting XML TermSet [" + termset.Name + "]");

          XElement _eltTermSet = new XElement("TermSet", string.Empty);
          _eltTermSet.Add(new XAttribute("name", termset.Name));
          _eltTermSet.Add(new XAttribute("id", termset.Id));
          _eltGroup.Add(_eltTermSet);          
          XmlExport.RecursiveTermManagement(_eltTermSet, termset.Terms);
        }
        else
        {
          Console.ForegroundColor = ConsoleColor.Yellow;
          Console.WriteLine("No Terms Found for TermSet[" + termset.Name + "]");
        }
      }

      Console.ResetColor();

    }

    private static void RecursiveTermManagement(XElement elt, TermCollection collectionTerm)
    {
      foreach (Term term in collectionTerm)
      {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Exporting XML Term [" + term.Name + "]");

        XElement __childEl = new XElement("Term", string.Empty);
        __childEl.Add(new XAttribute("name", term.Name));
        __childEl.Add(new XAttribute("id", term.Id));
        __childEl.Add(new XAttribute("language", term.TermSet.Group.TermStore.Languages[0]));

        foreach (Label lbl in term.Labels)
        {
          XElement __labelEl = new XElement("Label", string.Empty);
          __labelEl.Add(new XAttribute("name", lbl.Value));
          __labelEl.Add(new XAttribute("language", lbl.Language));
          __labelEl.Add(new XAttribute("isdefault", lbl.IsDefaultForLanguage));
          __childEl.Add(__labelEl);
        }


        elt.Add(__childEl);

        if (term.TermsCount > 0)
          XmlExport.RecursiveTermManagement(__childEl, term.Terms);
      }
    }
  }
}
