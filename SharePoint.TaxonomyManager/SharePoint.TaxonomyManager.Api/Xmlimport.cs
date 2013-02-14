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
  public class XmlImport
  {
    public static void Import(string configFile, string managedServiceName, string siteUrl, bool copyId)
    {
      using (SPSite site = new SPSite(siteUrl))
      {
        TaxonomySession __taxonomySession = new TaxonomySession(site, true);
        TermStore __defaultSiteCollectionTermStore = __taxonomySession.TermStores[managedServiceName];

        XDocument __xDoc = XDocument.Load(configFile);


        IEnumerable<XElement> __xGroups = __xDoc.Root.Elements("TermGroup");

        foreach (XElement xGroup in __xGroups)
        {
          if (xGroup.Descendants().Count() > 0)
          {

            

            Group __termGroup = (xGroup.Attribute("id") != null) ? (__defaultSiteCollectionTermStore.GetGroup(new Guid(xGroup.Attribute("id").Value))) : null ;

            if (__termGroup == null)
            {
              __termGroup = __defaultSiteCollectionTermStore.CreateGroup(xGroup.Attribute("name").Value);
            }

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Importing Group [" + __termGroup.Name + "]");

            IEnumerable<XElement> __xSets = xGroup.Elements("TermSet");

            foreach (XElement xSet in __xSets)
            {
              TermSet __termSet;
              if (copyId)
                __termSet = __termGroup.CreateTermSet(xSet.Attribute("name").Value, new Guid(xSet.Attribute("id").Value));
              else
                __termSet = __termGroup.CreateTermSet(xSet.Attribute("name").Value);

              Console.ForegroundColor = ConsoleColor.Magenta;
              Console.WriteLine("Importing TermSet [" + __termSet.Name + "]");

              IEnumerable<XElement> __xTerms = xSet.Elements("Term");
              foreach (XElement xTerm in __xTerms)
              {
                Term __term;

                if (copyId)
                  __term = __termSet.CreateTerm(xTerm.Attribute("name").Value, Convert.ToInt32(xTerm.Attribute("language").Value), new Guid(xTerm.Attribute("id").Value));
                else
                {
                  __term = __termSet.CreateTerm(xTerm.Attribute("name").Value, Convert.ToInt32(xTerm.Attribute("language").Value));
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Importing TermSet [" + __term.Name + "]");

                foreach (XElement label in xTerm.Elements("Label"))
                {
                  if (!label.Attribute("name").Value.Equals(__term.Name))
                    __term.CreateLabel(label.Attribute("name").Value, Convert.ToInt32(label.Attribute("language").Value), Convert.ToBoolean(label.Attribute("isdefault").Value));

                }

                if (xTerm.Elements("Term").Count() > 0)
                  RecursiveTermCreation(xTerm.Elements("Term"), __term, copyId, 1033);
              }
            }

          }
        }

        __defaultSiteCollectionTermStore.CommitAll();

      }
    }

    private static void RecursiveTermCreation(IEnumerable<XElement> descendants, Term term, bool copyId, int LCID)
    {
      foreach (XElement xTerm in descendants)
      {
        Term __childTerm;
        if (copyId)
        {
          __childTerm = term.CreateTerm(xTerm.Attribute("name").Value, LCID, new Guid(xTerm.Attribute("id").Value));
        }
        else
        {
          __childTerm = term.CreateTerm(xTerm.Attribute("name").Value, LCID);
        }

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Importing TermSet [" + __childTerm.Name + "]");

        foreach (XElement label in xTerm.Elements("Label"))
        {
          if ((!label.Attribute("name").Value.Equals(__childTerm.Name)))
          __childTerm.CreateLabel(label.Attribute("name").Value, Convert.ToInt32(label.Attribute("language").Value), Convert.ToBoolean(label.Attribute("isdefault").Value));
        }


        if (xTerm.Elements("Term").Count() > 0)
          RecursiveTermCreation(xTerm.Elements("Term"), __childTerm, copyId, LCID);

      }
    }
  }
}
