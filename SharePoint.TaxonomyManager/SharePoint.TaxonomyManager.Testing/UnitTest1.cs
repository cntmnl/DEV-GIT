using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.SharePoint;
using SharePoint.TaxonomyManager.Api;

namespace SharePoint.TaxonomyManager.Testing
{
  [TestClass]
  public class UnitTest1
  {
    string _siteUrl = "http://employeefr.nespresso.local";

    [TestInitialize]
    public void Initialize()
    {
     
    }
    

    [TestMethod]
    public void TestMethod1()
    {
      ExcelExport.ExportFullTaxonomy(@"C:\TEMP", "Managed Metadata Service", _siteUrl);


    }
  }
}
