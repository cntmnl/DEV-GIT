using SharePoint.TaxonomyManager.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharePoint.TaxonomyManager
{
  class Program
  {

    internal sealed class ConsoleArgument
    {

      public string SiteUrl
      {
        get;
        set;
      }
      public string Path
      {
        get;
        set;
      }
      public string Operation
      {
        get;
        set;
      }
      public string ExportType
      {
        get;
        set;
      }
      public string ManagedService
      { get; set; }
      public bool CopyId 
      { get; set; }
    }


    static void Main(string[] args)
    {
      ConsoleArgument __p = new ConsoleArgument();

      if (args.Length < 4)
      {
        showHelp();
        return;
      }

      if (args[0].Equals("?"))
      {
        showHelp();
        return;

      }


      

      foreach (string item in args)
      {
        List<string> __split = new List<string>(2);
        int __index = item.IndexOf(':');
        if (__index > -1)
        {
          __split.Add(item.Substring(0, __index));
          __split.Add(item.Substring(__index + 1, item.Length - __index - 1));
        }
        string[] __arg = __split.ToArray();
        if (__arg.Length == 2)
        {
          switch (__arg[0])
          {
            case "-site":
              __p.SiteUrl = __arg[1];
              break;
            case "-path":
              __p.Path = __arg[1];
              break;
            case "-managedservice":
              __p.ManagedService = __arg[1];
              break;
            case "-operation":
              __p.Operation = __arg[1];
              break;
            case "-exporttype":
              __p.ExportType = __arg[1];
              break;
            case "-copyTermsID":
              __p.CopyId = bool.Parse(__arg[1]);
              break;
          }
        }
      }

      try
      {
        switch (__p.Operation)
        {
          case "import":
            XmlImport.Import(__p.Path, __p.ManagedService, __p.SiteUrl, __p.CopyId);
            break;
          case "export":
            export(__p);
            break;
          default:
            throw new Exception("Operation parameter not recognized, please check the help");
        }



      }
      catch (Exception err)
      {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(err);
        Console.ResetColor();
      }


      Console.ReadLine();
    }

    private static void export(ConsoleArgument argument)
    {
      switch (argument.ExportType.ToLower())
      {
        case "xls":
          ExcelExport.ExportFullTaxonomy(argument.Path, argument.ManagedService, argument.SiteUrl);
          break;
        case "xml":
          XmlExport.ExportFullTaxonomy(argument.Path, argument.ManagedService, argument.SiteUrl);
          break;
        case "all":
          ExcelExport.ExportFullTaxonomy(argument.Path, argument.ManagedService, argument.SiteUrl);
          XmlExport.ExportFullTaxonomy(argument.Path, argument.ManagedService, argument.SiteUrl);
          break;
        default:          
          throw new Exception("Export Type parameter not recognized, please check the help");
      }
    }
    
    private static void showHelp()
    {
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.WriteLine("Usage: .exe");
      Console.WriteLine("-site:[SiteUrl]");
      Console.WriteLine("-managedservice:[Name of the Managed Metadata Service]");
      Console.WriteLine("-path:[Workspace destination folder; in the IMPORT operation is the name of the XML file]");
      Console.WriteLine("-operation[export/import]");
      Console.WriteLine("-exporttype:[xls/xml/all] OPTIONAL");
      Console.WriteLine("-copyTermsID:[true/false] OPTIONAL");


      Console.ResetColor();
      Console.WriteLine("Press any key to continue...");
      Console.ReadLine();
    }
  }
}
