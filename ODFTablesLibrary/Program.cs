using ODFTablesLib;
using System.IO;

public static class App
{

    static void Main()
    {
        new ODFTables(Path.GetFullPath(@"NewTemplate.odt")).Save(@"C:\1.odt");


    }
}
