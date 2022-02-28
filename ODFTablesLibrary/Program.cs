using HtmlAgilityPack;
using ODFTablesLib;
using System.Collections.Generic;
using System.IO;

public static class App
{

    static void Main()
    {
        var odf = new ODFTables(Path.GetFullPath(@"C:\Angstrem\results.htm"));

        //odf.Cells;
        //table = new List<Result>();
        //HtmlAgilityPack();
    }
    public struct Result
    {
        public string Date { get; set; }
        public string U { get; set; }
        public string F { get; set; }
        public string N { get; set; }
        public string Scheme { get; set; }
        public string Cx { get; set; }
        public string tgd { get; set; }
        public string Sko_cx { get; set; }
        public string Sco_tg { get; set; }
        public string R { get; set; }
        public string T { get; set; } //???
        public string CC { get; set; }
        public string DeltaTg { get; set; }
        public string Ka { get; set; }
        public string R1 { get; set; }
        public string R2 { get; set; }
        public string Rzo { get; set; }
        public string Rzx { get; set; }
    }
    static List<Result> table;
    public static void HtmlAgilityPack()
    {
        HtmlDocument htmlSnippet = new HtmlDocument();
        //using (var stream = File.OpenRead(@"C:\Angstrem\results.htm"))
            htmlSnippet.Load(@"C:\Angstrem\results.htm");

        GetInfoRows(htmlSnippet.DocumentNode);
    }
    static void GetInfoRows(HtmlNode doc)
    {
        if (doc.FirstChild == null) return;
        if (doc.FirstChild.Name == "td" && (doc.FirstChild.InnerText != "" && doc.FirstChild.InnerText != "Датавремя "))
        {            
            table.Add(new Result()
            {
                Date = doc.ChildNodes[0].InnerText,
                U = doc.ChildNodes[1].InnerText,
                F = doc.ChildNodes[2].InnerText,
                N = doc.ChildNodes[3].InnerText,
                Scheme = doc.ChildNodes[4].InnerText,
                Cx = doc.ChildNodes[5].InnerText,
                tgd = doc.ChildNodes[6].InnerText,
                Sko_cx = doc.ChildNodes[7].InnerText,
                Sco_tg = doc.ChildNodes[8].InnerText,
                R = doc.ChildNodes[9].InnerText,
                T = doc.ChildNodes[10].InnerText,
                CC = doc.ChildNodes[11].InnerText,
                DeltaTg = doc.ChildNodes[12].InnerText,
                Ka = doc.ChildNodes[13].InnerText,
                R1 = doc.ChildNodes[14].InnerText,
                R2 = doc.ChildNodes[15].InnerText,
                Rzo = doc.ChildNodes[16].InnerText,
                Rzx = doc.ChildNodes[17].InnerText
            });
        }
        else
            foreach (var child in doc.ChildNodes)
                GetInfoRows(child);
    }
}
