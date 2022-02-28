using System;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;

namespace ODFTablesLib
{
    public class ODFTables
    {
        private string temp;
        private XmlDocument doc;
        public CellRange Cells;
        public ODFTables(string path) => Load(path);
        /// <summary>
        /// Loading new ODF file 
        /// </summary>
        /// <param name="path">Path</param>
        public void Load(string path)
        {
            if (!Directory.Exists("Temp")) Directory.CreateDirectory("Temp");
            this.temp = path;
            using (ZipArchive zipArchive = ZipFile.OpenRead(path))
            using (var stream = zipArchive.Entries.FirstOrDefault(x => x.Name == "content.xml")?.Open())
            using (StreamReader sr = new StreamReader(stream, Encoding.UTF8))
            {
                string text = sr.ReadToEnd();
                doc = new XmlDocument();
                doc.LoadXml(text);
            }
            Cells = new CellRange(doc, temp);
        }
        /// <summary>
        /// Save current file
        /// </summary>
        /// <param name="path">Path</param>
        public void Save(string path)
        {
            File.Copy(temp, path, overwrite: true);
            var pathTemp = Path.GetFullPath("Temp/content.xml");
            using (ZipArchive zipArchive = ZipFile.OpenRead(path))
                zipArchive.Entries.FirstOrDefault(x => x.Name == "content.xml")?.
                    ExtractToFile(pathTemp, true);
            using (StreamWriter streamWriter = new StreamWriter(pathTemp)) streamWriter.Write(doc.OuterXml);
            using (ZipArchive zipArchive = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                zipArchive.Entries.FirstOrDefault((ZipArchiveEntry x) => x.Name == "content.xml")?.Delete();
                ZipFileExtensions.CreateEntryFromFile(zipArchive, pathTemp, "content.xml");
            }

            File.Delete(pathTemp);
        }
        public void SaveAsPDF()
        {
            var type = "odt";
            if (temp.EndsWith(".ods")) type = "ods";
            string filePath = Path.GetFullPath($@"Temp\print.{type}");
            if (!Directory.Exists("Temp"))
                Directory.CreateDirectory("Temp");
            Save(filePath);

            bool PDFPrinterFound = false;
            foreach (string strPrinter in PrinterSettings.InstalledPrinters)
                if (strPrinter.ToLower().Contains("pdf"))
                {
                    SetDefaultPrinter(strPrinter);
                    PDFPrinterFound = true;
                    break;
                }
            if (!PDFPrinterFound) throw new Exception("Вывод в PDF не доступен для данного компьютера");


            Process printProcess = new Process()
            {
                StartInfo = new ProcessStartInfo()
                {
                    Verb = "print",
                    CreateNoWindow = true,
                    FileName = filePath,
                    WindowStyle = ProcessWindowStyle.Hidden
                },
            };
            printProcess.Start();
            printProcess.WaitForExit(1000);
            try
            {
                printProcess.Kill();
                throw new Exception("Ошибка вывода в PDF");
            }
            catch { }
            new FileInfo(filePath).Delete();
        }
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetDefaultPrinter(string Printer);
    }

}
