using System.Xml;
using System.Text;
using System.Collections;
using System.Text.RegularExpressions;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO.Compression;

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
            Cells = new CellRange(doc);
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
                }
            };
            printProcess.Start();
            printProcess.WaitForExit();
            try
            {
                printProcess.Kill();
            }
            catch { }
            new FileInfo(filePath).Delete();
        }
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetDefaultPrinter(string Printer);
    }

    public class CellRange : IEnumerable
    {
        private int _totalRows = -1;
        private int _totalColumns = 0;
        private XmlDocument doc;
        XmlNode EmptyNode;
        string textNameSpaceURI = string.Empty;
        #region Public fields
        public List<Cell> Cells;
        public int TotalRows => _totalRows;
        public int TotalColumns => _totalColumns;
        public int Count => Cells.Count;
        public int Width => Cells[Cells.Count - 1].Column - (Cells[0].Column - 1);
        public int Height => Cells[Cells.Count - 1].Row - (Cells[0].Row - 1);
        public int FirstRowIndex => Cells[0].Row;
        public int FirstColumnIndex => Cells[0].Column;
        public int LastRowIndex => Cells[Cells.Count - 1].Row;
        public int LastColumnIndex => Cells[Cells.Count - 1].Column;
        /// <summary>
        /// Selecting a cell by index in the list
        /// the list is filled in line by line
        /// </summary>
        /// <param name="index">Index</param>
        /// <returns>Cell</returns>
        public Cell this[int index] => this.Cells[index];
        /// <summary>
        /// Selecting a cell by text inside
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Cell</returns>
        public Cell this[string value] => this.Cells.FirstOrDefault(c => c.Name == value);
        /// <summary>
        /// Selecting a cell by position 
        /// </summary>
        /// <param name="relativeRow">Row</param>
        /// <param name="relativeColumn">Column</param>
        /// <returns>Cell</returns>
        public Cell this[int relativeRow, int relativeColumn] => this.Cells.FirstOrDefault(c => c.Row == relativeRow && c.Column == relativeColumn)
                                                              ?? this.Cells.FirstOrDefault(c => c.Node == EmptyNode);

        /// <summary>
        /// Linq
        /// </summary>
        /// <param name="f">Predicate</param>
        /// <returns>Cell</returns>
        public Cell FirstOrDefault(Func<Cell, bool> f) => Cells.FirstOrDefault(f);
        public IEnumerator GetEnumerator() => Cells.GetEnumerator();
        #endregion
        private CellRange() => Cells = new List<Cell>();
        public CellRange(XmlDocument doc, bool inner = false)
        {
            this.doc = doc;
            Cells = new List<Cell>();

            textNameSpaceURI = FindTextURI(doc);
            EmptyNode = doc.CreateElement("table:p", $"{textNameSpaceURI}");
            if (!inner)
            {
                GetCell(GetBody(doc.DocumentElement));
                _totalRows++; // абсолютное значение кол-ва строк, ++ т.к. начинается с 0
            }
        }
        private string FindTextURI(XmlNode doc)
        {
            var item = doc.ChildNodes[1].OuterXml;
            int start = item.IndexOf("xmlns:text=");
            int end = item.IndexOf($" xmlns:", start + 15);
            return item.Substring(start + 12, end - (start + 13));
        }
        private XmlNode GetBody(XmlNode node)
        {
            foreach (XmlNode child in node.ChildNodes)
                if (child.Name == "office:body") return child;
            return null;
        }
        private void GetCell(XmlNode node, int column = -1, int rowSpan = 0, int columnSpan = 0)
        {
            foreach (XmlNode child in node.ChildNodes)
            {
                switch (child.Name)
                {
                    case "text:p":
                        if (child.NextSibling?.Name == "table:table")
                        {
                            _totalRows++;
                            AddNode(child, 0);
                        }
                        break;
                    case "table:table-cell":
                        rowSpan = int.Parse(child.Attributes["table:number-rows-spanned"]?.Value ?? "0");
                        columnSpan = int.Parse(child.Attributes["table:number-columns-spanned"]?.Value ?? "0");
                        column++;
                        AddCell(column, rowSpan, columnSpan, child);
                        break;
                    case "table:covered-table-cell":
                        int.TryParse((child.Attributes["table:number-columns-repeated"]?.Value ?? "1"), out int cc);
                        for (int i = 0; i < cc; i++)
                            AddNode(child, ++column);
                        break;
                    case "table:table-row":
                        _totalRows++;
                        break;
                    case "table:table":
                        int ColumnsCount = 0;
                        foreach (XmlNode item in child.ChildNodes)
                        {
                            if (item.Name == "table:table-column")
                            {
                                if (int.TryParse((item.Attributes["table:number-columns-repeated"]?.Value ?? ""), out int c)) ColumnsCount += c;
                                else ColumnsCount++;
                                if (_totalColumns < ColumnsCount) _totalColumns = ColumnsCount;
                            }
                        }
                        break;
                }
                GetCell(child, column, rowSpan, columnSpan);
            }
        }

        private void AddCell(int column, int rowSpan, int columnSpan, XmlNode child)
        {
            if (child.FirstChild?.Name.StartsWith("text") ?? false && column != -1)
            {
                AddNode(child.FirstChild, column, rowSpan, columnSpan);
                if (rowSpan > 0 || columnSpan > 0)
                {
                    var oldCell = Cells[Cells.Count - 1];
                    var m = oldCell.MergedRange = new CellRange(doc, true);
                    int i = 0, j = 0;
                    do
                    {
                        do
                        {
                            m.Cells.Add(new Cell()
                            {
                                Node = EmptyNode,
                                Value = "",
                                Row = _totalRows + j,
                                RowSpan = 0,
                                Column = column + i,
                                ColumnSpan = 0
                            });
                            j++;
                        } while (j < rowSpan);
                        i++; j = 0;
                    } while (i < columnSpan);
                    m.Cells[0].Node = child.FirstChild;
                }
            }
            else
            {

                var e = doc.CreateElement("text", "p", $"{textNameSpaceURI}");
                e.SetAttribute("style-name", $"{textNameSpaceURI}", $"P1"); // P1 - первый стиль для text:p

                child.PrependChild(e);
                AddNode(child.PrependChild(e), column);
            }

        }


        private void AddNode(XmlNode node, int column, int rowSpan = 0, int columnSpan = 0)
        {
            this.Cells.Add(new Cell()
            {
                Node = node,
                Value = node.InnerText,
                Row = _totalRows,
                RowSpan = rowSpan,
                Column = column,
                ColumnSpan = columnSpan
            });
        }

        /// <summary>
        /// Search for rexgex
        /// </summary>
        /// <param name="regex">regex</param>
        /// <param name="row">row</param>
        /// <param name="column">column</param>
        /// <returns>Found?</returns>
        public bool FindText(Regex regex, out int row, out int column)
        {
            row = 0; column = 0;
            foreach (var node in Cells)
                if (regex.IsMatch(node.Value))
                {
                    row = node.Row;
                    column = node.Column;
                    return true;
                }
            return false;
        }
        public CellRange GetSubrangeRelative(int relativeRow, int relativeColumn, int width, int height)
        {
            CellRange range = new CellRange(doc, true);
            int i = Cells.IndexOf(Cells.FirstOrDefault(x => x.Row == relativeRow));
            int last = Cells.IndexOf(Cells.LastOrDefault(x => x.Row == relativeRow + height - 1));
            for (; i <= last; i++)
                range.Cells.Add(Cells[i]);

            range._totalRows = height;
            range._totalColumns = width;
            range.EmptyNode = EmptyNode;
            return range;
        }
        public CellRange GetSubrangeAbsolute(int FirstRowIndex, int FirstColumnIndex, int LastRowIndex, int LastColumnIndex)
        {
            var range = new CellRange(doc, true);
            for (int i = FirstRowIndex; i <= LastRowIndex; i++)
                for (int j = FirstColumnIndex; j <= LastColumnIndex; j++)
                    range.Cells.Add(this[i, j]);

            range._totalColumns = LastColumnIndex - FirstColumnIndex + 1;
            range._totalRows = LastRowIndex - FirstRowIndex + 1;
            range.EmptyNode = EmptyNode;
            return range;
        }
    }
    public class Cell
    {
        /// <summary>
        /// Do not ask why name is [row,columt] instead of [column,row]
        /// I dont have right answer
        /// </summary>
        public Cell() { }
        public CellRange MergedRange { get; set; }
        internal XmlNode Node { get; set; }
        public int RowSpan { get; set; }
        public int ColumnSpan { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public string Name { get => $"{GetColumnName(this.Column)}{this.Row}"; }
        public string Value
        {
            get => Node.InnerText;
            set => Node.InnerText = value;
        }
        public override bool Equals(object obj) => obj is Cell && (obj as Cell).Name == this.Name;

        public override int GetHashCode()
        {
            int hashCode = 1232569091;
            hashCode = hashCode * -1521134295 + EqualityComparer<XmlNode>.Default.GetHashCode(Node);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Value);
            hashCode = hashCode * -1521134295 + RowSpan.GetHashCode();
            hashCode = hashCode * -1521134295 + ColumnSpan.GetHashCode();
            hashCode = hashCode * -1521134295 + Row.GetHashCode();
            hashCode = hashCode * -1521134295 + Column.GetHashCode();
            return hashCode;
        }//???

        private string GetColumnName(int number)
        {
            string data = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            int baseC = data.Length;
            string res = string.Empty;
            int index;
            if (number == 0) return "A";
            while (number >= 1)
            {
                if (res != string.Empty)
                    number = number - 1;
                index = number % baseC;
                number /= baseC;
                res = data[index] + res;
                if (number < 0)
                    res = data[0] + res;
                index = 0;
            }
            return res;
        }
    }

}
