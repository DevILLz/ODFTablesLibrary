using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace ODFTablesLib
{
    public class CellRange : IEnumerable
    {
        private int _totalRows = -1;
        private int _totalColumns = 0;
        private XmlDocument doc;
        private XmlNode EmptyNode;
        private string TextNS = string.Empty;
        private string TableNS = string.Empty;
        private string OfficeNS = string.Empty;
        private string CalcextNS = string.Empty;
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
                                                              ?? AddNewCell(relativeColumn, relativeRow);

        /// <summary>
        /// Linq
        /// </summary>
        /// <param name="f">Predicate</param>
        /// <returns>Cell</returns>
        public Cell FirstOrDefault(Func<Cell, bool> f) => Cells.FirstOrDefault(f);
        public IEnumerator GetEnumerator() => Cells.GetEnumerator();
        #endregion
        private CellRange() => Cells = new List<Cell>();
        private string temp;
        public CellRange(XmlDocument doc, string temp = "", bool inner = false, object NS = null)
        {
            this.temp = temp;
            this.doc = doc;
            
            Cells = new List<Cell>();
            if (NS == null)
                SearchForNameSpaces(XDocument.Parse(doc.ChildNodes[1].OuterXml));
            else
            {
                var nss = (NS as (string, string, string, string)?).Value;
                TableNS = nss.Item1;
                OfficeNS = nss.Item2;
                CalcextNS = nss.Item3;
                TextNS = nss.Item4;
            }
            EmptyNode = doc.CreateElement("table:p", $"{TextNS}");
            if (!inner)
            {
                GetCell(GetBody(doc.DocumentElement));
                _totalRows++; // абсолютное значение кол-ва строк, ++ т.к. начинается с 0
            }
        }
        private void SearchForNameSpaces(XDocument doc)
        {
            TableNS = doc.Root.GetNamespaceOfPrefix("table").ToString();
            OfficeNS = doc.Root.GetNamespaceOfPrefix("office").ToString();
            CalcextNS = doc.Root.GetNamespaceOfPrefix("calcext").ToString();
            TextNS = doc.Root.GetNamespaceOfPrefix("text").ToString();
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
                        if (child.NextSibling?.Name == "table:table" )
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
                                if (int.TryParse(item.Attributes["table:number-columns-repeated"]?.Value ?? "", out int c)) ColumnsCount += c;
                                else ColumnsCount++;
                                if (_totalColumns < ColumnsCount) _totalColumns = ColumnsCount;
                            }
                        }
                        break;
                }
                GetCell(child, column, rowSpan, columnSpan);
            }
        }
        internal XmlNode AddTextBlock(XmlNode node) => node.PrependChild(doc.CreateElement("text", "p", $"{TextNS}"));
        internal XmlNode ChangeCellType(XmlNode node, string type)
        {
            if (string.IsNullOrEmpty(type)) return node;
            if (node?.Name.StartsWith("text") ?? false)
                node = node.ParentNode;
            var n = node as XmlElement;
            if (node?.Name.StartsWith("table:table-cell") ?? false)
            {
                var o = n.Attributes.GetNamedItem("office:value-type");
                var c = n.Attributes.GetNamedItem("calcext:value-type");
                if (o != null)
                    o.Value = type;
                else
                {
                    var off = doc.CreateAttribute("office", "value-type", $"{OfficeNS}");
                    off.Value = type;
                    n.Attributes.Append(off);
                }
                if (c != null)
                    c.Value = type;
                else
                {
                    var calc = doc.CreateAttribute("calcext", "value-type", $"{CalcextNS}");
                    calc.Value = type;
                    n.Attributes.Append(calc);
                }
            }
            if (type == "float")
            {
                var v = n.Attributes.GetNamedItem("office:value");
                if (v != null)
                    v.Value = n.InnerText;
                else
                {
                    var calc = doc.CreateAttribute("office", "value", $"{OfficeNS}");
                    calc.Value = n.InnerText;
                    n.Attributes.Append(calc);
                }
            }
            return node;
        }

        private void AddCell(int column, int rowSpan, int columnSpan, XmlNode child)
        {
            if (child.FirstChild?.Name.StartsWith("text") ?? false && column != -1)
            {
                AddNode(child.FirstChild, column, rowSpan, columnSpan);
                if (rowSpan > 0 || columnSpan > 0)
                {
                    var oldCell = Cells[Cells.Count - 1];
                    var m = oldCell.MergedRange = new CellRange(doc, temp, true, (TableNS, TableNS, CalcextNS, TextNS));
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
                if (temp.EndsWith(".ods"))
                    AddNode(child, column);
                else
                    AddNewNode(child, column);
            }
        }
        private Cell AddNewCell(int column, int row)
        {
            ProtectdTask(() => { AddNewNode(this.Cells.LastOrDefault(c => c.Row <= row).Node, column, row); });
            
            return Cells[Cells.Count - 1];
        }
        private XmlNode AddNewNode(XmlNode node, int column, int row = -1)
        {
            var t = doc.CreateElement("text", "p", $"{TextNS}");
            //t.SetAttribute("style-name", $"{TextNS}", $"P1"); // P1 - первый стиль для text:p
            if (row == -1) AddNode(node.PrependChild(t), column);
            else
            {
                XmlNode newRow = node;
                while (newRow?.Name != "table:table") // поиск таблицы
                    newRow = newRow.ParentNode;

                if (row > (Cells[Cells.Count - 1].Row + 1))
                    for (int i = 0; i < (row - Cells[Cells.Count - 1].Row) - 1; i++) // добавление пустых строк
                        newRow.AppendChild(doc.CreateElement("table", "table-row", TableNS));

                if (row > Cells[Cells.Count - 1].Row)
                    newRow = newRow.AppendChild(doc.CreateElement("table", "table-row", TableNS)); // добавление нужной строки
                else
                {
                    int count = -1;
                    foreach (XmlNode child in newRow.ChildNodes) // ищем нужную строку
                    {
                        if (child.Name == "table:table-row") count++;
                        if (count == row)
                        {
                            newRow = child;
                            break;
                        }
                    }
                }
                    
                int start = this.Cells.LastOrDefault(x => x.Row == row)?.Column ?? 0;
                for (int i = start; i < column - 1; i++)
                {
                    var cell = doc.CreateElement("table", "table-cell", $"{TableNS}"); // создать ячейку
                    var text = doc.CreateElement("text", "p", $"{TextNS}"); // создать текстовое поле
                    cell.AppendChild(text);
                    newRow.AppendChild(cell); 
                    AddNode(text, i, 0, 0, row);
                }
                var c = doc.CreateElement("table", "table-cell", $"{TableNS}");
                c.AppendChild(t);
                newRow.AppendChild(c);
                AddNode(t, column, 0, 0, row);
            }
            return t;
        }

        private void AddNode(XmlNode node, int column, int rowSpan = 0, int columnSpan = 0, int row = -1)
        {
            ProtectdTask(() =>
            {
                string type = null;
                if (temp.EndsWith(".ods"))
                {
                    if (node?.Name.StartsWith("text") ?? false)
                        node = node.ParentNode;
                    if (node?.Name.StartsWith("table:table-cell") ?? false)
                    {
                        var o = node.Attributes.GetNamedItem("office:value-type");
                        if (o != null)
                        {
                            type = o.Value;
                        }
                        if (node.FirstChild != null)
                            node = node.FirstChild;
                    }
                }                
                this.Cells.Add(new Cell()
                {
                    Node = node,
                    Range = this,
                    CellType = type,
                    Value = node.InnerText,
                    Row = row == -1 ? _totalRows : row,
                    RowSpan = rowSpan,
                    Column = column,
                    ColumnSpan = columnSpan
                });
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
            CellRange range = new CellRange(doc, temp, true, (TableNS, TableNS, CalcextNS, TextNS));
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
            var range = new CellRange(doc, temp, true, (TableNS, TableNS, CalcextNS, TextNS));
            for (int i = FirstRowIndex; i <= LastRowIndex; i++)
                for (int j = FirstColumnIndex; j <= LastColumnIndex; j++)
                    range.Cells.Add(this[i, j]);

            range._totalColumns = LastColumnIndex - FirstColumnIndex + 1;
            range._totalRows = LastRowIndex - FirstRowIndex + 1;
            range.EmptyNode = EmptyNode;
            return range;
        }
        private bool ProtectdTask(Action a)
        {
            try
            {
                a?.Invoke();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex} {ex.Message}");
                return false;
            }
            
        }
    }

}
