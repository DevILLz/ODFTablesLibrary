using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Xml;

namespace ODFTablesLib
{
    public class Cell : INotifyPropertyChanged
    {
        private string cellType;

        /// <summary>
        /// Do not ask why name is [row,columt] instead of [column,row]
        /// I dont have right answer
        /// </summary>
        public Cell() { }
        public CellRange MergedRange { get; set; }
        public CellRange Range { get; set; }
        internal XmlNode Node { get; set; }
        public string CellType
        { 
            get => cellType; 
            set
            {
                if (cellType != value)
                {
                    cellType = value;
                    Range.ChangeCellType(Node, value);
                }               
            } 
        }
        public int RowSpan { get; set; }
        public int ColumnSpan { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public string Name { get => $"{GetColumnName(this.Column)}{this.Row}"; }
        public string Value
        {
            get => Node.InnerText;
            set
            {
                if (!string.IsNullOrWhiteSpace(value) && (Node?.Name.StartsWith("table:table-cell") ?? false) && !(Node.FirstChild?.Name.StartsWith("text") ?? false))
                {
                    Node = Range.AddTextBlock(Node);
                }
                Node.InnerText = value;
                var currentCulture = CultureInfo.CurrentCulture;
                CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
                if (float.TryParse(value, out float o))
                {
                    CellType = "float";
                    Node.ParentNode.Attributes.GetNamedItem("office:value").Value = value;
                }                    
                CultureInfo.CurrentCulture = currentCulture;
                NotifyPropertyChanged();
            }
        }
        #region Don't look inside
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
        #endregion
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
        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
