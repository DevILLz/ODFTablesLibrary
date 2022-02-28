using ODFTablesLib;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window , INotifyPropertyChanged
    {
        private ObservableCollection<Cell> _items;

        public ObservableCollection<Cell> Items { get => _items; set => SetProperty(ref _items, value); }
        private int _column;
        private int _row;
        public int Column { get => _column; set => SetProperty(ref _column, value); }
        public int Row { get => _row; set => SetProperty(ref _row, value); }
        private CellRange cells;
        private ODFTables odf;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            odf = new ODFTables(Path.GetFullPath(@"C:\Users\serov.KBNT-SEROV\Desktop\Работа\ProtocolTemplates\SA640 temp.ods"));
            cells = odf.Cells;
            Items = new ObservableCollection<Cell>(cells.Cells);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        protected bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(storage, value))
            {
                return false;
            }

            storage = value;
            NotifyPropertyChanged(propertyName);
            return true;
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            cells[Row, Column].Value = "xx";
            odf.Save(@"C:\Users\serov.KBNT-SEROV\Desktop\test.ods");

        }
    }
}
