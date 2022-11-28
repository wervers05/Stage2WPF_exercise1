using Microsoft.Win32;
using Stage2WPF.MVVM.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Stage2WPF.Core;

namespace Stage2WPF.MVVM.View
{
    /// <summary>
    /// Interaction logic for ExcelToDataGridView.xaml
    /// </summary>
    public partial class ExcelToDataGridView : UserControl, INotifyPropertyChanged
    {
        private ICollectionView _dataGridCollection;
        private string _filterString;

        public ExcelToDataGridView()
        {
            InitializeComponent();
            DataGridCollection = CollectionViewSource.GetDefaultView(TestData);
            DataGridCollection.Filter = new Predicate<object>(Filter);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        ExcelParsing excel = new ExcelParsing();

        public ICollectionView DataGridCollection
        {
            get { return _dataGridCollection; }
            set { _dataGridCollection = value; NotifyPropertyChanged("DataGridCollection"); }
        }

        public string FilterString
        {
            get { return _filterString; }
            set
            {
                _filterString = value;
                NotifyPropertyChanged("FilterString");
                FilterCollection();
            }
        }
        private void FilterCollection()
        {
            if (_dataGridCollection != null)
            {
                _dataGridCollection.Refresh();
            }
        }
        public bool Filter(object obj)
        {
            var data = obj as ExcelModel;
            if (data != null)
            {
                if (!string.IsNullOrEmpty(_filterString))
                {
                    return data.Rep.Contains(_filterString); // || data.Email.Contains(_filterString)
                }
                return true;
            }
            return false;
        }
        public IEnumerable<ExcelModel> TestData
        {
            get
            {
                if (excelDataGrid.Items.Count > 0)
                    foreach (var item in excelDataGrid.Items.OfType<ExcelModel>())
                    {
                        yield return item;
                    }
            }
        }

        private void NotifyPropertyChanged(string property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }

        private void OpenFileDialog_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFD = new OpenFileDialog();

            string xlFileName;

            openFD.Filter = "Excel Files | *.xls; *xlsx";
            openFD.Title = "Import Excel File";
            var browseFile = openFD.ShowDialog();
            xlFileName = openFD.FileName;
            textFileName.Text = openFD.FileName;
            
            if (browseFile == true)
            {
                excel.XlsxFileName = xlFileName;
                excelDataGrid.ItemsSource = excel.getExcelData();
            }
            #region < ----- >
            
            #endregion
        }

        private void searchByName_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                excelDataGrid.Items.Clear();
                excel.XlsxFileName = textFileName.Text;
                excelDataGrid.ItemsSource = excel.getExcelDataByRep(textBoxSearch.Text);
            }
            catch (Exception ex) { }
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                excelDataGrid.Items.Clear();
                excel.XlsxFileName = textFileName.Text;
                excelDataGrid.ItemsSource = excel.getExcelDataByRegion(cmbSelectRegion.Text);
            }
            catch (Exception ex) { }
        }
    }

    
}
