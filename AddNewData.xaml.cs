using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
using ExcelDataReader;
using Microsoft.Win32;

namespace ExcelData
{
    /// <summary>
    /// Логика взаимодействия для AddNewData.xaml
    /// </summary>
    public partial class AddNewData : Window
    {
        public int _index;
        public string _file;
        IExcelDataReader edr;
        private DataTableCollection tableCollection = null;
        public AddNewData(int index,string fileName)
        {
            InitializeComponent();
            _index = index;
            _file = fileName;
            DataList.Text = "Номер листа - " + Convert.ToString(_index+1);
            var extension = _file.Substring(_file.LastIndexOf('.'));
            FileStream stream = File.Open(_file, FileMode.Open, FileAccess.Read);
            // Читатель для файлов с расширением *.xlsx.
            if (extension == ".xlsx")
                edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
            // Читатель для файлов с расширением *.xls.
            else if (extension == ".xls")
                edr = ExcelReaderFactory.CreateBinaryReader(stream);
            //// reader.IsFirstRowAsColumnNames

            DataSet dataSet = edr.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });
            tableCollection = dataSet.Tables;
            edr.Close();
            DataView dataView = tableCollection[_index].AsDataView();
            DataTable dataTable = dataView.Table;
            CmbDate.ItemsSource = dataTable.Columns.ToString();


        }

       

    }
}
