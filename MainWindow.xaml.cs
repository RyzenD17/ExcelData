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
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelData
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IExcelDataReader edr;
        private DataTableCollection tableCollection = null;
        public string fileName = "";
        public bool refFlag = false;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenExcelbtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() != true) return;
                fileName = openFileDialog.FileName;
                readFile(fileName);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!",MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void readFile(string fileNames)
        {
            try
            {
                if (refFlag == false)
                {
                    var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
                    FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
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

                    CBChooseList.Items.Clear();

                    foreach (DataTable dt in tableCollection)
                    {
                        CBChooseList.Items.Add(dt.TableName);
                    }
                    CBChooseList.SelectedIndex = 0;
                    edr.Close();
                }
                if (refFlag == true)
                {
                    var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
                    FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
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
                    CBChooseList.SelectedIndex = CBChooseList.SelectedIndex;
                    DataView dataView = tableCollection[Convert.ToString(CBChooseList.SelectedItem)].AsDataView();
                    edr.Close();
                    DbGrig.ItemsSource = null;
                    DbGrig.ItemsSource = dataView;
                    refFlag = false;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CBChooseList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                DataView dataView = tableCollection[Convert.ToString(CBChooseList.SelectedItem)].AsDataView();
                DbGrig.ItemsSource = dataView;
            }
            catch { return; }

        }

        private void SaveChangesbtn_Click(object sender, RoutedEventArgs e)
        {
            try 
            { 
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(fileName);
            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[Convert.ToString(CBChooseList.SelectedItem)];

            for (int j = 0; j < DbGrig.Columns.Count; j++)
            {
                Excel.Range myRange = (Excel.Range)sheet1.Cells[1, j + 1];
                sheet1.Columns[j + 1].ColumnWidth = 15;
                myRange.Value2 = DbGrig.Columns[j].Header;
            }

            for (int i = 0; i < DbGrig.Columns.Count; i++)
            {
                for (int j = 0; j < DbGrig.Items.Count; j++)
                {
                    TextBlock b = DbGrig.Columns[i].GetCellContent(DbGrig.Items[j]) as TextBlock;
                    Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                        if (b != null){
                            if (b.Text.ToString() != null || b.Text.ToString() != "")
                            {
                                myRange.Value2 = b.Text.ToString();
                            }
                            if (b.Text.ToString() == null || b.Text.ToString() == "")
                            {
                                myRange.Value2 = "";
                            }
                        }
                        else
                        {
                            myRange.Value2 = "";
                        }
                }
            }
            workbook.Save();
            workbook.Close(false);
            excel.Quit();
            excel = null;
            workbook = null;
            sheet1 = null;
            refFlag = true;
            readFile(fileName);
        }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void CalcPassbtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int pass=0; 
                int passUv = 0;
                int summpass = 0;
                int summpassuv = 0;
                string passuv ,passstr = "";
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(fileName);
                Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[Convert.ToString(CBChooseList.SelectedItem)];

                for (int i = 1; i < DbGrig.Items.Count-1; i++)
                {
                    for (int j = 2; j < DbGrig.Columns.Count-2; j++)
                    {
                        TextBlock b = DbGrig.Columns[j].GetCellContent(DbGrig.Items[i]) as TextBlock;

                        if (b.Text == "2" || b.Text == "4" || b.Text == "6" || b.Text== "8")
                        {
                            pass += Convert.ToInt32(b.Text.ToString());
                        }
                        if (b.Text.ToString() == "2*" || b.Text.ToString() == "4*" || b.Text.ToString() == "6*" || b.Text.ToString() == "8*")
                        {
                            passuv = b.Text.ToString().Replace("*", "");
                            pass += Convert.ToInt32(passuv);
                            passUv += Convert.ToInt32(passuv);
                        }
                    }

                    sheet1.Cells[i+2, DbGrig.Columns.Count-1]= Convert.ToString(pass);
                    sheet1.Cells[i+2, DbGrig.Columns.Count] = Convert.ToString(pass-passUv);
                    summpass += pass;
                    summpassuv += passUv;
                    passUv = 0;
                    pass = 0;
                    sheet1.Cells[DbGrig.Items.Count, DbGrig.Columns.Count - 1] = Convert.ToString(summpass);
                    sheet1.Cells[DbGrig.Items.Count, DbGrig.Columns.Count] = Convert.ToString(summpassuv);
                }
              
                workbook.Save();
                workbook.Close(false);
                excel.Quit();
                excel = null;
                workbook = null;
                sheet1 = null;
                refFlag = true;
                readFile(fileName);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }
    }
}
