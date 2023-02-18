


using ExcelDataReader;
using System.Collections;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
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
using System.Data;
using System.IO;
// дата создания 18.02.2023 время 8-23
// https://blogadminday.ru/wpf-import-xls-xlsx-faylov/
// пойду на лыжах кататься. нога болит пиии, гружа опять
namespace Autocad_exel_read_and_to_create_lisp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IExcelDataReader edr;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void load_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            // меню загрузки, сначала все файлы
            openFileDialog.Filter = "EXCEL Files (All files (*.*)|*.*|*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls)";
            if (openFileDialog.ShowDialog() != true)
                return;

            DbGrig.ItemsSource = readFile(openFileDialog.FileName);
        }
        private DataView readFile(string fileNames)
        {

            // to do меняем кодировку
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
            // Создаем поток для чтения.
           
            FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
            // В зависимости от расширения файла Excel, создаем того или иного способа прочтения.
            // Читатель для файлов с расширением *.xlsx.
            if (extension == ".xlsx")
              edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //// Читатель для файлов с расширением *.xls.
            else if (extension == ".xls")
              edr = ExcelReaderFactory.CreateBinaryReader(stream);

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            // Читаем, получаем DataView и работаем с ним.
            DataSet dataSet = edr.AsDataSet(conf);
            DataView dtView = dataSet.Tables[0].AsDataView();

            // После завершения чтения освобождаем ресурсы.
            edr.Close();
            return dtView;
        }


        // https://www.youtube.com/watch?v=JiEKqLdvnyY
        //// экспорт в excel
        //private void btnExportExcel_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        //        excel.Visible = true;
        //        Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
        //        Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
        //        int StartCol = 1;
        //        int StartRow = 1;
        //        int j = 0, i = 0;

        //        //Write Headers
        //        for (j = 0; j < dgvSource.Columns.Count; j++)
        //        {
        //            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
        //            myRange.Value2 = dgvSource.Columns[j].HeaderText;
        //        }

        //        StartRow++;

        //        //Write datagridview content
        //        for (i = 0; i < dgvSource.Rows.Count; i++)
        //        {
        //            for (j = 0; j < dgvSource.Columns.Count; j++)
        //            {
        //                try
        //                {
        //                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
        //                    myRange.Value2 = dgvSource[j, i].Value == null ? "" : dgvSource[j, i].Value;
        //                }
        //                catch
        //                {
        //                    ;
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}
    }

}
