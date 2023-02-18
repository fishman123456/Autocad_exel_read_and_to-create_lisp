


using ExcelDataReader;
using System.Collections;
using Microsoft.Win32;

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
            openFileDialog.Filter = "EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";
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
    }

}
