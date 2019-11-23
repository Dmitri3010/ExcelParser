using System;
using System.IO;
using ExcelParser.MVVM;
using System.Windows;
using ExcelParser.Tools;
using Spire.Xls;


namespace ExcelParser
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DefaultDialogService DefaultDialog { get; }
//        private  Workbook NewBook => new Workbook()

//        private static Workbook CurrentBook => new Workbook();


        public MainWindow()
        {
            InitializeComponent();
            DefaultDialog = new DefaultDialogService();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var Newbook = new Workbook();
            var path = string.Empty;
            var CurrentBook = new Workbook();
            try
            {
                DefaultDialog.OpenFileDialog();
                path = DefaultDialog.FilePath;
//                var workbook = new Workbook();

                CurrentBook.LoadFromFile(path);
            }
            catch (Exception ex)
            {
                DefaultDialog.ShowMessage(ex.Message);
            }

//            var finalBook = new Workbook();
//            finalBook.LoadFromFile("../../../EmptyTable.xlsx");
//
//            var finalSheet = finalBook.Worksheets[0];

//            var newBook = new Workbook();
//
//            var newSheet = newBook.Worksheets[0];

            try
            {
                Newbook = Parser.Parse(CurrentBook);
            }
            catch (Exception ex)
            {
                DefaultDialog.ShowMessage(ex.Message);
                DefaultDialog.ShowMessage("Что-то пошло не так. Попробуйте еще раз");
                throw new Exception();
            }
//            finally
//            {
//                
//            } 

            DefaultDialog.ShowMessage("Отчет успешно сформирован! Выберите название для файла");
            DefaultDialog.SaveFileDialog();
            path = DefaultDialog.FilePath;
            Newbook.SaveToFile(path + ".xlsx", ExcelVersion.Version2010);
            
            DefaultDialog.ShowMessage($"Отчет успешно сохранен в папке: {Directory.GetCurrentDirectory()}");
        }
    }
}