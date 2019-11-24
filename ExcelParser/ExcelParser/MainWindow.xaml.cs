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

        public MainWindow()
        {
            InitializeComponent();
            DefaultDialog = new DefaultDialogService();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Workbook Newbook;
            string path;
            var currentBook = new Workbook();
            try
            {
                DefaultDialog.OpenFileDialog();
                path = DefaultDialog.FilePath;

                currentBook.LoadFromFile(path);
            }
            catch (Exception ex)
            {
                DefaultDialog.ShowMessage(ex.Message);
            }

            try
            {
                Newbook = Parser.Parse(currentBook);
            }
            catch (Exception ex)
            {
                DefaultDialog.ShowMessage(ex.Message);
                DefaultDialog.ShowMessage("Что-то пошло не так. Попробуйте еще раз");
                throw new Exception();
            }

            DefaultDialog.ShowMessage("Отчет успешно сформирован! Выберите название для файла");
            DefaultDialog.SaveFileDialog();
            path = DefaultDialog.FilePath;
            Newbook.SaveToFile(path + ".xlsx", ExcelVersion.Version2010);
        }
    }
}