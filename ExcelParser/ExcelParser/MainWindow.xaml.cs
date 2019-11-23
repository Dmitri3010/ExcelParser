using ExcelParser.MVVM;
using System.Linq;
using System.Windows;
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
            DefaultDialog.OpenFileDialog();
            var path = DefaultDialog.FilePath;
            MessageBox.Show(path);

            var newBook = new Workbook();

            var newSheet = newBook.Worksheets[0];

            var workbook = new Workbook();

            workbook.LoadFromFile(path);

            var sheet = workbook.Worksheets[0];

            var i = 1;

            var columnCount = sheet.Columns.Count();
            foreach (CellRange range in sheet.Columns[2])
            {
//                if (range.Text == "teacher")
//
//                {
                CellRange sourceRange = sheet.Range[range.Row, 1, range.Row, columnCount];

                CellRange destRange = newSheet.Range[i, 1, i, columnCount];

                sheet.Copy(sourceRange, destRange, true);

                i++;
//                }
            }

            newBook.SaveToFile("NewForm.xlsx", ExcelVersion.Version2010);
        }
    }
}