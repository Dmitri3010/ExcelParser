using System;
using System.Collections.Generic;
using ExcelParser.MVVM;
using System.Linq;
using System.Windows;
using Spire.Xls;
using Spire.Xls.Core;


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
            var Worksheets = new Dictionary<string, IWorksheet>();
            var path = string.Empty;
            try
            {
                DefaultDialog.OpenFileDialog();
                path = DefaultDialog.FilePath;
            }
            catch (Exception ex)
            {
                DefaultDialog.ShowMessage(ex.Message);
            }

            var finalBook = new Workbook();
            finalBook.LoadFromFile("../../../EmptyTable.xlsx");

            var finalSheet = finalBook.Worksheets[0];

            var newBook = new Workbook();

            var newSheet = newBook.Worksheets[0];

            var workbook = new Workbook();

            workbook.LoadFromFile(path);

            foreach (var worksheet in workbook.Worksheets)
            {
                if (worksheet.CodeName.ToUpper() != "ИТОГ")
                {
                    Worksheets.Add(worksheet?.CodeName, worksheet);
                }
            }

//            var sheet = workbook.Worksheets[0];

            var i = 1;

            var columnCount = Worksheets.Count;

            CellRange source = finalSheet.Range[finalSheet.Columns[0].Row, 1, finalSheet.Columns[0].Row, columnCount];
            CellRange dest = newSheet.Range[i, 1, i, columnCount];
            finalSheet.Copy(source, dest, true);

            i = 2;
            foreach (var sheet in Worksheets)
            {
                newSheet.Range["B" + i].Value = sheet.Key;
                var ccc = sheet.Value.Cells.FirstOrDefault(p => p.Value.ToLower() == "итого");
                
                if (sheet.Value.Range.Value?.ToLower() == "итого")
                {
//                    CellRange sourceRange = sheet.Value.Range[sheet.Value.Range.Row, 1, sheet.Value.Range.Row];
//                    newSheet.InsertRow(i);
                    var cc = sheet.Value.Range;
                    // newSheet.Copy(sheet.Value.Range[sheet.Value.Range.Row, i, sheet.Value.Range.Row, "AI"+i ], newSheet.Range[$"D{i}:AI{i}"], true);

                }

//                newSheet.Range[sheet.Value.Columns[2].Row, 1, sheet.Value.Columns[0].Row, sheet.Value.Columns.Length]
//                for (int j = 0; j <= 32; j++)
//                {
//                    
//                }
                i++;
            }

//            foreach (var worksheet in Worksheets)
//            {
////                if (range.Text == "teacher")
////
////                {
//                CellRange sourceRange = sheet.Range[range.Row, 1, range.Row, columnCount];
//
//                CellRange destRange = newSheet.Range[i, 1, i, columnCount];
//
//                sheet.Copy(sourceRange, destRange, true);
//
//                i++;
////                }
//            }

            newBook.SaveToFile(Guid.NewGuid() + ".xlsx", ExcelVersion.Version2010);
        }
    }
}