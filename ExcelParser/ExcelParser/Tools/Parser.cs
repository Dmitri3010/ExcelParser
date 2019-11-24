using System;
using System.Collections.Generic;
using System.Linq;
using Spire.Xls;
using Spire.Xls.Core;

namespace ExcelParser.Tools
{
    public static class Parser
    {
        private static Dictionary<string, IWorksheet> Worksheets { get; set; }

        public static Workbook Parse(Workbook workbook)
        {
            Worksheets = new Dictionary<string, IWorksheet>();
            foreach (var worksheet in workbook.Worksheets)
            {
                if (worksheet.CodeName.ToUpper() != "ИТОГ" && new string(worksheet.CodeName.Take(4).ToArray())
                    != "Лист")
                {
                    Worksheets.Add(worksheet?.CodeName, worksheet);
                }
            }

            var i = 1;


            var finalBook = new Workbook();
            finalBook.LoadFromFile("../../../EmptyTable.xlsx");

            var finalSheet = finalBook.Worksheets[0];
            var newBook = new Workbook();
            var columnCount = finalBook.Worksheets[0].Columns.Length;
//            var emptyrow = finalSheet.Columns.FirstOrDefault(p => p.Value.ToLower() == "итого");


            var source = finalSheet.Range[finalSheet.Columns[0].Row, 1, finalSheet.Columns[0].Row, columnCount];
            var dest = newBook.Worksheets[0].Range[i, 1, i, columnCount];
            finalSheet.Copy(source, dest, true);

            i = 2;
            foreach (var (key, value) in Worksheets)
            {
                newBook.Worksheets[0].Range["B" + i].Value = key;
                var currentSheet = value.Cells.FirstOrDefault(p => p.Value.ToLower() == "итого");
                if (currentSheet != null)
                {
                    var worksheet = (Worksheet) value;
                    //Loop through cells
                    foreach (CellRange cell in worksheet.Range)
                    {
                        //If the cell contain formula, get the formula value, clear cell content, and then fill the formula value into the cell 
                        if (cell.HasFormula)
                        {
                            Object values = cell.FormulaValue;                        
                            cell.Clear(ExcelClearOptions.ClearContent);
                            cell.Value2 = values;
                        }
                    }
                    newBook.Worksheets[0].Copy(worksheet.Range[currentSheet.Row, 4, currentSheet.Row, 32],
                        newBook.Worksheets[0].Range[$"D{i}:AI{i}"], true);
                }

                i++;
            }

            for (var j = 0; j <= 1; j++)
            {
                newBook.Worksheets[1].Remove();
            }

            newBook.ActiveSheetIndex = 0;
//            var rowCount = newBook.Worksheets[0].Rows.Length;
//            var destiny = newBook.Worksheets[0].Range[i, 1, i, columnCount];
            var sourcesheet = finalSheet.Range[$"A21:AI27"];
            var desty = newBook.Worksheets[0].Range[$"A{(Worksheets.Count+8)}:AI{Worksheets.Count+8}"];
            finalSheet.Copy(sourcesheet, desty, true);
//            finalSheet.Copy(emptyrow, dest, true);
            //newBook.Worksheets[0].Range["B" + (Worksheets.Count+2)].Value = "Итого";
//            var colums = newBook.Worksheets[0].Columns;
//            var cells = newBook.Worksheets[0].Columns[4].Cells;

//            var bbb = new Dictionary<int, int>();
//            foreach (var cell in cells.Where(p => !p.IsBlank).Skip(1))
//            {
//                bbb.Add(cell.Column, Convert.ToInt32(cell.DisplayedText));
//            }
            for (int k = newBook.Worksheets[0].Rows.Count() - 1; k >= 0; k--)
            {
                if (newBook.Worksheets[0].Rows[k].IsBlank)
                {
                    newBook.Worksheets[0].DeleteRow(k+1);
                }
            }

            var finalCells =  newBook.Worksheets[0].Range[$"A{(Worksheets.Count+2)}:AI{Worksheets.Count+2}"];;
            foreach (var cell in finalCells.Skip(3))
            {
                cell.Value = newBook.Worksheets[0]
                    .Columns[cell.Column-1]
                    .Cells.Where(p=> p.DisplayedText != string.Empty)
                    .Skip(1)
                    .Sum(p => Convert.ToInt32(p.DisplayedText)).ToString();
            }

            return newBook;
        }
    }
}