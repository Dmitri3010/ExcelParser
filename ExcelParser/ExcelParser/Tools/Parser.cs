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
                if (worksheet.CodeName.ToUpper() != "ИТОГ")
                {
                    Worksheets.Add(worksheet?.CodeName, worksheet);
                }
            }

            var i = 1;

            var columnCount = Worksheets.Count;

            var finalBook = new Workbook();
            finalBook.LoadFromFile("../../../EmptyTable.xlsx");

            var finalSheet = finalBook.Worksheets[0];
            var newBook = new Workbook();


            var source = finalSheet.Range[finalSheet.Columns[0].Row, 1, finalSheet.Columns[0].Row, columnCount];
            var dest = newBook.Worksheets[0].Range[i, 1, i, columnCount];
            finalSheet.Copy(source, dest, true);

            i = 2;
            foreach (var sheet in Worksheets)
            {
                newBook.Worksheets[0].Range["B" + i].Value = sheet.Key;
                var currentSheet = sheet.Value.Cells.FirstOrDefault(p => p.Value.ToLower() == "итого");
                if (currentSheet != null)
                {
                    var worksheet = (Worksheet) sheet.Value;
                    newBook.Worksheets[0].Copy(worksheet.Range[currentSheet.Row, 4, currentSheet.Row, 32],
                        newBook.Worksheets[0].Range[$"D{i}:AI{i}"], true);
                }

                i++;
            }

            return newBook;
        }
    }
}