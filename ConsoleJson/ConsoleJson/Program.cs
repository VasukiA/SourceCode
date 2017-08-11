using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data.Common;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;



namespace ConsoleJson
{
    class Program
    {
        static void Main(string[] args)
        {
            var ExcelPath = @"E:\input1.xlsx";
            var destinationPath = @"E:\output.json";
            int sheetNo = 1;
            var xlsheetName = "AssetList";

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range = null;


            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(ExcelPath);

            var connectionString = String.Format(@"
                Provider=Microsoft.ACE.OLEDB.12.0;
                Data Source={0};
                Extended Properties=""Excel 12.0 Xml;HDR=YES""", ExcelPath);
            File.AppendAllText(destinationPath, xlsheetName);

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in xlWorkBook.Sheets)
            {
                string sheetName = sheet.Name;
                File.AppendAllText(destinationPath, sheetName);
                //Creating and opening a data connection to the Excel sheet 
                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    var cmd = conn.CreateCommand();
                    cmd.CommandText = String.Format(
                        @"SELECT * FROM [{0}$]",
                        sheetName
                    );

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetNo);
                    sheetNo++;
                    range = xlWorkSheet.UsedRange;
                    int rowRange = range.Rows.Count;
                    int columnRange = range.Columns.Count;

                    int columnCount = 0;

                    using (var rdr = cmd.ExecuteReader())
                    {
                        var query =
                       (from DbDataRecord row in rdr
                        select row).Select(x =>
                        {
                            Dictionary<string, object> item = new Dictionary<string, object>();
                            for (columnCount = 0; columnCount < columnRange; columnCount++)
                            {
                                item.Add(rdr.GetName(columnCount), x[columnCount]);
                            }
                            return item;
                        });

                        var json = JsonConvert.SerializeObject(query);

                        // to write a json format text in the file
                        File.AppendAllText(destinationPath, json);
                    }

                }
                File.AppendAllText(destinationPath, "]");
            }
            File.AppendAllText(destinationPath, "}");
        }
    }
}
