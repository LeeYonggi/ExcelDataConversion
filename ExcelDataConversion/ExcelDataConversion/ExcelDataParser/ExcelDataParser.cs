using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDataConversion
{

    public static class ExcelDataParser
    {
        public static DataTable ExcelFileRead(string filePath)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Open(filePath);
            Excel.Worksheet sheet = workbook.Worksheets.Item[1] as Excel.Worksheet;

            Excel.Range range = sheet.UsedRange;

            object[,] data = range.Value;

            DataTable dataTable = new DataTable();
            for(int i = 1; i <= range.Columns.Count; i++)
            {
                dataTable.Columns.Add(i.ToString(), typeof(string));
            }

            for (int r = 1; r <= range.Rows.Count; r++)
            {
                DataRow dr = dataTable.Rows.Add();

                for (int c = 1; c <= range.Columns.Count; c++)
                {
                    dr[c - 1] = data[r, c];
                }
            }

            workbook.Close(true);
            app.Quit();
            DeleteObject(sheet);
            DeleteObject(app);
            DeleteObject(workbook);

            return dataTable;
        }

        public static void DataTableToJson(string savePath, DataTable dataTable)
        {
            string data = string.Empty;
            JObject jObjects = new JObject();

            for(int obj = 1; obj < dataTable.Rows.Count; obj++)
            {
                JObject sonSpec = new JObject();

                for(int i = 0; i < dataTable.Columns.Count; i++)
                {
                    sonSpec.Add(new JProperty(dataTable.Rows[0].ItemArray[i] as string, dataTable.Rows[obj].ItemArray[i]));
                }

                jObjects.Add(dataTable.Rows[obj].ItemArray[0] as string, sonSpec);
            }

            data += JsonConvert.SerializeObject(jObjects, Formatting.Indented);

            File.WriteAllText(savePath, data);
        }

        private static void DeleteObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch(Exception ex)
            {
                obj = null;
                Console.WriteLine(ex.ToString(), obj);
            }
        }
    }
}
