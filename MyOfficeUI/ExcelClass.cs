using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Microsoft.Office.Interop;
using System.IO;
namespace MyOffice
{
    public class ExcelClass
    {
        public static System.Data.DataTable LoadFile(string FileName)
        {
           
            System.Data.DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                System.Data.DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }

        public static void SaveFile(System.Data.DataTable dataTable, string filepath)
        {
            SaveFile(dataTable, filepath, new int[] { });
        }
        public static void SaveFile(System.Data.DataTable dataTable, string filepath ,int[] ColunmWidth )
        {
            int[] colunmWidth = new int[dataTable.Columns.Count];
            for (int i = 0; i < colunmWidth.Length; i++)
            {
                colunmWidth[i] = 10;
            }
            for(int i = 0 ; i < ColunmWidth.Length ; i ++)
            {
                if (i < colunmWidth.Length) colunmWidth[i] = ColunmWidth[i];
            }
            try
            {
                //建立Excel應用程式類的一個例項，相當於從電腦開始選單開啟Excel
                Microsoft.Office.Interop.Excel.ApplicationClass xlsxapp = new Microsoft.Office.Interop.Excel.ApplicationClass();
                //新建一張Excel工作簿
                Microsoft.Office.Interop.Excel.Workbook wbook = xlsxapp.Workbooks.Add(true);
                //第一個sheet頁
                Microsoft.Office.Interop.Excel.Worksheet wsheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets.get_Item(1);
                //將DataTable的列名顯示在Excel表第一行
                wsheet.Cells.NumberFormatLocal = "@";

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    //注意Excel表的行和列的索引都是從1開始的
                    ((Microsoft.Office.Interop.Excel.Range)wsheet.Columns[Convert_Colums(i), System.Type.Missing]).ColumnWidth = colunmWidth[i];
                    wsheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName;
                }
                //遍歷DataTable，給Excel賦值
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        //從第二行第一列開始寫入資料

                        wsheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j].ToString();
                    }
                }

                try
                {
                    //儲存檔案
                    wbook.SaveAs(filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                    null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
                    false, false, null, null, null);
                }
                catch
                {
                    Console.WriteLine("Excel 存檔失敗!");
                }
                finally
                {
                    string ProcessName = "WINWORD";//換成想要結束的進程名字
                    System.Diagnostics.Process[] MyProcess = System.Diagnostics.Process.GetProcessesByName(ProcessName);
                    for (int i = 0; i < MyProcess.Length; i++)
                    {
                        MyProcess[i].Kill();
                    }
                    ProcessName = "EXCEL";//換成想要結束的進程名字
                    MyProcess = System.Diagnostics.Process.GetProcessesByName(ProcessName);
                    for (int i = 0; i < MyProcess.Length; i++)
                    {
                        MyProcess[i].Kill();
                    }
                    //釋放資源
                    xlsxapp.Quit();
                }
            }
            catch
            {
                Console.WriteLine("找不到Excel安裝資訊!");
            }

         
      
        }
        private static string Convert_Colums(int index)
        {
            string str = "";
            int value_00 = -1;
            int value_01 = -1;
            value_00 = index % 26;
            if (index >= 26)
            {
                value_01 = index / 26;
                value_01--;
            }
            if (value_01 != -1)
            {
                str += (char)(value_01 + 65);
            }
            if (value_00 != -1)
            {
                str += (char)(value_00 + 65);
            }
            
            return str;
        }
    }
}
