using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.IO;
using NPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using System.Drawing;
using Basic;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.ComponentModel;

namespace MyOffice
{
    [Serializable]
    public class SheetClass
    {
        List<CellValue> cellValues = new List<CellValue>();

        public List<CellValue> CellValues { get => cellValues; set => cellValues = value; }

        public List<ICellStyle> cellStyles = new List<ICellStyle>();


        public bool Add(CellValue cellValue)
        {
            List<CellValue> CellValues_buf = new List<CellValue>();
            CellValues_buf = (from value in CellValues
                              where value.RowStart == cellValue.RowStart
                              where value.RowEnd == cellValue.RowEnd
                              where value.ColStart == cellValue.ColStart
                              where value.ColEnd == cellValue.ColEnd
                              select value).ToList();
            if (CellValues_buf.Count == 0)
            {
                this.CellValues.Add(cellValue);
                return true;
            }
            return false;
        }
        public void Init()
        {
            HSSFWorkbook workbook = new NPOI.HSSF.UserModel.HSSFWorkbook();

        }
    }
    [Serializable]
    public class CellValue
    {
        static public class ColorSerializationHelper
        {
            static public Color FromString(string value)
            {
                var parts = value.Split(':');

                int A = 0;
                int R = 0;
                int G = 0;
                int B = 0;
                int.TryParse(parts[0], out A);
                int.TryParse(parts[1], out R);
                int.TryParse(parts[2], out G);
                int.TryParse(parts[3], out B);
                return Color.FromArgb(A, R, G, B);
            }
            static public string ToString(Color color)
            {
                return color.A + ":" + color.R + ":" + color.G + ":" + color.B;

            }
        }
        [TypeConverter(typeof(FontConverter))]
        static public class FontSerializationHelper
        {
            static public Font FromString(string value)
            {
                var parts = value.Split(':');
                return new Font(
                    parts[0],                                                   // FontFamily.Name
                    float.Parse(parts[1]),                                      // Size
                    EnumSerializationHelper.FromString<FontStyle>(parts[2]),    // Style
                    EnumSerializationHelper.FromString<GraphicsUnit>(parts[3]), // Unit
                    byte.Parse(parts[4]),                                       // GdiCharSet
                    bool.Parse(parts[5])                                        // GdiVerticalFont
                );
            }
            static public string ToString(Font font)
            {
                return font.FontFamily.Name
                        + ":" + font.Size
                        + ":" + font.Style
                        + ":" + font.Unit
                        + ":" + font.GdiCharSet
                        + ":" + font.GdiVerticalFont
                        ;
            }
        }
        [TypeConverter(typeof(EnumConverter))]
        static public class EnumSerializationHelper
        {
            static public T FromString<T>(string value)
            {
                return (T)Enum.Parse(typeof(T), value, true);
            }
        }
        
        private string text = "";
        private int rowStart = 0;
        private int rowEnd = 0;
        private int colStart = 0;
        private int colEnd = 0;

        private MyFont font;
        private MyCellStyle cellStyle;
        public string Text { get => text; set => text = value; }
        public int RowStart { get => rowStart; set => rowStart = value; }
        public int RowEnd { get => rowEnd; set => rowEnd = value; }
        public int ColStart { get => colStart; set => colStart = value; }
        public int ColEnd { get => colEnd; set => colEnd = value; }
        public MyFont Font { get => font; set => font = value; }
        public MyCellStyle CellStyle { get => cellStyle; set => cellStyle = value; }


 
        public static ICellStyle ToICellStytle(NPOI.SS.UserModel.IWorkbook workbook, MyCellStyle myCellStyle, MyFont font)
        {
            ICellStyle cellStyle = MyCellStyle.ToICellStyle(workbook, myCellStyle);
            cellStyle.SetFont(MyFont.ToIFont(workbook, font));
            return cellStyle;
            
        }
    }
    [Serializable]
    public class MyFont
    {
        public string FontName { get; set; }
        public double FontHeight { get; set; }
        public double FontHeightInPoints { get; set; }
        public bool IsItalic { get; set; }
        public bool IsStrikeout { get; set; }
        public short Color { get; set; }
        public FontSuperScript TypeOffset { get; set; }
        public FontUnderlineType Underline { get; set; }
        public short Charset { get; set; }
        public short Index { get; set; }
        public short Boldweight { get; set; }
        public bool IsBold { get; set; }

        public static MyFont ToMyFont(IFont font)
        {
            MyFont myFont = new MyFont();
            myFont.FontName = font.FontName;
            myFont.FontHeight = font.FontHeight;
            myFont.FontHeightInPoints = font.FontHeightInPoints;
            myFont.IsItalic = font.IsItalic;
            myFont.IsStrikeout = font.IsStrikeout;
            myFont.Color = font.Color;
            myFont.TypeOffset = font.TypeOffset;
            myFont.Underline = font.Underline;
            myFont.Charset = font.Charset;
            myFont.Index = font.Index;
            myFont.Boldweight = font.Boldweight;
            myFont.IsBold = font.IsBold;

            return myFont;
        }
        public static IFont ToIFont(NPOI.SS.UserModel.IWorkbook workbook ,MyFont font)
        {
            IFont Ifont = workbook.FindFont(font.Boldweight, font.Color, (short)font.FontHeight, font.FontName, font.IsItalic, font.IsStrikeout, font.TypeOffset, font.Underline);
            if (Ifont != null) return Ifont;



            IFont myFont = workbook.CreateFont();
            myFont.FontName = font.FontName;
            myFont.FontHeight = font.FontHeight;
            myFont.FontHeightInPoints = font.FontHeightInPoints;
            myFont.IsItalic = font.IsItalic;
            myFont.IsStrikeout = font.IsStrikeout;
            myFont.Color = font.Color;
            myFont.TypeOffset = font.TypeOffset;
            myFont.Underline = font.Underline;
            myFont.Charset = font.Charset;
            myFont.Boldweight = font.Boldweight;
            myFont.IsBold = font.IsBold;

            return myFont;
        }
     
    }
    [Serializable]
    public class MyCellStyle
    {
        public BorderStyle BorderLeft { get; set; }
        public BorderDiagonal BorderDiagonal { get; set; }
        public BorderStyle BorderDiagonalLineStyle { get; set; }
        public short BorderDiagonalColor { get; set; }
        public short FillForegroundColor { get; set; }
        public short FillBackgroundColor { get; set; }
        public FillPattern FillPattern { get; set; }
        public short BottomBorderColor { get; set; }
        public short TopBorderColor { get; set; }
        public short RightBorderColor { get; set; }
        public short LeftBorderColor { get; set; }
        public BorderStyle BorderBottom { get; set; }
        public BorderStyle BorderTop { get; set; }
        public BorderStyle BorderRight { get; set; }
        //public IColor FillBackgroundColorColor { get; set; }
        //public IColor FillForegroundColorColor { get; set; }
        public short Rotation { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }
        public bool WrapText { get; set; }
        public HorizontalAlignment Alignment { get; set; }
        public bool IsLocked { get; set; }
        public bool IsHidden { get; set; }
        public short FontIndex { get; set; }
        public short DataFormat { get; set; }
        public short Index { get; set; }
        public bool ShrinkToFit { get; set; }
        public short Indention { get; set; }
        public static MyCellStyle ToMyCellStyle(ICellStyle cellStyle)
        {
            MyCellStyle myCellStyle = new MyCellStyle();
            myCellStyle.BorderLeft = cellStyle.BorderLeft;
            myCellStyle.BorderDiagonal = cellStyle.BorderDiagonal;
            myCellStyle.BorderDiagonalLineStyle = cellStyle.BorderDiagonalLineStyle;
            myCellStyle.BorderDiagonalColor = cellStyle.BorderDiagonalColor;
            myCellStyle.FillForegroundColor = cellStyle.FillForegroundColor;
            myCellStyle.FillBackgroundColor = cellStyle.FillBackgroundColor;
            myCellStyle.FillPattern = cellStyle.FillPattern;
            myCellStyle.BottomBorderColor = cellStyle.BottomBorderColor;
            myCellStyle.TopBorderColor = cellStyle.TopBorderColor;
            myCellStyle.RightBorderColor = cellStyle.RightBorderColor;
            myCellStyle.LeftBorderColor = cellStyle.LeftBorderColor;
            myCellStyle.BorderBottom = cellStyle.BorderBottom;
            myCellStyle.BorderTop = cellStyle.BorderTop;
            myCellStyle.BorderRight = cellStyle.BorderRight;
            //myCellStyle.FillBackgroundColorColor = cellStyle.FillBackgroundColorColor;
            //myCellStyle.FillForegroundColorColor = cellStyle.FillForegroundColorColor;
            myCellStyle.Rotation = cellStyle.Rotation;
            myCellStyle.VerticalAlignment = cellStyle.VerticalAlignment;
            myCellStyle.WrapText = cellStyle.WrapText;
            myCellStyle.Alignment = cellStyle.Alignment;
            myCellStyle.IsLocked = cellStyle.IsLocked;
            myCellStyle.IsHidden = cellStyle.IsHidden;
            myCellStyle.FontIndex = cellStyle.FontIndex;
            myCellStyle.DataFormat = cellStyle.DataFormat;
            myCellStyle.Index = cellStyle.Index;
            myCellStyle.ShrinkToFit = cellStyle.ShrinkToFit;
            myCellStyle.Indention = cellStyle.Indention;

            return myCellStyle;
        }
        public static ICellStyle ToICellStyle(NPOI.SS.UserModel.IWorkbook workbook, MyCellStyle cellStyle)
        {
            //for (short i = 0; i < workbook.NumCellStyles; i++)
            //{
            //    ICellStyle myCellStyle_buf = workbook.GetCellStyleAt(i);
            //    bool flag_ok = true;
            //    if (myCellStyle_buf.Alignment != cellStyle.Alignment) flag_ok = false;
            //    if (myCellStyle_buf.VerticalAlignment != cellStyle.VerticalAlignment) flag_ok = false;
            //    if (flag_ok)
            //    {
            //        return myCellStyle_buf;
            //    }
            //}
            ICellStyle myCellStyle = workbook.CreateCellStyle();
            myCellStyle.BorderLeft = cellStyle.BorderLeft;
            myCellStyle.BorderDiagonal = cellStyle.BorderDiagonal;
            myCellStyle.BorderDiagonalLineStyle = cellStyle.BorderDiagonalLineStyle;
            myCellStyle.BorderDiagonalColor = cellStyle.BorderDiagonalColor;
            myCellStyle.FillForegroundColor = cellStyle.FillForegroundColor;
            myCellStyle.FillBackgroundColor = cellStyle.FillBackgroundColor;
            myCellStyle.FillPattern = cellStyle.FillPattern;
            myCellStyle.BottomBorderColor = cellStyle.BottomBorderColor;
            myCellStyle.TopBorderColor = cellStyle.TopBorderColor;
            myCellStyle.RightBorderColor = cellStyle.RightBorderColor;
            myCellStyle.LeftBorderColor = cellStyle.LeftBorderColor;
            myCellStyle.BorderBottom = cellStyle.BorderBottom;
            myCellStyle.BorderTop = cellStyle.BorderTop;
            myCellStyle.BorderRight = cellStyle.BorderRight;
   
            myCellStyle.Rotation = cellStyle.Rotation;
            myCellStyle.VerticalAlignment = cellStyle.VerticalAlignment;
            myCellStyle.WrapText = cellStyle.WrapText;
            myCellStyle.Alignment = cellStyle.Alignment;
            myCellStyle.IsLocked = cellStyle.IsLocked;
            myCellStyle.IsHidden = cellStyle.IsHidden;
            myCellStyle.DataFormat = cellStyle.DataFormat;
            myCellStyle.ShrinkToFit = cellStyle.ShrinkToFit;
            myCellStyle.Indention = cellStyle.Indention;

            return myCellStyle;
        }
    }
    static public class ExcelClass
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
        public static void SaveFile(this System.Data.DataTable dataTable, string filepath)
        {
            try
            {
                SaveFile(dataTable, filepath, new int[] { });
            }
            catch
            {
                Console.WriteLine("找不到Excel安裝資訊!");
            }
    
        }
        public static void SaveFile(this System.Data.DataTable dataTable, string filepath ,int[] ColunmWidth )
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

        public static void NPOI_SaveFile(this System.Data.DataTable dt, string filepath)
        {
            NPOI.SS.UserModel.IWorkbook workbook;
            string fileExt = Path.GetExtension(filepath).ToLower();
            if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return; }
            NPOI.SS.UserModel.ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName);

            //表头  
            NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                NPOI.SS.UserModel.ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    NPOI.SS.UserModel.ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件  
            using (FileStream fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }


        }

        public static void NPOI_SaveFile(this string json, string file)
        {
            Basic.Time.MyTimerBasic myTimerBasic = new Time.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();
            SheetClass sheetClass = json.JsonDeserializet<SheetClass>();
            if (sheetClass == null) return;

            NPOI.SS.UserModel.IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return; }
            NPOI.SS.UserModel.ISheet sheet = string.IsNullOrEmpty("Sheet1") ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet("Sheet1");
            for (int i = 0; i < sheetClass.CellValues.Count; i++)
            {
                CellValue cellValue = sheetClass.CellValues[i];
            
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(cellValue.RowStart, cellValue.RowEnd, cellValue.ColStart, cellValue.ColEnd));
                if (sheet.GetRow(cellValue.RowStart) == null) sheet.CreateRow(cellValue.RowStart);
                if (sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart) == null) sheet.GetRow(cellValue.RowStart).CreateCell(cellValue.ColStart);
              
            }
            for (int i = 0; i < sheetClass.CellValues.Count; i++)
            {
                CellValue cellValue = sheetClass.CellValues[i];
                sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart).SetCellValue(cellValue.Text);
                sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart).CellStyle.SetFont(MyFont.ToIFont(workbook, cellValue.Font));
                sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart).CellStyle.Alignment = cellValue.CellStyle.Alignment;
                sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart).CellStyle.VerticalAlignment = cellValue.CellStyle.VerticalAlignment;

            }

            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件  
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
 
            Console.WriteLine($"存檔耗時{myTimerBasic.ToString()}");
        }
        public static string NPOI_LoadToJson(this string file)
        {
            Basic.Time.MyTimerBasic myTimerBasic = new Time.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();

            string result = "";
            NPOI.SS.UserModel.IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            try
            {
                FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read);
                if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
                SheetClass sheetClass = new SheetClass();
                for (int r = 0; r <= sheet.LastRowNum; r++)
                {

                    for (int c = 0; c < sheet.GetRow(r).LastCellNum; c++)
                    {
                        CellValue cellValue = new CellValue();
                        ICell cell = sheet.GetRow(r).GetCell(c);
                        object obj = NPOI_GetValueType(cell);
                        if (obj != null)
                        {
                            cellValue.Text = obj.ObjectToString();
                        }

                        if (!sheet.NPOI_IsMergeCell(r, c, ref cellValue))
                        {
                            cellValue.RowStart = r;
                            cellValue.RowEnd = r;
                            cellValue.ColStart = c;
                            cellValue.ColEnd = c;
                        }

                        cellValue.Font = MyFont.ToMyFont(cell.CellStyle.GetFont(workbook));
                        cellValue.CellStyle = MyCellStyle.ToMyCellStyle(cell.CellStyle);
                        sheetClass.Add(cellValue);
                       
                    }
                }
                result = sheetClass.JsonSerializationt(false);
                //Console.WriteLine($"{result}");
                fs.Close();
                fs.Dispose();
                workbook.Close();
                Console.WriteLine($"讀檔耗時{myTimerBasic.ToString()}");
            }
            catch
            {
                Console.WriteLine($"NPOI_LoadHeader 檔案已開啟!無法讀取! , 位置 : {file}");
                return "[]";
            }
            finally
            {

            }
 
    
            return result;
        }
        public static DataTable NPOI_LoadFile(this string file)
        {
            try
            {
                DataTable dt = new DataTable();
                NPOI.SS.UserModel.IWorkbook workbook;
                string fileExt = Path.GetExtension(file).ToLower();
                using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                    if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs); } else { workbook = null; }
                    if (workbook == null) { return null; }
                    NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
                    //表头
                    NPOI.SS.UserModel.IRow header = sheet.GetRow(sheet.FirstRowNum);
                    List<int> columns = new List<int>();
                    for (int i = 0; i < header.LastCellNum; i++)
                    {
                        object obj = NPOI_GetValueType(header.GetCell(i));
                        if (obj == null || obj.ToString() == string.Empty)
                        {
                            dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        }
                        else
                        {
                            dt.Columns.Add(new DataColumn(obj.ToString()));
                        }
                        columns.Add(i);
                    }
                    //数据
                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                    {
                        DataRow dr = dt.NewRow();
                        bool hasValue = false;
                        foreach (int j in columns)
                        {
                            dr[j] = NPOI_GetValueType(sheet.GetRow(i).GetCell(j));
                            if (dr[j] != null && dr[j].ToString() != string.Empty)
                            {
                                hasValue = true;
                            }
                        }
                        if (hasValue)
                        {
                            dt.Rows.Add(dr);
                        }
                    }
                }
                return dt;
            }
            catch
            {
                return null;
            }        
        }
        private static object NPOI_GetValueType(this NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case NPOI.SS.UserModel.CellType.Blank: //BLANK:  
                    return null;
                case NPOI.SS.UserModel.CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case NPOI.SS.UserModel.CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case NPOI.SS.UserModel.CellType.String: //STRING:  
                    return cell.StringCellValue;
                case NPOI.SS.UserModel.CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case NPOI.SS.UserModel.CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }
        private static bool NPOI_IsMergeCell(this ISheet sheet, int rowIndex, int colIndex, ref CellValue cellValue)
        {
            bool result = false;
            cellValue.RowStart = 0;
            cellValue.RowEnd = 0;
            cellValue.ColStart = 0;
            cellValue.ColEnd = 0;
            if ((rowIndex < 0) || (colIndex < 0)) return result;
            int regionsCount = sheet.NumMergedRegions;
            for (int i = 0; i < regionsCount; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                //sheet.IsMergedRegion(range); 
                if (rowIndex >= range.FirstRow && rowIndex <= range.LastRow && colIndex >= range.FirstColumn && colIndex <= range.LastColumn)
                {
                    cellValue.RowStart = range.FirstRow;
                    cellValue.RowEnd = range.LastRow;
                    cellValue.ColStart = range.FirstColumn;
                    cellValue.ColEnd = range.LastColumn;
                    result = true;
                    break;
                }
            }
            return result;
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
