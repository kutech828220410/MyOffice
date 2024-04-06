using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using MyUI;
using SQLUI;
using Basic;
using MyOffice;
using NPOI.XWPF.UserModel;
using System.IO;
namespace Txt2Excel
{
    public partial class Form1 : Form
    {
        public enum enum_藥品資料
        {
            GUID,
            藥品名稱,
            中文名稱,
            藥品代碼,
            廠牌,
            成分及含量,
            用法及用量,
            適應症,
            副作用,
            注意事項,
            懷孕用等級,

        }
        public enum enum_藥品資料_匯出
        {
            藥品名稱,
            中文名稱,
            藥品代碼,
            廠牌,
            成分及含量,
            用法及用量,
            適應症,
            副作用,
            注意事項,
            懷孕用等級,

        }
        public Form1()
        {
            InitializeComponent();
            MyMessageBox.音效 = false;
            this.Load += Form1_Load;
            this.button_Loadfile.Click += Button_Loadfile_Click;
            this.button_Savefile.Click += Button_Savefile_Click;
        }

  

        private void Form1_Load(object sender, EventArgs e)
        {
            Table table = new Table("");
            table.AddColumnList("GUID", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("藥品名稱", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("中文名稱", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("藥品代碼", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("廠牌", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("成分及含量", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("用法及用量", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("適應症", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("副作用", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("注意事項", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("懷孕用等級", Table.StringType.VARCHAR, 50, Table.IndexType.None);

            this.sqL_DataGridView_藥品資料.Init(table);
            sqL_DataGridView_藥品資料.Set_ColumnVisible(false, new enum_藥品資料().GetEnumNames());
            sqL_DataGridView_藥品資料.Set_ColumnWidth(120, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.藥品名稱);
            sqL_DataGridView_藥品資料.Set_ColumnWidth(120, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.中文名稱);
            sqL_DataGridView_藥品資料.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.藥品代碼);
            sqL_DataGridView_藥品資料.Set_ColumnWidth(150, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.廠牌);
            sqL_DataGridView_藥品資料.Set_ColumnWidth(200, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.成分及含量);
            sqL_DataGridView_藥品資料.Set_ColumnWidth(300, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.用法及用量);
            sqL_DataGridView_藥品資料.Set_ColumnWidth(200, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.適應症);
            sqL_DataGridView_藥品資料.Set_ColumnWidth(300, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.副作用);
            sqL_DataGridView_藥品資料.Set_ColumnWidth(300, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.注意事項);
            sqL_DataGridView_藥品資料.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleLeft, enum_藥品資料.懷孕用等級);

        }

        private void Button_Loadfile_Click(object sender, EventArgs e)
        {
            if(this.openFileDialog_LoadExcel.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog_LoadExcel.FileName;

                string Extension = System.IO.Path.GetExtension(fileName);
                if (Extension == ".txt")
                {
                    string allText = Basic.MyFileStream.LoadFileAllText(fileName);
                    string input = allText;
                    string pattern = @"\r+";
                    string replacement = "\r";
                    string result = Regex.Replace(input, pattern, replacement);

                    Console.WriteLine("Original string: " + input);
                    Console.WriteLine("String after merging \\r: " + result);

                    List<object[]> list_藥品資料 = ExtractDrugs(input, "藥品名稱:");


                    
                }
                if (Extension == ".docx")
                {
                    FileStream fs = new FileStream(fileName, FileMode.Open);
                    XWPFDocument doc = new XWPFDocument(fs);
                    IList<XWPFParagraph> paragraphs = doc.Paragraphs;
                    string allText = "";
                    allText =  ExtractTextFromWord(fileName);
            
                    string input = allText;
                    string pattern = @"\r+";
                    string replacement = "\r";
                    string result = Regex.Replace(input, pattern, replacement);

                    Console.WriteLine("Original string: " + input);
                    Console.WriteLine("String after merging \\r: " + result);

                    List<object[]> list_藥品資料 = ExtractDrugs(input, "藥品名稱:");
                    sqL_DataGridView_藥品資料.RefreshGrid(list_藥品資料);
                }
             


            }

        }
        static string ExtractTextFromWord(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                XWPFDocument doc = new XWPFDocument(fs);
                string text = "";

                // 处理段落
                foreach (XWPFParagraph paragraph in doc.Paragraphs)
                {
                    text += paragraph.Text + Environment.NewLine;
                }

                // 处理表格
                foreach (XWPFTable table in doc.Tables)
                {
                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.GetTableCells())
                        {
                            text += cell.GetText() + "\t";
                        }
                        text += Environment.NewLine;
                    }
                    text += Environment.NewLine;
                }

                return text;
            }
        }
        private void Button_Savefile_Click(object sender, EventArgs e)
        {
            if (this.saveFileDialog_SaveExcel.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog_SaveExcel.FileName;
                DataTable dataTable = sqL_DataGridView_藥品資料.GetDataTable();
                dataTable = dataTable.ReorderTable(new enum_藥品資料_匯出());
                MyOffice.ExcelClass.SaveFile(dataTable, fileName);
                MyMessageBox.ShowDialog("儲存成功!");
            }
        }
        static Regex drugRegex = new Regex(@"藥品名稱:\s*(.*?)\s*\[.*?\]", RegexOptions.Singleline);
        static string RemoveSpacesBetweenChineseCharacters(string input, string str)
        {
            return RemoveSpacesBetweenChineseCharacters(input, str.ToArray());
        }
        static string RemoveSpacesBetweenChineseCharacters(string input, params char[] characters)
        {
            string pattern = string.Join(@"[\t\s]+", characters); // 修改正则表达式模式
            return Regex.Replace(input, pattern, match => match.Value.Replace(" ", "").Replace("\t", ""));
        }
        static List<object[]> ExtractDrugs(string text , string regxStr)
        {
            List<object[]> temp = new List<object[]>();

            drugRegex = new Regex($@"{regxStr}\s*(.*?)\s*\[.*?\]", RegexOptions.Multiline);
            var matches = drugRegex.Matches(text);
            int strart_index = 0;
            int endIndex = 0;
            List<string> list_str = new List<string>();
            text = RemoveSpacesBetweenChineseCharacters(text, "藥品名稱:");
            text = RemoveSpacesBetweenChineseCharacters(text, "中文名稱:");
            text = RemoveSpacesBetweenChineseCharacters(text, "藥品代碼:");
            text = RemoveSpacesBetweenChineseCharacters(text, "廠牌:");
            text = RemoveSpacesBetweenChineseCharacters(text, "處置代碼:");
            text = RemoveSpacesBetweenChineseCharacters(text, "健保價:");
            text = RemoveSpacesBetweenChineseCharacters(text, "成份及含量:");
            text = RemoveSpacesBetweenChineseCharacters(text, "成分及含量:");
            text = RemoveSpacesBetweenChineseCharacters(text, "用法及用量:");
            text = RemoveSpacesBetweenChineseCharacters(text, "適應症");
            text = RemoveSpacesBetweenChineseCharacters(text, "副作用:");
            text = RemoveSpacesBetweenChineseCharacters(text, "適應症:");
            text = RemoveSpacesBetweenChineseCharacters(text, "注意事項:");
            text = RemoveSpacesBetweenChineseCharacters(text, "懷孕用藥級:");
            
            while (true)
            {
                int Index = text.IndexOf($"藥品名稱:", strart_index);
                int next_index = text.IndexOf($"藥品名稱:", Index + 1);
                if (next_index == -1)
                {
                    string str_temp = text.Substring(Index, text.Length - Index);
                
                    str_temp = Regex.Replace(str_temp, @"\r+", "\r");
                    str_temp = Regex.Replace(str_temp, @"\n+", "\n");
                    str_temp = Regex.Replace(str_temp, @"\t+", "\t");
                    str_temp = replace(str_temp, "\r\n", "$");
                    str_temp = replace(str_temp, "\n\r", "$");
                    str_temp = replace(str_temp, "$", "$");
                    str_temp = replace(str_temp, "\t", "$");
                    str_temp = replace(str_temp, ": ", ":$");
                    string pattern = @"\$+";
                    string replacement = "$";
                    str_temp = Regex.Replace(str_temp, pattern, replacement);
                    str_temp = Regex.Replace(str_temp, @":\t+", ":$");
                    str_temp = Regex.Replace(str_temp, @":\s+", ":$");
                    list_str.Add(str_temp);
   
                    break;
                }
                else
                {
               
                    string str_temp = text.Substring(Index, next_index - Index);
                    if (str_temp.Contains("景安"))
                    {
                    }
                    str_temp = Regex.Replace(str_temp, @"\r+", "\r");
                    str_temp = Regex.Replace(str_temp, @"\n+", "\n");
                    str_temp = Regex.Replace(str_temp, @"\t+", "\t");
                    str_temp = replace(str_temp, "\r\n", "$");
                    str_temp = replace(str_temp, "\n\r", "$");
                    str_temp = replace(str_temp, "$", "$");
                    str_temp = replace(str_temp, "\t", "$");
                    str_temp = replace(str_temp, ": ", ":$");
                    string pattern = @"\$+";
                    string replacement = "$";
                    str_temp = Regex.Replace(str_temp, pattern, replacement);
                    str_temp = Regex.Replace(str_temp, @":[\t\s]+", ":$");

                    list_str.Add(str_temp);
                    strart_index = next_index;
                }
            }

            foreach (string str in list_str)
            {
      
                string GUID = Guid.NewGuid().ToString();
                string 藥品名稱 = GetValue(str, "藥品名稱:");
                string 中文名稱 = GetValue(str, "中文名稱:"); 
                string 藥品代碼 = GetValue(str, "藥品代碼:");
                string 廠牌 = GetValue(str, "廠牌:");
                string 成份及含量 = GetValue(str, "成份及含量:");
                if(成份及含量.StringIsEmpty()) 成份及含量 = GetValue(str, "成分及含量:");
                string 用法及用量 = GetValue(str, "用法及用量:");
                string 適應症 = GetValue(str, "適應症:");
                string 副作用 = GetValue(str, "副作用:");
                string 注意事項 = GetValue(str, "注意事項:");
                string 懷孕用藥級 = GetValue(str, "懷孕用藥級:");

                temp.Add(new object[] { GUID, 藥品名稱, 中文名稱, 藥品代碼 , 廠牌 , 成份及含量, 用法及用量, 適應症 , 副作用, 注意事項, 懷孕用藥級 });
            }
            return temp;
        }
        static string GetValue(string text ,string title)
        {
            string result = "";
            int index = text.IndexOf(title);
            string temp = "";
   
            int index_ex = text.IndexOf(":$", index + title.Length);
            int Each_index = text.IndexOf(":$", index + title.Length);
            if(Each_index >= 0 && title == "成份及含量:" && title == "成分及含量:")
            {
                index_ex = text.IndexOf(":$", Each_index + ":$".Length);
            }
            if(title == "懷孕用藥級:")
            {
                index_ex = text.IndexOf("$", index + title.Length + 1);
            }
            if (index_ex >= 0 && index >= 0)
            {
                temp = text.Substring(index, index_ex - index);
             
                temp = temp.Replace("藥品名稱", "");
                temp = temp.Replace("處 置 代 碼", "");
                temp = temp.Replace("處置代碼", "");
                temp = temp.Replace("單    位", "");
                temp = temp.Replace("單位", "");
                temp = temp.Replace("中文名稱", "");
                temp = temp.Replace("藥品代碼", "");
                temp = temp.Replace("廠    牌", "");
                temp = temp.Replace("廠牌", "");
                temp = temp.Replace("健保價", "");
                temp = temp.Replace("成份及含量", "");
                temp = temp.Replace("成分及含量", "");
                temp = temp.Replace("用法及用量", "");
                temp = temp.Replace("適應症", "");
                temp = temp.Replace("副作用", "");
                temp = temp.Replace("注意事項", "");
                temp = temp.Replace("懷孕用藥級", "");
                temp = temp.Replace(":$", "");
                if (temp.Length < 1)
                {
                    return "";
                }
                if (temp.Substring(0, 1) == "$")
                {
                    temp = temp.Remove(0, 1);
                }
                temp = temp.Replace("$", " ");
                result = temp.Trim();
   
            }
            return result;
        }

        static string replace(string input , string serch_str , string replace_str)
        {
            string pattern = $@"{serch_str}+";
            string replacement = $"{replace_str}";
            string result = Regex.Replace(input, pattern, replacement);
            return result;
        }
        static string Extract(string text, int startIndex, string serch_start_text, string serch_end_text , ref int endIndex)
        {
            int Index = text.IndexOf($"{serch_start_text}", startIndex);
            if (Index == -1)
                return "";

            int EndIndex = text.IndexOf($"{serch_end_text}", startIndex);
            if (EndIndex == -1)
                return "";
            if (EndIndex - (Index + $"{serch_start_text}:".Length) <= 0)
            {
                return "";
            }
            string Extracttext = text.Substring(Index + $"{serch_start_text}:".Length, EndIndex - (Index + $"{serch_start_text}:".Length));
            endIndex = EndIndex;
            Extracttext = Extracttext.Trim();
            return Extracttext;
        }
        static string Extract(string text, int startIndex, string serch_start_text ,int len)
        {
            int Index = text.IndexOf($"{serch_start_text}", startIndex);
            if (Index == -1)
                return "";

            string Extracttext = text.Substring(Index, len);
            Extracttext = Extracttext.Trim();
            return Extracttext;
        }
    }
}
