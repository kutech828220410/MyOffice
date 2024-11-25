using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MyOffice;
using Basic;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
namespace Form_Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.button_Json解碼.Click += Button_Json解碼_Click;
            this.button_讀取Excel.Click += Button_讀取Excel_Click;
            this.button_存檔Excel.Click += Button_存檔Excel_Click;
            this.button_API_GET.Click += Button_API_GET_Click;
            this.button_Json解碼_new.Click += Button_Json解碼_new_Click;
        }

        private void Button_Json解碼_new_Click(object sender, EventArgs e)
        {
            List<SheetClass> sheetClasses = this.textBox_Json.Text.JsonDeserializet<List<SheetClass>>();


        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = "";
            if(this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                path = this.folderBrowserDialog1.SelectedPath;
                MyOffice.ExcelClass.NPOI_SaveFiles2Folder(@"C:\Users\Administrator\Desktop\屏榮20241025盤點\盤點表格-列印用OK.xlsx", path);
                MessageBox.Show("完成!");
            }
          
   
            
        }


        private void Button_Json解碼_Click(object sender, EventArgs e)
        {
            SheetClass sheetClass = this.textBox_Json.Text.JsonDeserializet<SheetClass>();
            if(sheetClass == null)
            {
                MessageBox.Show("解碼失敗!");
                return;
            }
            using (Bitmap bitmap = sheetClass.GetBitmap())
            {
                using (Graphics g = panel1.CreateGraphics())
                {
                    g.DrawImage(bitmap, new PointF());
                }
            }
        }
        SheetClass sheetClass;
        private void Button_讀取Excel_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string extension = System.IO.Path.GetExtension(openFileDialog1.FileName);
                if(extension == ".txt")
                {
                    string json = MyFileStream.LoadFileAllText(openFileDialog1.FileName , "big5");
                    List<SheetClass> sheetClasses = MyFileStream.LoadFileAllText(openFileDialog1.FileName).JsonDeserializet<List<SheetClass>>();
                    byte[] excelData = sheetClasses.NPOI_GetBytes(Excel_Type.xlsx);
                }
                else
                {
                    MyOffice.ExcelClass.NPOI_LoadFile(openFileDialog1.FileName);
                    sheetClass = MyOffice.ExcelClass.NPOI_LoadToSheetClass(openFileDialog1.FileName);
                    this.textBox_Json.Text = sheetClass.JsonSerializationt(false);
                }
           
            }
        }
        private void Button_存檔Excel_Click(object sender, EventArgs e)
        {
            if (saveFileDialog_SaveExcel.ShowDialog() == DialogResult.OK)
            {
                sheetClass.NPOI_SaveFile(saveFileDialog_SaveExcel.FileName);
                MessageBox.Show("完成!");
            }
        }
    
        private void Button_API_GET_Click(object sender, EventArgs e)
        {
            string str = Basic.Net.WEBApiGet(@"https://localhost:44318/api/test/excel");
            this.textBox_Json.Text = str;


            List<SheetClass> sheetClass = this.textBox_Json.Text.JsonDeserializet<List<SheetClass>>();
            if (sheetClass == null)
            {
                MessageBox.Show("解碼失敗!");
                return;
            }
            using (Bitmap bitmap = sheetClass[0].GetBitmap())
            {
                using (Graphics g = panel1.CreateGraphics())
                {
                    g.DrawImage(bitmap, new PointF());
                }
            }
        }
    }
}
