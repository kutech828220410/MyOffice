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
        }


        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = MyOffice.ExcelClass.LoadFile(@"C:\Users\User\Desktop\TEST.xlsx");
            SheetClass sheetClass1 = dt.NPOI_GetSheetClass();
            sheetClass1.NPOI_SaveFile(@"C:\Users\User\Desktop\TEST1.xlsx");

            //List<SheetClass> sheetClasses = new List<SheetClass>();
            //SheetClass sheetClass = new SheetClass("1");
            //sheetClass.ColumnsWidth.Add(10000);
            //sheetClass.ColumnsWidth.Add(10000);
            //sheetClass.ColumnsWidth.Add(10000);
            //sheetClass.ColumnsWidth.Add(10000);
            //for (int col = 0; col < 4; col++)
            //{
            //    for (int row = 0; row < 8; row++)
            //    {
            //        sheetClass.AddNewCell(row, col, $"A{col}-{row}", new Font("微軟正黑體", 14));
            //    }
            //}
            //for (int i = 0; i < sheetClass.Rows.Count; i++)
            //{
            //    sheetClass.Rows[i].Height = 1000;
            //}
            //sheetClasses.Add(sheetClass);
            //sheetClass = new SheetClass("2");
            //sheetClass.ColumnsWidth.Add(10000);
            //sheetClass.ColumnsWidth.Add(10000);
            //sheetClass.ColumnsWidth.Add(10000);
            //sheetClass.ColumnsWidth.Add(10000);
            //for (int col = 0; col < 4; col++)
            //{
            //    for (int row = 0; row < 8; row++)
            //    {
            //        sheetClass.AddNewCell(row, col, $"B{col}-{row}", new Font("微軟正黑體", 14));
            //    }
            //}
            //for (int i = 0; i < sheetClass.Rows.Count; i++)
            //{
            //    sheetClass.Rows[i].Height = 1000;
            //}
            //sheetClasses.Add(sheetClass);
            //sheetClasses.NPOI_SaveFile(@"C:\Users\User\Desktop\TEST.xlsx");


            //List<SheetClass> sheetClasses_buf = MyOffice.ExcelClass.NPOI_LoadToSheetClasses(@"C:\Users\User\Desktop\TEST.xlsx");
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
                sheetClass = MyOffice.ExcelClass.NPOI_LoadToSheetClass(openFileDialog1.FileName);
                this.textBox_Json.Text = sheetClass.JsonSerializationt(false);
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
