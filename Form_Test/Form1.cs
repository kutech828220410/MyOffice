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
            this.button_API_GET.Click += Button_API_GET_Click;
        }



        private void button1_Click(object sender, EventArgs e)
        {
           // dt = MyOffice.ExcelClass.LoadFile(@"C:\Users\User\Desktop\TEST.xls");

            SheetClass sheetClass = MyOffice.ExcelClass.NPOI_LoadToSheetClass(@"C:\Users\User\Desktop\台北榮總管制結存紀錄表.xls");
            sheetClass.ReplaceCell(1, 1, "TTTTT");
            //sheetClass.AddNewCell(3, 3, 0, 10, "測試字體", new Font("標楷體", 20), NPOI_Color.RED, 800);
            //MyOffice.ExcelClass.NPOI_SaveFile(sheetClass, @"C:\Users\User\Desktop\藥品資料1.xls");
            //Rectangle rectangle = new Rectangle();
            //for (int i = 0; i < 3; i++)
            //{
            //    using (Bitmap bitmap = sheetClass.GetBitmap(i, ref rectangle))
            //    {
            //        using (Graphics g = panel1.CreateGraphics())
            //        {
            //            g.SmoothingMode = SmoothingMode.HighQuality; //使繪圖質量最高，即消除鋸齒
            //            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            //            g.CompositingQuality = CompositingQuality.HighQuality;
            //            g.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;
            //            g.DrawImage(bitmap, rectangle);
            //        }
            //    }
            //}

            using (Bitmap bitmap = sheetClass.GetBitmap(1370,756,0.8, H_Alignment.Center, V_Alignment.Center))
            {
                using (Graphics g = panel1.CreateGraphics())
                {
                    g.DrawImage(bitmap, new PointF());
                }
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
        private void Button_讀取Excel_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SheetClass sheetClass = MyOffice.ExcelClass.NPOI_LoadToSheetClass(openFileDialog1.FileName);
                this.textBox_Json.Text = sheetClass.JsonSerializationt(false);
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
