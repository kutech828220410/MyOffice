using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Form_Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
           // dt = MyOffice.ExcelClass.LoadFile(@"C:\Users\User\Desktop\TEST.xls");
            dataGridView1.DataSource = dt;

            MyOffice.ExcelClass.SaveFile(dt, @"C:\Users\User\Desktop\TEST123");
        }
    }
}
