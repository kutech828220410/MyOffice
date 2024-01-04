namespace Form_Test
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBox_Json = new System.Windows.Forms.TextBox();
            this.button_Json解碼 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button_讀取Excel = new System.Windows.Forms.Button();
            this.button_API_GET = new System.Windows.Forms.Button();
            this.comboBox_ExcelType = new System.Windows.Forms.ComboBox();
            this.saveFileDialog_SaveExcel = new System.Windows.Forms.SaveFileDialog();
            this.button_存檔Excel = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1388, 225);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 134);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1370, 756);
            this.panel1.TabIndex = 2;
            // 
            // textBox_Json
            // 
            this.textBox_Json.Location = new System.Drawing.Point(734, 774);
            this.textBox_Json.Multiline = true;
            this.textBox_Json.Name = "textBox_Json";
            this.textBox_Json.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox_Json.Size = new System.Drawing.Size(680, 263);
            this.textBox_Json.TabIndex = 3;
            // 
            // button_Json解碼
            // 
            this.button_Json解碼.Location = new System.Drawing.Point(1420, 888);
            this.button_Json解碼.Name = "button_Json解碼";
            this.button_Json解碼.Size = new System.Drawing.Size(100, 134);
            this.button_Json解碼.TabIndex = 4;
            this.button_Json解碼.Text = "Json解碼";
            this.button_Json解碼.UseVisualStyleBackColor = true;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button_讀取Excel
            // 
            this.button_讀取Excel.Location = new System.Drawing.Point(1420, 774);
            this.button_讀取Excel.Name = "button_讀取Excel";
            this.button_讀取Excel.Size = new System.Drawing.Size(100, 108);
            this.button_讀取Excel.TabIndex = 7;
            this.button_讀取Excel.Text = "讀取Excel";
            this.button_讀取Excel.UseVisualStyleBackColor = true;
            // 
            // button_API_GET
            // 
            this.button_API_GET.Location = new System.Drawing.Point(81, 774);
            this.button_API_GET.Name = "button_API_GET";
            this.button_API_GET.Size = new System.Drawing.Size(100, 134);
            this.button_API_GET.TabIndex = 8;
            this.button_API_GET.Text = "API(GET)";
            this.button_API_GET.UseVisualStyleBackColor = true;
            // 
            // comboBox_ExcelType
            // 
            this.comboBox_ExcelType.FormattingEnabled = true;
            this.comboBox_ExcelType.Items.AddRange(new object[] {
            "xls",
            "xlsx"});
            this.comboBox_ExcelType.Location = new System.Drawing.Point(1526, 888);
            this.comboBox_ExcelType.Name = "comboBox_ExcelType";
            this.comboBox_ExcelType.Size = new System.Drawing.Size(121, 20);
            this.comboBox_ExcelType.TabIndex = 9;
            // 
            // saveFileDialog_SaveExcel
            // 
            this.saveFileDialog_SaveExcel.DefaultExt = "txt";
            this.saveFileDialog_SaveExcel.Filter = "Excel File (*.xlsx)|*.xlsx|Excel File (*.xls)|*.xls|txt File (*.txt)|*.txt;";
            // 
            // button_存檔Excel
            // 
            this.button_存檔Excel.Location = new System.Drawing.Point(1420, 660);
            this.button_存檔Excel.Name = "button_存檔Excel";
            this.button_存檔Excel.Size = new System.Drawing.Size(100, 108);
            this.button_存檔Excel.TabIndex = 10;
            this.button_存檔Excel.Text = "存檔Excel";
            this.button_存檔Excel.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1751, 1061);
            this.Controls.Add(this.button_存檔Excel);
            this.Controls.Add(this.comboBox_ExcelType);
            this.Controls.Add(this.button_API_GET);
            this.Controls.Add(this.button_讀取Excel);
            this.Controls.Add(this.button_Json解碼);
            this.Controls.Add(this.textBox_Json);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox textBox_Json;
        private System.Windows.Forms.Button button_Json解碼;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button_讀取Excel;
        private System.Windows.Forms.Button button_API_GET;
        private System.Windows.Forms.ComboBox comboBox_ExcelType;
        private System.Windows.Forms.SaveFileDialog saveFileDialog_SaveExcel;
        private System.Windows.Forms.Button button_存檔Excel;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}

