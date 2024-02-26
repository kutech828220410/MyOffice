
namespace Txt2Excel
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
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.saveFileDialog_SaveExcel = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog_LoadExcel = new System.Windows.Forms.OpenFileDialog();
            this.button_Loadfile = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.sqL_DataGridView_藥品資料 = new SQLUI.SQL_DataGridView();
            this.button_Savefile = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // saveFileDialog_SaveExcel
            // 
            this.saveFileDialog_SaveExcel.DefaultExt = "xls";
            this.saveFileDialog_SaveExcel.Filter = "Excel File (*.xls)|*.xls";
            // 
            // openFileDialog_LoadExcel
            // 
            this.openFileDialog_LoadExcel.DefaultExt = "txt";
            this.openFileDialog_LoadExcel.Filter = "Word File (*.docx)|*.docx|txt File (*.txt)|*.txt;";
            // 
            // button_Loadfile
            // 
            this.button_Loadfile.Location = new System.Drawing.Point(22, 22);
            this.button_Loadfile.Name = "button_Loadfile";
            this.button_Loadfile.Size = new System.Drawing.Size(178, 56);
            this.button_Loadfile.TabIndex = 0;
            this.button_Loadfile.Text = "Loadfile";
            this.button_Loadfile.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button_Savefile);
            this.panel1.Controls.Add(this.button_Loadfile);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 823);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1264, 99);
            this.panel1.TabIndex = 8;
            // 
            // sqL_DataGridView_藥品資料
            // 
            this.sqL_DataGridView_藥品資料.AutoSelectToDeep = true;
            this.sqL_DataGridView_藥品資料.backColor = System.Drawing.Color.Silver;
            this.sqL_DataGridView_藥品資料.BorderColor = System.Drawing.Color.Silver;
            this.sqL_DataGridView_藥品資料.BorderRadius = 0;
            this.sqL_DataGridView_藥品資料.BorderSize = 2;
            this.sqL_DataGridView_藥品資料.cellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Raised;
            this.sqL_DataGridView_藥品資料.cellStylBackColor = System.Drawing.Color.PowderBlue;
            this.sqL_DataGridView_藥品資料.cellStyleFont = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.sqL_DataGridView_藥品資料.cellStylForeColor = System.Drawing.Color.Black;
            this.sqL_DataGridView_藥品資料.columnHeaderBackColor = System.Drawing.Color.DarkGray;
            this.sqL_DataGridView_藥品資料.columnHeaderFont = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.sqL_DataGridView_藥品資料.columnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Raised;
            this.sqL_DataGridView_藥品資料.columnHeadersHeight = 4;
            this.sqL_DataGridView_藥品資料.columnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.sqL_DataGridView_藥品資料.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sqL_DataGridView_藥品資料.Font = new System.Drawing.Font("新細明體", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.sqL_DataGridView_藥品資料.ImageBox = false;
            this.sqL_DataGridView_藥品資料.Location = new System.Drawing.Point(0, 0);
            this.sqL_DataGridView_藥品資料.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.sqL_DataGridView_藥品資料.Name = "sqL_DataGridView_藥品資料";
            this.sqL_DataGridView_藥品資料.OnlineState = SQLUI.SQL_DataGridView.OnlineEnum.Online;
            this.sqL_DataGridView_藥品資料.Password = "user82822040";
            this.sqL_DataGridView_藥品資料.Port = ((uint)(3306u));
            this.sqL_DataGridView_藥品資料.rowHeaderBackColor = System.Drawing.Color.Gray;
            this.sqL_DataGridView_藥品資料.rowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Raised;
            this.sqL_DataGridView_藥品資料.RowsColor = System.Drawing.SystemColors.Control;
            this.sqL_DataGridView_藥品資料.RowsHeight = 80;
            this.sqL_DataGridView_藥品資料.SaveFileName = "SQL_DataGridView";
            this.sqL_DataGridView_藥品資料.Server = "127.0.0.0";
            this.sqL_DataGridView_藥品資料.Size = new System.Drawing.Size(1264, 823);
            this.sqL_DataGridView_藥品資料.SSLMode = MySql.Data.MySqlClient.MySqlSslMode.None;
            this.sqL_DataGridView_藥品資料.TabIndex = 9;
            this.sqL_DataGridView_藥品資料.UserName = "root";
            this.sqL_DataGridView_藥品資料.可拖曳欄位寬度 = false;
            this.sqL_DataGridView_藥品資料.可選擇多列 = false;
            this.sqL_DataGridView_藥品資料.單格樣式 = System.Windows.Forms.DataGridViewCellBorderStyle.Raised;
            this.sqL_DataGridView_藥品資料.自動換行 = true;
            this.sqL_DataGridView_藥品資料.表單字體 = new System.Drawing.Font("新細明體", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.sqL_DataGridView_藥品資料.邊框樣式 = System.Windows.Forms.BorderStyle.Fixed3D;
            this.sqL_DataGridView_藥品資料.顯示CheckBox = false;
            this.sqL_DataGridView_藥品資料.顯示首列 = true;
            this.sqL_DataGridView_藥品資料.顯示首行 = true;
            this.sqL_DataGridView_藥品資料.首列樣式 = System.Windows.Forms.DataGridViewHeaderBorderStyle.Raised;
            this.sqL_DataGridView_藥品資料.首行樣式 = System.Windows.Forms.DataGridViewHeaderBorderStyle.Raised;
            // 
            // button_Savefile
            // 
            this.button_Savefile.Location = new System.Drawing.Point(206, 22);
            this.button_Savefile.Name = "button_Savefile";
            this.button_Savefile.Size = new System.Drawing.Size(178, 56);
            this.button_Savefile.TabIndex = 1;
            this.button_Savefile.Text = "Savefile";
            this.button_Savefile.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1264, 922);
            this.Controls.Add(this.sqL_DataGridView_藥品資料);
            this.Controls.Add(this.panel1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SaveFileDialog saveFileDialog_SaveExcel;
        private System.Windows.Forms.OpenFileDialog openFileDialog_LoadExcel;
        private System.Windows.Forms.Button button_Loadfile;
        private System.Windows.Forms.Panel panel1;
        private SQLUI.SQL_DataGridView sqL_DataGridView_藥品資料;
        private System.Windows.Forms.Button button_Savefile;
    }
}

