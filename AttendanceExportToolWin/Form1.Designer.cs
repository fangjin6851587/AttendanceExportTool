using System.Windows.Forms;

namespace AttendanceExportToolWin
{
    partial class AttendanceExportWindow
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.DateMonthTitle = new System.Windows.Forms.Label();
            this.CurrentMonthDateTimePiacker = new System.Windows.Forms.DateTimePicker();
            this.ImportSignPathTip = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.ImportSignButton = new System.Windows.Forms.Button();
            this.ImportMemberButton = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.OverTimeButton = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.PayPathButton = new System.Windows.Forms.Button();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.ExportButton = new System.Windows.Forms.Button();
            this.ExportDirButton = new System.Windows.Forms.Button();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.openExcelFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.explortBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // DateMonthTitle
            // 
            this.DateMonthTitle.AutoSize = true;
            this.DateMonthTitle.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.DateMonthTitle.Location = new System.Drawing.Point(21, 37);
            this.DateMonthTitle.Name = "DateMonthTitle";
            this.DateMonthTitle.Size = new System.Drawing.Size(67, 14);
            this.DateMonthTitle.TabIndex = 1;
            this.DateMonthTitle.Text = "选择月份";
            // 
            // CurrentMonthDateTimePiacker
            // 
            this.CurrentMonthDateTimePiacker.CustomFormat = "yyy-MM";
            this.CurrentMonthDateTimePiacker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.CurrentMonthDateTimePiacker.Location = new System.Drawing.Point(25, 54);
            this.CurrentMonthDateTimePiacker.Name = "CurrentMonthDateTimePiacker";
            this.CurrentMonthDateTimePiacker.Size = new System.Drawing.Size(200, 21);
            this.CurrentMonthDateTimePiacker.TabIndex = 2;
            this.CurrentMonthDateTimePiacker.ValueChanged += new System.EventHandler(this.CurrentMonthValueChanged);
            // 
            // ImportSignPathTip
            // 
            this.ImportSignPathTip.AutoSize = true;
            this.ImportSignPathTip.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ImportSignPathTip.Location = new System.Drawing.Point(22, 95);
            this.ImportSignPathTip.Name = "ImportSignPathTip";
            this.ImportSignPathTip.Size = new System.Drawing.Size(142, 14);
            this.ImportSignPathTip.TabIndex = 3;
            this.ImportSignPathTip.Text = "万信达导出数据表格";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(24, 112);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(340, 21);
            this.textBox1.TabIndex = 4;
            // 
            // ImportSignButton
            // 
            this.ImportSignButton.Location = new System.Drawing.Point(370, 110);
            this.ImportSignButton.Name = "ImportSignButton";
            this.ImportSignButton.Size = new System.Drawing.Size(75, 23);
            this.ImportSignButton.TabIndex = 5;
            this.ImportSignButton.Text = "选择";
            this.ImportSignButton.UseVisualStyleBackColor = true;
            this.ImportSignButton.Click += new System.EventHandler(this.ImportSignCick);
            // 
            // ImportMemberButton
            // 
            this.ImportMemberButton.Location = new System.Drawing.Point(371, 158);
            this.ImportMemberButton.Name = "ImportMemberButton";
            this.ImportMemberButton.Size = new System.Drawing.Size(75, 23);
            this.ImportMemberButton.TabIndex = 8;
            this.ImportMemberButton.Text = "选择";
            this.ImportMemberButton.UseVisualStyleBackColor = true;
            this.ImportMemberButton.Click += new System.EventHandler(this.ImportMemberCick);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(25, 158);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(340, 21);
            this.textBox2.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(23, 143);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 14);
            this.label1.TabIndex = 6;
            this.label1.Text = "人事档案";
            // 
            // OverTimeButton
            // 
            this.OverTimeButton.Location = new System.Drawing.Point(370, 207);
            this.OverTimeButton.Name = "OverTimeButton";
            this.OverTimeButton.Size = new System.Drawing.Size(75, 23);
            this.OverTimeButton.TabIndex = 11;
            this.OverTimeButton.Text = "选择";
            this.OverTimeButton.UseVisualStyleBackColor = true;
            this.OverTimeButton.Click += new System.EventHandler(this.OverTimeCick);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(25, 209);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(340, 21);
            this.textBox3.TabIndex = 10;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(22, 192);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 14);
            this.label2.TabIndex = 9;
            this.label2.Text = "加班记录";
            // 
            // PayPathButton
            // 
            this.PayPathButton.Location = new System.Drawing.Point(372, 263);
            this.PayPathButton.Name = "PayPathButton";
            this.PayPathButton.Size = new System.Drawing.Size(75, 23);
            this.PayPathButton.TabIndex = 14;
            this.PayPathButton.Text = "选择";
            this.PayPathButton.UseVisualStyleBackColor = true;
            this.PayPathButton.Click += new System.EventHandler(this.PayCick);
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(26, 265);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(340, 21);
            this.textBox4.TabIndex = 13;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(24, 246);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 14);
            this.label3.TabIndex = 12;
            this.label3.Text = "工资信息表";
            // 
            // ExportButton
            // 
            this.ExportButton.Location = new System.Drawing.Point(26, 367);
            this.ExportButton.Name = "ExportButton";
            this.ExportButton.Size = new System.Drawing.Size(75, 23);
            this.ExportButton.TabIndex = 15;
            this.ExportButton.Text = "确认";
            this.ExportButton.UseVisualStyleBackColor = true;
            this.ExportButton.Click += new System.EventHandler(this.StartExportCick);
            // 
            // ExportDirButton
            // 
            this.ExportDirButton.Location = new System.Drawing.Point(372, 317);
            this.ExportDirButton.Name = "ExportDirButton";
            this.ExportDirButton.Size = new System.Drawing.Size(75, 23);
            this.ExportDirButton.TabIndex = 18;
            this.ExportDirButton.Text = "选择";
            this.ExportDirButton.UseVisualStyleBackColor = true;
            this.ExportDirButton.Click += new System.EventHandler(this.ExportDirCick);
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(26, 319);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(340, 21);
            this.textBox5.TabIndex = 17;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(24, 300);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(67, 14);
            this.label4.TabIndex = 16;
            this.label4.Text = "导出目录";
            // 
            // openExcelFileDialog
            // 
            this.openExcelFileDialog.Filter = "表格文件 (*.xlsx)|*.xlsx";
            this.openExcelFileDialog.Title = "表格选择";
            this.openExcelFileDialog.CheckFileExists = true;
            // 
            // AttendanceExportWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(692, 430);
            this.Controls.Add(this.ExportDirButton);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.ExportButton);
            this.Controls.Add(this.PayPathButton);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.OverTimeButton);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ImportMemberButton);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ImportSignButton);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.ImportSignPathTip);
            this.Controls.Add(this.CurrentMonthDateTimePiacker);
            this.Controls.Add(this.DateMonthTitle);
            this.Name = "AttendanceExportWindow";
            this.Text = "考勤表格导出工具";
            this.Load += new System.EventHandler(this.AttendanceExportWindow_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        public System.Windows.Forms.Label DateMonthTitle;
        private System.Windows.Forms.DateTimePicker CurrentMonthDateTimePiacker;
        private Label ImportSignPathTip;
        private TextBox textBox1;
        private Button ImportSignButton;
        private Button ImportMemberButton;
        private TextBox textBox2;
        private Label label1;
        private Button OverTimeButton;
        private TextBox textBox3;
        private Label label2;
        private Button PayPathButton;
        private TextBox textBox4;
        private Label label3;
        private Button ExportButton;
        private Button ExportDirButton;
        private TextBox textBox5;
        private Label label4;
        private OpenFileDialog openExcelFileDialog;
        private FolderBrowserDialog explortBrowserDialog;
    }
}

