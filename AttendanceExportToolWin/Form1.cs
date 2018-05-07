using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AttendanceExportTool;
using AttendanceExportTool.Util;

namespace AttendanceExportToolWin
{
    public partial class AttendanceExportWindow : Form
    {
        public AttendanceExportWindow()
        {
            InitializeComponent();

        }

        private void AttendanceExportWindow_Load(object sender, EventArgs e)
        {
            InitCode code = GlobalDefine.Instance.Init();
            if (code != InitCode.Ok)
            {
                if (MessageBox.Show("配置数据初始化失败.", "", MessageBoxButtons.OK) == DialogResult.OK)
                {
                    Close();
                }
                return;
            }

            UpdateCurrentMonth();
        }

        private void CurrentMonthValueChanged(object sender, EventArgs e)
        {
            UpdateCurrentMonth();
        }

        private void UpdateCurrentMonth()
        {
            GlobalDefine.Instance.Config.CurrentMonth = Convert.ToDateTime(CurrentMonthDateTimePiacker.Value).Month;
            LogController.Log("Current month: " + GlobalDefine.Instance.Config.CurrentMonth);
        }

        private void ImportSignCick(object sender, EventArgs e)
        {
            openExcelFileDialog.Title = ImportSignPathTip.Text;
            openExcelFileDialog.Multiselect = false;
            if (openExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openExcelFileDialog.FileName;
                GlobalDefine.Instance.Config.ImportSignPath = openExcelFileDialog.FileName;
            }
        }

        private void ImportMemberCick(object sender, EventArgs e)
        {
            openExcelFileDialog.Title = label1.Text;
            openExcelFileDialog.Multiselect = false;
            if (openExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openExcelFileDialog.FileName;
                GlobalDefine.Instance.Config.ImportMemberPath = openExcelFileDialog.FileName;
            }
        }

        private void OverTimeCick(object sender, EventArgs e)
        {
            openExcelFileDialog.Title = label2.Text;
            openExcelFileDialog.Multiselect = false;
            if (openExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = openExcelFileDialog.FileName;
                GlobalDefine.Instance.Config.OverTimePath = openExcelFileDialog.FileName;
            }
        }

        private void PayCick(object sender, EventArgs e)
        {
            openExcelFileDialog.Title = label3.Text;
            openExcelFileDialog.Multiselect = true;
            if (openExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = String.Join(";" ,openExcelFileDialog.FileNames);
                GlobalDefine.Instance.Config.PayPathList = openExcelFileDialog.FileNames;
            }
        }

        private void StartExportCick(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(GlobalDefine.Instance.Config.ImportSignPath))
            {
                return;
            }

            if (string.IsNullOrEmpty(GlobalDefine.Instance.Config.ImportMemberPath))
            {
                return;
            }

            if (string.IsNullOrEmpty(GlobalDefine.Instance.Config.OverTimePath))
            {
                return;
            }

            if (GlobalDefine.Instance.Config.PayPathList == null)
            {
                return;
            }

            List<IInit> initList = new List<IInit>
            {
                AttendanceDataManager.Instance,
                MemberDataManager.Instance,
                OvertimeDataManager.Instance,
                MemberPayDataManager.Instance,
            };

            foreach (var init in initList)
            {
                InitCode code = init.Init();
                if (code != InitCode.Ok)
                {
                    LogController.Log("Excel export failed code = " + code);
                    MessageBox.Show("表格数据读取失败.");
                    return;
                }
            }

            try
            {
                List<ExcelWriter> excelWriters = new List<ExcelWriter>
                {
                    new BusinessExcelExporter(),
                    new ShoppingGuideExporter(),
                    new AdministrativeExport(),
                };

                foreach (var excelWriter in excelWriters)
                {
                    excelWriter.Save();
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show("导出失败.");
                throw;
            }

            MessageBox.Show("导出成功.");
        }

        private void ExportDirCick(object sender, EventArgs e)
        {
            if (explortBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = explortBrowserDialog.SelectedPath;
                GlobalDefine.Instance.ExportDir = explortBrowserDialog.SelectedPath;
            }
        }
    }
}
