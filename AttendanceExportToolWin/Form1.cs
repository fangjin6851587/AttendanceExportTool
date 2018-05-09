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

        private BackgroundWorker bkWorker = new BackgroundWorker();
        private ProgressForm notifyForm = new ProgressForm();

        public AttendanceExportWindow()
        {
            InitializeComponent();

            CheckForIllegalCrossThreadCalls = false;
            bkWorker.WorkerReportsProgress = true;  
            bkWorker.DoWork += new DoWorkEventHandler(DoWork);  
            bkWorker.ProgressChanged += new ProgressChangedEventHandler(ProgessChanged);  
            bkWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CompleteWork);  

        }

        private void CompleteWork(object sender, RunWorkerCompletedEventArgs e)
        {
            notifyForm.Close();
            int result = (int)e.Result;
            if (result == 0)
            {
                MessageBox.Show("导出成功.");
            }
            else if (result == 1)
            {
                MessageBox.Show("表格数据读取失败.");
            }
            else if (result == 2)
            {
                MessageBox.Show("表格数据导出失败.");
            }
        }

        private void ProgessChanged(object sender, ProgressChangedEventArgs e)
        {
            notifyForm.SetNotifyInfo(e.ProgressPercentage, "处理进度:" + Convert.ToString(e.ProgressPercentage) + "%");
        }

        private void DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = ProcessProgress(bkWorker, e);
        }

        private int ProcessProgress(object sender, DoWorkEventArgs e)  
        {  
            List<IInit> initList = new List<IInit>
            {
                AttendanceDataManager.Instance,
                MemberDataManager.Instance,
                OvertimeDataManager.Instance,
                MemberPayDataManager.Instance,
            };

            List<ExcelWriter> excelWriters = new List<ExcelWriter>
            {
                new BusinessExcelExporter(),
                new ShoppingGuideExporter(),
                new AdministrativeExport(),
                new PayForgetMemberExporter()
            };

            int totalProgress = initList.Count + excelWriters.Count;

            for (int i = 0; i < totalProgress; i++)
            {
                if (i < initList.Count)
                {
                    InitCode code = initList[i].Init();
                    if (code != InitCode.Ok)
                    {
                        LogController.Log("Excel export failed code = " + code);
                        return 1;
                    }
                }
                else
                {
                    try
                    {
                        excelWriters[i - initList.Count].Save();
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine(exception);
                        return 2;
                    }
                }

                if (bkWorker.CancellationPending)  
                {  
                    e.Cancel = true;  
                    return -1;  
                }

                bkWorker.ReportProgress((int) ((float)(i + 1) / totalProgress * 100));  
                System.Threading.Thread.Sleep(1);
            }
  
            return 0;  
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

            notifyForm.StartPosition = FormStartPosition.CenterParent;  
            bkWorker.RunWorkerAsync();  
            notifyForm.ShowDialog();  
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
