using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AttendanceExportTool
{
    abstract class ExportExcelSheet
    {
        public abstract void Create(ExcelPackage package, string sheetName);
    }

    class BusinessSheet : ExportExcelSheet
    {
        public override void Create(ExcelPackage package, string sheetName)
        {
            var excelSetting = GlobalDefine.Instance.Config.ExportExcelSetting;

            var sheet = package.Workbook.Worksheets.Add(sheetName);
            for (int i = 1; i <= GlobalDefine.BUSINESS_EXCEL_TITLE.Length; i++)
            {
                var cell = sheet.Cells[1, i];
                cell.Value = GlobalDefine.BUSINESS_EXCEL_TITLE[i - 1];
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(excelSetting.SignTitleBackgroundColor));
                cell.Style.Font.Bold = true;
            }

            int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year , GlobalDefine.Instance.Config.CurrentMonth); 
            for (int i = 0; i < days; i++)
            {
                var cell = sheet.Cells[1, GlobalDefine.BUSINESS_EXCEL_TITLE.Length + i + 1];
                cell.Value = i + 1;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;

                DateTime dt = new DateTime(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth, i + 1);
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(dt.DayOfWeek == DayOfWeek.Monday ? GlobalDefine.MONDAY_COLOR : excelSetting.SignTitleBackgroundColor));
                cell.Style.Font.Bold = true;
            }

            var config = GlobalDefine.Instance.Config;
            int memberCount = AttendanceDataManager.Instance.BusinessMemberNameList.Count;
            int memberIndex = 0;
            foreach (var key in AttendanceDataManager.Instance.BusinessMemberNameList.Keys)
            {
                int col = 1;
                var cell = sheet.Cells[memberIndex + 2, col++];
                cell.Value = AttendanceDataManager.Instance.BusinessMemberNameList[key];
                cell.AddComment("人员编号: " + key, config.ExportExcelSetting.Author);
                cell = sheet.Cells[memberIndex + 2, col++];
                int num = AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.Normal) +
                          AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkLeaveEarly) +
                          AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkLate) +
                          AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkLateAndLeveaEarly) +
                          AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkUnClockIn) +
                          AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkUnClockOff);
                if (num > 0)
                {
                    cell.Value = num;
                }

                cell = sheet.Cells[memberIndex + 2, col++];
                num = AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkRest) + AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkUnClockInAndOff);
                if (num > 0)
                {
                    cell.Value = num;
                }

                cell = sheet.Cells[memberIndex + 2, col++];
                num = 0;
                if (num > 0)
                {
                    cell.Value = num;
                }

                cell = sheet.Cells[memberIndex + 2, col++];
                num = AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkLeaveEarly) +
                      AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkLate) +
                      AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkLateAndLeveaEarly);
                if (num > 0)
                {
                    cell.Value = num;
                }

                cell = sheet.Cells[memberIndex + 2, col];
                num = AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkUnClockIn) +
                      AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkUnClockOff) +
                      AttendanceDataManager.Instance.GetBusinessWorkCount(key, WorkTimeType.WorkUnClockInAndOff);
                if (num > 0)
                {
                    cell.Value = num;
                }

                for (int j = 0; j < days; j++)
                {
                    cell = sheet.Cells[memberIndex + 2, j + GlobalDefine.BUSINESS_EXCEL_TITLE.Length + 1];
                    var workTypeInfo = AttendanceDataManager.Instance.GetBusinessWorkType(key, j + 1);
                    WorkTimeType workTimeType = workTypeInfo.WorkTimeType;
                    cell.Value = GlobalDefine.WORK_TIME_TYPE_STRINGS[(int)workTimeType];
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(GlobalDefine.WORK_TYPE_COLOR[(int)workTimeType]));

                    if (workTimeType == WorkTimeType.WorkUnClockInAndOff)
                    {
                        cell.Style.Font.Color.SetColor(Color.Red);
                    }

                    cell.AddComment(workTypeInfo.GetComment(), config.ExportExcelSetting.Author);
                    cell.Comment.Font.Size = excelSetting.SignTitleFontSize * 0.8f;
                }

                sheet.Row(memberIndex + 2).Height = excelSetting.SignCellHeight;
                memberIndex++;
            }


            for (int i = 0; i < GlobalDefine.WORK_TIME_TYPE_STRINGS.Length; i++)
            {
                var cell = sheet.Cells[memberCount + 3, GlobalDefine.BUSINESS_EXCEL_TITLE.Length + i + 1];
                cell.Value = GlobalDefine.WORK_TIME_TYPE_STRINGS[i];
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(GlobalDefine.WORK_TYPE_COLOR[i]));
                sheet.Row(memberCount + 3).Height = excelSetting.SignCellHeight;

                if (i == (int) WorkTimeType.WorkUnClockInAndOff)
                {
                    cell.Style.Font.Color.SetColor(Color.Red);
                }
            }


            sheet.Row(1).Height = excelSetting.SignCellHeight;

            string startColumn = Util.Util.ToNumberSystem26(1);
            string endColumn = Util.Util.ToNumberSystem26(sheet.Dimension.End.Column);
            string columnRange = $"{startColumn}:{endColumn}";

            sheet.Cells[columnRange].Style.Font.Size = excelSetting.SignTitleFontSize;
            sheet.Cells[columnRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[columnRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells[columnRange].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            sheet.Cells[columnRange].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[columnRange].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            sheet.Cells[columnRange].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            sheet.Cells[columnRange].Style.WrapText = true;

            for (int i = 1; i <= sheet.Dimension.End.Column; i++)
            {
                if (i == 1)
                {
                    sheet.Column(i).AutoFit(8);
                }
                else
                {
                    sheet.Column(i).Width = config.ExportExcelSetting.SignCellWidth;
                }
            }
        }
    }
}
