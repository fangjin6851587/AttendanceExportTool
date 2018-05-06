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
    class AdministrativeExport : ExcelWriter
    {
        protected override string GetPath()
        {
            return GlobalDefine.Instance.Config.AdministrativeExportPath;
        }

        protected override void Export(ExcelPackage package)
        {
            var excelSetting = GlobalDefine.Instance.Config.ExportExcelSetting;
            var sheet = package.Workbook.Worksheets.Add("sheet1");
            int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth);

            sheet.Cells[1, 1].Value = "日期";
            sheet.Cells["A1:B1"].Merge = true;
            string columnRange = "A:B";
            sheet.Cells[columnRange].Style.Font.Size = excelSetting.SignTitleFontSize;
            sheet.Cells[columnRange].Style.Font.Bold = true;
            sheet.Cells[columnRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[columnRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            for (int i = 1; i <= days; i++)
            {
                var cell = sheet.Cells[1, i + 2];
                cell.Value = i;
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Empty);
            }

            var memberList = AttendanceDataManager.Instance.AdministrativeMemberNameList.Keys;
            int startRow = 2;
            foreach (var id in memberList)
            {
                startRow = CreateMember(sheet, startRow, id);
            }

            sheet.Cells[1, 3 + days].Value = "月季";
            sheet.Cells[1, 4 + days].Value = "年计";
            sheet.Cells[1, 5 + days].Value = "累计年假";
            string startColumn = Util.Util.GetCellAddress(1, 3 + days);
            string endColumn = Util.Util.GetCellAddress(sheet.Dimension.End.Row, sheet.Dimension.End.Column);
            columnRange = $"{startColumn}:{endColumn}";
            sheet.Cells[columnRange].Style.Font.Size = excelSetting.SignTitleFontSize;
            sheet.Cells[columnRange].Style.Font.Bold = true;
            sheet.Cells[columnRange].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Empty);
            sheet.Cells[columnRange].AutoFitColumns(10);

            startColumn = Util.Util.ToNumberSystem26(1);
            endColumn = Util.Util.ToNumberSystem26(sheet.Dimension.End.Column);
            columnRange = $"{startColumn}:{endColumn}";
            sheet.Cells[columnRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[columnRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            startColumn = Util.Util.GetCellAddress(1, 2);
            endColumn = Util.Util.GetCellAddress(sheet.Dimension.End.Row, 2);
            columnRange = $"{startColumn}:{endColumn}";
            sheet.Cells[columnRange].Style.Border.Right.Style = ExcelBorderStyle.Thick;

            startColumn = Util.Util.GetCellAddress(1, 2);
            endColumn = Util.Util.GetCellAddress(sheet.Dimension.End.Row, 2);
            columnRange = $"{startColumn}:{endColumn}";
            sheet.Cells[columnRange].AutoFitColumns();

            for (int i = 1; i <= days; i++)
            {
                startColumn = Util.Util.GetCellAddress(1, i + 2);
                endColumn = Util.Util.GetCellAddress(sheet.Dimension.End.Row, i + 2);
                columnRange = $"{startColumn}:{endColumn}";

                DateTime dt = new DateTime(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth, i);
                if (GlobalDefine.Instance.Config.HolidayTimeList.FirstOrDefault(p => p.IsInHolidayTime(i)) != null ||
                    dt.DayOfWeek == DayOfWeek.Sunday || dt.DayOfWeek == DayOfWeek.Saturday)
                {

                    sheet.Cells[columnRange].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[columnRange].Style.Fill.BackgroundColor.SetColor(Color.Bisque);
                }
                sheet.Cells[columnRange].AutoFitColumns(3);
            }
        }

        private int CreateMember(ExcelWorksheet sheet, int startRow, int id)
        {
            string start = Util.Util.GetCellAddress(startRow, 1);
            string end = Util.Util.GetCellAddress(startRow + 3, 1);
            string columnRange = $"{start}:{end}";
            sheet.Cells[columnRange].Merge = true;
            sheet.Cells[columnRange].Value = AttendanceDataManager.Instance.AdministrativeMemberNameList[id];
            sheet.Cells[columnRange].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Empty);
            sheet.Cells[columnRange].AutoFitColumns(3, 3);
            sheet.Cells[columnRange].Style.WrapText = true;


            sheet.Cells[startRow, 2].Value = "迟到";
            sheet.Cells[startRow, 2].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            sheet.Cells[startRow, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[startRow + 1, 2].Value = "请假";
            sheet.Cells[startRow + 1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[startRow + 2, 2].Value = "加班";
            sheet.Cells[startRow + 2, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[startRow + 3, 2].Value = "备注";

            start = Util.Util.GetCellAddress(startRow + 3, 1);
            end = Util.Util.GetCellAddress(startRow + 3, sheet.Dimension.End.Column);
            columnRange = $"{start}:{end}";
            sheet.Cells[columnRange].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;

            int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth);

            for (int i = 1; i <= days; i++)
            {
                var workTimeType = AttendanceDataManager.Instance.GetAdministrationWorkType(id, i);
                if (workTimeType.WorkTimeType == WorkTimeType.WorkLate ||
                    workTimeType.WorkTimeType == WorkTimeType.WorkLeaveEarly ||
                    workTimeType.WorkTimeType == WorkTimeType.WorkLateAndLeveaEarly)
                {
                    sheet.Cells[startRow, 2 + i].Value = 1;
                }

                DateTime dt = new DateTime(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth, i);

                if (dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday)
                {
                    if (workTimeType.ClockOffTime != null || workTimeType.ClockInTime != null)
                    {
                        sheet.Cells[startRow + 2, 2 + i].Value = 1;
                    }
                }

                if (workTimeType.WorkTimeType == WorkTimeType.WorkUnClockIn ||
                    workTimeType.WorkTimeType == WorkTimeType.WorkUnClockInAndOff ||
                    workTimeType.WorkTimeType == WorkTimeType.WorkUnClockOff)
                {
                    sheet.Cells[startRow + 3, 2 + i].Value = 1;
                }

                sheet.Cells[startRow + 3, 2 + i].AddComment(workTimeType.GetComment(), GlobalDefine.Instance.Config.ExportExcelSetting.Author);
            }

            return startRow + 4;
        }
    }
}
