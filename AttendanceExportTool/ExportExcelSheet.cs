using System;
using System.Collections.Generic;
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

    class SignSheet : ExportExcelSheet
    {
        public override void Create(ExcelPackage package, string sheetName)
        {
            var excelSetting = GlobalDefine.Instance.Config.ExportExcelSetting;

            var sheet = package.Workbook.Worksheets.Add(sheetName);
            for (int i = 1; i < GlobalDefine.SignExcelTitle.Length + 1; i++)
            {
                var cell = sheet.Cells[1, i];
                cell.Value = GlobalDefine.SignExcelTitle[i - 1];
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(excelSetting.SignTitleBackgroundColor));
                cell.Style.Font.Bold = true;
            }

            int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year , GlobalDefine.Instance.Config.CurrentMonth); 
            for (int i = 0; i <= days; i++)
            {
                var cell = sheet.Cells[1, GlobalDefine.SignExcelTitle.Length + i];
                cell.Value = i + 1;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(excelSetting.SignTitleBackgroundColor));
                cell.Style.Font.Bold = true;
            }

            var config = GlobalDefine.Instance.Config;
            for (int i = 0; i < config.MemberInfoList.Length; i++)
            {
                var member = config.MemberInfoList[i];

                int col = 1;
                var cell = sheet.Cells[i + 2, col++];
                cell.Value = member.Name;
                cell = sheet.Cells[i + 2, col++];
                int num = AttendanceDataManager.Instance.GetWorkCount(member.Name, WorkType.Normal);
                if (num > 0)
                {
                    cell.Value = num;
                }

                cell = sheet.Cells[i + 2, col++];
                num = AttendanceDataManager.Instance.GetWorkCount(member.Name, WorkType.WorkRest);
                if (num > 0)
                {
                    cell.Value = num;
                }

                cell = sheet.Cells[i + 2, col++];
                num = 0;
                if (num > 0)
                {
                    cell.Value = num;
                }

                cell = sheet.Cells[i + 2, col++];
                num = AttendanceDataManager.Instance.GetWorkCount(member.Name, WorkType.WorkLeaveEarly) +
                      AttendanceDataManager.Instance.GetWorkCount(member.Name, WorkType.WorkLate) + 
                      AttendanceDataManager.Instance.GetWorkCount(member.Name, WorkType.WorkLateAndLeveaEarly);
                if (num > 0)
                {
                    cell.Value = num;
                }

                cell = sheet.Cells[i + 2, col];
                num = AttendanceDataManager.Instance.GetWorkCount(member.Name, WorkType.WorkUnClockIn) +
                      AttendanceDataManager.Instance.GetWorkCount(member.Name, WorkType.WorkUnClockOff) + 
                      AttendanceDataManager.Instance.GetWorkCount(member.Name, WorkType.WorkUnClockInAndOff);
                if (num > 0)
                {
                    cell.Value = num;
                }

                for (int j = 0; j < days; j++)
                {
                    cell =  sheet.Cells[i + 2, j + GlobalDefine.SignExcelTitle.Length];
                    var workType = AttendanceDataManager.Instance.GetWorkType(member.Name, j + 1);
                    cell.Value = GlobalDefine.WorkTypeStrings[(int) workType.WorkType];
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(GlobalDefine.WorkTypeColor[(int)workType.WorkType]));
                    cell.AddComment(workType.GetComment(), GlobalDefine.Instance.Config.ExportExcelSetting.Author);
                }

                sheet.Row(i + 2).Height = excelSetting.SignCellHeight;
            }

            for (int i = 0; i < GlobalDefine.WorkTypeStrings.Length; i++)
            {
                var cell = sheet.Cells[config.MemberInfoList.Length + 3, GlobalDefine.SignExcelTitle.Length + i];
                cell.Value = GlobalDefine.WorkTypeStrings[i];
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(GlobalDefine.WorkTypeColor[i]));
                sheet.Row(config.MemberInfoList.Length + 3).Height = excelSetting.SignCellHeight;
            }


            sheet.Row(1).Height = excelSetting.SignCellHeight;
            sheet.Cells["A1:AJ"].Style.Font.Size = excelSetting.SignTitleFontSize;
            sheet.Cells["A1:AJ"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["A1:AJ"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells["A1:AJ"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            sheet.Cells["A1:AJ"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells["A1:AJ"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            sheet.Cells["A1:AJ"].Style.Border.Left.Style = ExcelBorderStyle.Thin;


            for (int i = 1; i < sheet.Dimension.End.Column; i++)
            {
                sheet.Column(i).AutoFit();
            }
        }
    }
}
