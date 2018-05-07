using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AttendanceExportTool
{
    class ShoppingGuideExporter : ExcelWriter
    {
        class ShoppingSheet : ExportExcelSheet
        {
            private const string COMMENT_1 = "【本月共休息{0}天】{1}\n【本月共加班{2}天】{3}\n【漏报】{4}上午未报、{5}下午未报\n【本月共报岗{6}天】{7}号开始有报岗记录，{8}号开始无报岗记录";

            public override void Create(ExcelPackage package, string sheetName)
            {
                var excelSetting = GlobalDefine.Instance.Config.ExportExcelSetting;
                var sheet = package.Workbook.Worksheets.Add(sheetName);

                for (int i = 1; i <= GlobalDefine.SHOPPING_EXCEL_TITLES.Length; i++)
                {
                    var cell = sheet.Cells[1, i];
                    cell[1, i].Value = GlobalDefine.SHOPPING_EXCEL_TITLES[i - 1];
                    cell.Style.Font.Size = GlobalDefine.Instance.Config.ExportExcelSetting.SignTitleFontSize;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(excelSetting.SignTitleBackgroundColor));
                    cell.Style.Font.Bold = true;
                }

                var shoppingMemberList = OvertimeDataManager.Instance.GetOvertimeShoppingNameList();
                int index = 0;
                foreach (var payData in MemberPayDataManager.Instance.GetMemberPayDataList())
                {
                    if (string.IsNullOrEmpty(payData.Name))
                    {
                        continue;
                    }

                    shoppingMemberList.TryGetValue(payData.Name, out var member);
                    var memberInfo = MemberDataManager.Instance.GetShoppingGuideMemberInfo(payData.Name);

                    int col = 1;
                    int row = index + 2;
                    var cell = sheet.Cells[row, col++];
                    cell.Value = payData.Name;
                    if (memberInfo != null)
                    {
                        cell.AddComment("人员编号: " + memberInfo.Id, excelSetting.Author);
                    }

                    cell = sheet.Cells[row, col++];
                    cell.Value = payData.WorkType;


                    cell = sheet.Cells[row, col++];

                    if (memberInfo != null)
                    {
                        cell.Value = memberInfo.JoinTime;
                    }
                    cell = sheet.Cells[row, col++];
                    if (memberInfo != null)
                    {
                        cell.Value = memberInfo.ExitTime;
                    }

                    cell = sheet.Cells[row, col++];

                    if (member != null)
                    {
                        cell.Value = member.GetOvertime().ToString("F1");
                    }

                    List<int> unClockTimeList = AttendanceDataManager.Instance.GetUnClockTimeList(payData.Name);
                    
                    int restDays = 0;
                    for (int i = 0; i < GlobalDefine.WORK_TYPE_STRINGS.Length; i++)
                    {
                        if (GlobalDefine.WORK_TYPE_STRINGS[i] == payData.WorkType)
                        {
                            restDays = GlobalDefine.WORK_TYPE_DAY[i];
                        }
                    }

                    cell = sheet.Cells[row, col++];
                    int leaveDay = unClockTimeList.Count - restDays;
                    if (leaveDay > 0)
                    {
                        cell.Value = leaveDay;
                    }

                    cell = sheet.Cells[row, col];
                    var unClockInList = AttendanceDataManager.Instance.GetUnClockInTimeList(payData.Name);
                    var unClockOffList = AttendanceDataManager.Instance.GetUnClockOffTimeList(payData.Name);
                    var signTimeList = AttendanceDataManager.Instance.GetSignTimeList(payData.Name);

                    int startSign = 0;
                    if (signTimeList.Count > 0)
                    {
                        startSign = signTimeList[0];
                    }

                    int endSign = 0;
                    if (signTimeList.Count > 0)
                    {
                        endSign = signTimeList[signTimeList.Count - 1] + 1;
                        int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth);
                        if (endSign > days)
                        {
                            endSign = 0;
                        }
                    }

                    int overTimeCount = 0;
                    string overTimeStr = string.Empty;
                    if (member != null)
                    {
                        overTimeCount = member.GetOvertimeList().Count;
                        overTimeStr = member.GetOvertimeDateTimeString();
                    }

                    string unClockInListString = "0";
                    if (unClockInList.Count > 0)
                    {
                        unClockInListString = String.Join(",", unClockInList.Select(p => p.SignTime.Day));
                    }

                    string unClockOffListString = "0";
                    if (unClockOffList.Count > 0)
                    {
                        unClockOffListString = String.Join(",", unClockOffList.Select(p => p.SignTime.Day));
                    }

                    cell.Value = string.Format(COMMENT_1, unClockTimeList.Count, String.Join(",", unClockTimeList), overTimeCount,
                        overTimeStr, unClockInListString, unClockOffListString, signTimeList.Count, startSign, endSign);

                    string comment = "上午漏报: \n" +
                                     String.Join(",", unClockInList.Select(p => p.ToShoppingGuideCommentString()));

                    comment += "下午漏报: \n" +
                               String.Join(",", unClockOffList.Select(p => p.ToShoppingGuideCommentString()));

                    cell.AddComment(comment, excelSetting.Author);

                    index++;
                }


                for (int i = 1; i <= sheet.Dimension.End.Row; i++)
                {
                    var row = sheet.Row(i);
                    row.Height = 120;
                }

                string startColumn = Util.Util.ToNumberSystem26(1);
                string endColumn = Util.Util.ToNumberSystem26(sheet.Dimension.End.Column -1);
                string columnRange = $"{startColumn}:{endColumn}";

                sheet.Cells[columnRange].Style.Font.Size = excelSetting.SignTitleFontSize;
                sheet.Cells[columnRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[columnRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[columnRange].Style.WrapText = true;
                sheet.Cells[columnRange].AutoFitColumns(15);

                startColumn = Util.Util.ToNumberSystem26(sheet.Dimension.End.Column);
                endColumn = Util.Util.ToNumberSystem26(sheet.Dimension.End.Column);
                columnRange = $"{startColumn}:{endColumn}";
                sheet.Cells[columnRange].Style.Font.Size = excelSetting.SignTitleFontSize;
                sheet.Cells[columnRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells[columnRange].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                sheet.Cells[columnRange].Style.WrapText = true;
                sheet.Cells[columnRange].AutoFitColumns(65);

                startColumn = Util.Util.ToNumberSystem26(1);
                endColumn = Util.Util.ToNumberSystem26(sheet.Dimension.End.Column);
                columnRange = $"{startColumn}:{endColumn}";
                sheet.Cells[columnRange].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[columnRange].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[columnRange].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[columnRange].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            }
        }


        protected override string GetPath()
        {
            return GlobalDefine.Instance.ExportDir + "/" + GlobalDefine.Instance.Config.ShoppingGuideExportPath;
        }

        protected override void Export(ExcelPackage package)
        {
            new ShoppingSheet().Create(package, "sheet1");
        }
    }
}
