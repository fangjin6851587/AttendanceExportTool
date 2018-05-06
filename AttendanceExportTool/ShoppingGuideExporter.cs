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

                var shoppingMemberList = OvertimeDataManager.Instance.GetOvertimeDataByShoppingType(sheetName);
                shoppingMemberList.Sort((data1, data2) => String.Compare(data1.ShoppingName, data2.ShoppingName, StringComparison.Ordinal));
                for (int i = 0; i < shoppingMemberList.Count; i++)
                {
                    int col = 1;
                    var cell = sheet.Cells[i + 2, col++];
                    cell.Value = shoppingMemberList[i].ShoppingName;
                    cell = sheet.Cells[i + 2, col++];
                    cell.Value = shoppingMemberList[i].Name;
                }


                for (int i = 1; i <= sheet.Dimension.End.Column; i++)
                {
                    var column = sheet.GetColumn(i);
                    if (column == null)
                    {
                        continue;
                    }
                    if (i == sheet.Dimension.End.Column)
                    {
                        column.Width = 65;
                    }
                    else
                    {
                        column.AutoFit(20);
                    }
                }


                string startColumn = Util.Util.ToNumberSystem26(1);
                string endColumn = Util.Util.ToNumberSystem26(sheet.Dimension.End.Column -1);
                string columnRange = $"{startColumn}:{endColumn}";

                sheet.Cells[columnRange].Style.Font.Size = excelSetting.SignTitleFontSize;
                sheet.Cells[columnRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[columnRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[columnRange].Style.WrapText = true;

                startColumn = Util.Util.ToNumberSystem26(sheet.Dimension.End.Column);
                endColumn = Util.Util.ToNumberSystem26(sheet.Dimension.End.Column);
                columnRange = $"{startColumn}:{endColumn}";
                sheet.Cells[columnRange].Style.Font.Size = excelSetting.SignTitleFontSize;
                sheet.Cells[columnRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells[columnRange].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                sheet.Cells[columnRange].Style.WrapText = true;

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
            return GlobalDefine.Instance.Config.ShoppingGuideExportPath;
        }

        protected override void Export(ExcelPackage package)
        {
            foreach (var shoppingType in OvertimeDataManager.Instance.GetShopTypeList())
            {
                new ShoppingSheet().Create(package, shoppingType);
            }
        }
    }
}
