using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AttendanceExportTool;
using OfficeOpenXml;

namespace AttendanceExportToolWin
{
    class PayForgetMemberExporter : ExcelWriter
    {
        protected override string GetPath()
        {
            return GlobalDefine.Instance.ExportDir + "/促销员工资信息补漏表.xlsx";
        }

        protected override void Export(ExcelPackage package)
        {
            var excelSetting = GlobalDefine.Instance.Config.ExportExcelSetting;
            var sheet = package.Workbook.Worksheets.Add("sheet1");

            sheet.Cells[1, 1].Value = "员工编号";
            sheet.Cells[1, 2].Value = "姓名";
            sheet.Cells[1, 3].Value = "门店";


            int row = 2;
            foreach (KeyValuePair<int, AttendanceImportData> keyValuePair in AttendanceDataManager.Instance.ShoppingGuideMemberNameList)
            {
                int col = 1;
                var payMember = MemberPayDataManager.Instance.GetPayData(keyValuePair.Value.Name);
                if (payMember == null)
                {
                    sheet.Cells[row, col].Value = keyValuePair.Key;
                    sheet.Cells[row, col].AutoFitColumns(10);
                    col++;
                    sheet.Cells[row, col].Value = keyValuePair.Value.Name;
                    sheet.Cells[row, col].AutoFitColumns(10);
                    col++;
                    sheet.Cells[row, col].Value = keyValuePair.Value.ShopName;
                    sheet.Cells[row, col].AutoFitColumns(10);

                    row++;
                }
            }
        }
    }
}
