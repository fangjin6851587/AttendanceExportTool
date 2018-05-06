using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AttendanceExportTool.Util;

namespace AttendanceExportTool
{
    class Program
    {
        static void Main(string[] args)
        {
            LogController.SetDebug(true);

            List<IInit> initList = new List<IInit>
            {
                GlobalDefine.Instance, 
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
                    LogController.Log("Init failed code = " + code);
                    return;
                }
            }

            List<ExcelWriter> excelWriters = new List<ExcelWriter>
            {
                new BusinessExcelExporter(),
                new ShoppingGuideExporter(),
                new AdministrativeExport(),
            };
            foreach (ExcelWriter excelWriter in excelWriters)
            {
                excelWriter.Save();
            }

            LogController.Log("===============================Init success=================================");
        }
    }
}
