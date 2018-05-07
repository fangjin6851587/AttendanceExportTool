using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace AttendanceExportTool
{
    class BusinessExcelExporter : ExcelWriter
    {
        protected override string GetPath()
        {
            return GlobalDefine.Instance.ExportDir + "/" + GlobalDefine.Instance.Config.BusinessExportPath;
        }

        protected override void Export(ExcelPackage package)
        {
            new BusinessSheet().Create(package, "sheet1");
        }
    }
}
