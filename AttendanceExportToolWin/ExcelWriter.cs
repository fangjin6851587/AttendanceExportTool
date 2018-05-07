using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AttendanceExportTool.Util;
using OfficeOpenXml;

namespace AttendanceExportTool
{
    abstract class ExcelWriter
    {
        protected abstract string GetPath();

        public void Save()
        {
            string path = GetPath();
            LogController.Log(path + " exporting...");
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            FileInfo newFile = new FileInfo(path);
            using (ExcelPackage package = new ExcelPackage())
            {
                Export(package);
                package.SaveAs(newFile);
            }
            LogController.Log(path + " export finished.");
        }

        protected abstract void Export(ExcelPackage package);
    }
}
