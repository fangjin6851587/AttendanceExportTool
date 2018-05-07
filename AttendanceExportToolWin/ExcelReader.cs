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
    abstract class ExcelReader<T>where T : IInit, new()
    {
        protected static T _instance;
        public static T Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new T();
                return _instance;
            }
        }

        public virtual InitCode Init()
        {
            try
            {
                string path = GetPath();
                LogController.Log(path + "importing...");
                FileInfo importFileInfo = new FileInfo(path);
                if (!importFileInfo.Exists)
                {
                    LogController.Log(path + "no exists.");
                    return InitCode.ExcelInitPathNoExist;
                }
                using (var package = new ExcelPackage(importFileInfo))
                {
                    Load(package);
                }
                LogController.Log(path + "import finished.");
            }
            catch (Exception e)
            {
                LogController.Log(e.ToString());
                return InitCode.ExcelInitFailed;
            }

            return InitCode.Ok;
        }

        protected abstract string GetPath();

        protected abstract void Load(ExcelPackage package);
    }
}
