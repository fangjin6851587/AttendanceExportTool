using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceExportTool
{
    enum InitCode
    {
        Ok = 0,
        ConfigDataLoadFailed,
        ExcelInitFailed,
        ExcelInitPathNoExist,
    }

    interface IInit
    {
        InitCode Init();
    }
}
