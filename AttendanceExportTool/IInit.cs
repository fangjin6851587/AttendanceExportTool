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
        AttendanceDataLoadFailed = 100,
        ConfigDataLoadFailed,
    }

    interface IInit
    {
        InitCode Init();
    }
}
