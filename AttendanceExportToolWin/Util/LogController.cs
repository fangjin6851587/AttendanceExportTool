using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceExportTool.Util
{
    class LogController
    {
        private static bool sDebug;

        public static void SetDebug(bool debug)
        {
            sDebug = debug;
        }

        public static void Log(string str)
        {
            if (sDebug)
            {
                Console.WriteLine(str);
            }
        }
    }
}
