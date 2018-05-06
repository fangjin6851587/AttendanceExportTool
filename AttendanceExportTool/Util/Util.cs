using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceExportTool.Util
{
    class Util
    {
        public static string ToNumberSystem26(int n)
        {
            string s = string.Empty;
            while (n > 0)
            {
                int m = n % 26;
                if (m == 0) m = 26;
                s = (char)(m + 64) + s;
                n = (n - m) / 26;
            }
            return s;
        }

        public static string GetCellAddress(int row, int column)
        {
            return string.Format("{0}{1}", ToNumberSystem26(column), row);
        }
    }
}
