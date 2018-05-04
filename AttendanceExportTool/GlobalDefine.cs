using System;
using System.Globalization;
using System.IO;
using System.Linq;
using AttendanceExportTool.Util;
using Newtonsoft.Json;

namespace AttendanceExportTool
{
    class ExportExcelSetInfo
    {
        public string Title;
        public string Author;
        public string Company;
        public int SignTitleFontSize;
        public int SignCellHeight;
        public string SignTitleBackgroundColor;
    }

    class MemberInfo
    {
        public string Name;
        public int Id;
        public int WorkAddress;
        public int WorkTimeType;
        public DateTime DimissionTime;
    }

    class MemberWorkTime
    {
        public int Id;
        public string Type;
        public string StartTime;
        public string EndTime;

        public bool IsHolidayTime(int day)
        {
            int index = day % Type.Length;
            char c = Type[index];
            return int.Parse(c.ToString()) == 0;
        }

        public bool IsClockInDelay(DateTime t)
        {
            DateTimeFormatInfo dtFormatInfo = new DateTimeFormatInfo();
            dtFormatInfo.ShortDatePattern = "HH:mm";
            DateTime startTime = DateTime.Parse(StartTime, dtFormatInfo);
            int daySec = startTime.Hour * 3600 + startTime.Minute * 60;
            int signSec = t.Hour * 3600 + t.Minute * 60 + t.Second;
            return daySec + GlobalDefine.Instance.Config.ClockThreshold * 60 < signSec;
        }

        public bool IsClockOffDelay(DateTime t)
        {
            DateTimeFormatInfo dtFormatInfo = new DateTimeFormatInfo();
            dtFormatInfo.ShortDatePattern = "HH:mm";
            DateTime startTime = DateTime.Parse(StartTime, dtFormatInfo);
            int daySec = startTime.Hour * 3600 + startTime.Minute * 60;
            int signSec = t.Hour * 3600 + t.Minute * 60 + t.Second;
            return daySec - GlobalDefine.Instance.Config.ClockThreshold * 60 > signSec;
        }
    }

    class WorkAddress
    {
        public int Id;
        public string[] AddressList;

        public bool IsInAddressRange(string address)
        {
            if (AddressList == null || AddressList.Length == 0)
            {
                return true;
            }

            return AddressList.FirstOrDefault(p => p != null && address.Contains(p)) != null;
        }
    }

    class ConfigData
    {
        public int CurrentMonth;
        public string ImportPath;
        public string ExportPath;
        public ExportExcelSetInfo ExportExcelSetting;
        public WorkAddress[] WorkAddress;
        public MemberInfo[] MemberInfoList;
        public MemberWorkTime[] MemberWorkTimeList;
        public string[] HolidayTimeList;
        public int ClockThreshold;

        public MemberInfo GetMemberInfo(string name)
        {
            var member = MemberInfoList.FirstOrDefault(p => p.Name == name);
            if (member == null)
            {
                LogController.Log("Can not find " + name);
            }

            return member;
        }

        public WorkAddress GetWorkAddress(int addressId)
        {
            var address = WorkAddress.FirstOrDefault(p => p.Id == addressId);
            if (address == null)
            {
                LogController.Log(addressId + " work address type error.");
            }

            return address;
        }

        public MemberWorkTime GetWorkTime(int timeId)
        {
            var time = MemberWorkTimeList.FirstOrDefault(p => p.Id == timeId);
            if (time == null)
            {
                LogController.Log(timeId + " work time type error.");
            }
            return time;
        }
    }


    class GlobalDefine : Singleton<GlobalDefine>, IInit
    {
        public static string[] WorkTypeStrings =
        {
            "上班",
            "上午未报",
            "下午未报",
            "全天未报",
            "迟到",
            "早退",
            "迟到早退",
            "休息",
            "离职"
        };

        public static string[] WorkTypeColor =
        {
            "#FFFFFF",
            "#FF99FF",
            "#FF99FF",
            "#FF99FF",
            "#99CCFF",
            "#99CCFF",
            "#99CCFF",
            "#FFF2CC",
            "#C00000"
        };

        public static string[] SignExcelTitle =
        {
            "姓名",
            "出勤天数",
            "休息天数",
            "请假天数",
            "迟到/早退次数",
            "漏报天数",
        };

        public const int MIDDAY = 12;

        public ConfigData Config;

        public InitCode Init()
        {
            try
            {
                ReadConfigData();
            }
            catch (Exception e)
            {
                LogController.Log(e.ToString());
                return InitCode.ConfigDataLoadFailed;
            }

            return InitCode.Ok;
        }

        private void ReadConfigData()
        {
            string dataPath = Environment.CurrentDirectory + "/Data.json";
            LogController.Log(dataPath + "loading...");
            using (var sr = new StreamReader(dataPath))
            {
                JsonTextReader reader = new JsonTextReader(sr);
                JsonSerializer se = new JsonSerializer();
                Config = se.Deserialize<ConfigData>(reader);
            }
            LogController.Log(dataPath + "loaded.");
        }
    }
}
