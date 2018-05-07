using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using AttendanceExportTool.Util;
using Newtonsoft.Json;

namespace AttendanceExportTool
{
    enum WorkAddressType
    {
        None,
        InCompany,
        ShopPatrol,
        InShop,
    }


    class ExportExcelSetInfo
    {
        public string Title;
        public string Author;
        public string Company;
        public int SignTitleFontSize;
        public float SignCellHeight;
        public float SignCellWidth;
        public string SignTitleBackgroundColor;
    }

    class WorkAddress
    {
        public int Id;
        public string Type;
        public string SubType;
        public string TypeValue;
        public string StartTime;
        public string EndTime;
        public string[] AddressList;

        public bool IsWeekendTime(int day)
        {
            if (string.IsNullOrEmpty(TypeValue))
            {
                return false;
            }

            DateTime dt = new DateTime(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth, day);

            int index = (int)dt.DayOfWeek - 1;
            if (index < 0)
            {
                index = Enum.GetValues(typeof(DayOfWeek)).Length - 1;
            }
            char c = TypeValue[index];
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
            DateTime endTime = DateTime.Parse(EndTime, dtFormatInfo);
            int daySec = endTime.Hour * 3600 + endTime.Minute * 60;
            int signSec = t.Hour * 3600 + t.Minute * 60 + t.Second;
            return daySec - GlobalDefine.Instance.Config.ClockThreshold * 60 > signSec;
        }

        public bool IsInAddressRange(string address)
        {
            if (AddressList == null || AddressList.Length == 0)
            {
                return true;
            }

            return AddressList.FirstOrDefault(p => p != null && address.Contains(p)) != null;
        }
    }

    class  HolidayTime
    {
        public string Title;
        public DateTime StartTime;
        public DateTime EndTime;

        public bool IsInHolidayTime(int day)
        {
            DateTime currDt = new DateTime(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth, day);

            if (StartTime <= currDt && EndTime >= currDt)
            {
                return true;
            }

            return false;
        }
    }

    class ConfigData
    {
        public int CurrentMonth;
        public string ImportSignPath;
        public string ImportMemberPath;
        public string OverTimePath;
        public string[] PayPathList;
        public string BusinessExportPath;
        public string ShoppingGuideExportPath;
        public string AdministrativeExportPath;
        public ExportExcelSetInfo ExportExcelSetting;
        public SpecialMember[] SpecailMemberList;
        public WorkAddress[] WorkAddress;
        public HolidayTime[] HolidayTimeList;
        public int ClockThreshold;

        public List<WorkAddress> GetWorkAddresses(WorkerType workerType)
        {
            return WorkAddress.Where(address => address.Type == workerType.ToString()).ToList();
        }

        public WorkAddress GetWorkAddress(WorkerType workerType, WorkAddressType workAddressType)
        {
            var address = WorkAddress.FirstOrDefault(p => p.Type == workerType.ToString() && p.SubType == workAddressType.ToString());
            if (address == null)
            {
                LogController.Log(workerType + " " + workAddressType + " work address type error.");
            }

            return address;
        }

        public SpecialMember FindSpeicalMember(int id)
        {
            return SpecailMemberList.FirstOrDefault(p => p.Id == id);
        }
    }

    enum WorkerType
    {
        None = -1,
        Business = 0,
        Administration,
        ShoppingGuide,
    }

    class SpecialMember
    {
        public int Id;
        public string Type;

        public WorkerType GetWorkerType()
        {
            return (WorkerType)Enum.Parse(typeof(WorkerType), Type);
        }
    }


    class GlobalDefine : Singleton<GlobalDefine>, IInit
    {
        public string ExportDir;

        public static readonly string[] WORK_TYPE_STRINGS = new[]
        {
            "做六休一",
            "做一休一",
            "做五休二",
        };

        public static readonly int[] WORK_TYPE_DAY = new[]
        {
            4,
            15,
            8,
        };

        public static readonly string[] SHOPPING_EXCEL_TITLES = new[]
        {
            "姓名",
            "上班类型",
            "入职时间",
            "离职时间",
            "加班",
            "请假",
            "考勤备注"
        };

        public static readonly string[][] BUSINESS_WORK_TYPE_STRINGS =
        {
            new []{ "业务", "督导", "市场部" },
            new []{ "行政主管", "行政助理", "财务", "人事" },
            new []{ "精英队", "促销流动", "导购员" },
        };

        public static readonly string[] WORK_TIME_TYPE_STRINGS =
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

        public static readonly string MONDAY_COLOR = "#FFFF00";

        public static readonly string[] WORK_TYPE_COLOR =
        {
            "#FFFFFF",
            "#FF66FF",
            "#FF66FF",
            "#FF66FF",
            "#99CCFF",
            "#99CCFF",
            "#99CCFF",
            "#FFF2CC",
            "#B1AAAA"
        };

        public static readonly string[] BUSINESS_EXCEL_TITLE =
        {
            "姓名",
            "出勤天数",
            "休息天数",
            "请假天数",
            "迟到/早退次数",
            "漏报天数",
        };

        public const int MIDDAY_HOUR = 12;

        internal ConfigData Config { get; set; }

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
