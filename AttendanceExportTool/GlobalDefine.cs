using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AttendanceExportTool.Util;
using Newtonsoft.Json;

namespace AttendanceExportTool
{
    partial class MemberInfo
    {
        public string Name;
        public int Id;
        public int WorkAddress;
        public int WorkTimeType;
    }

    partial class MemberWorkTime
    {
        public int Id;
        public string Type;
        public string StartTime;
        public string EndTime;
    }

    class WorkAddress
    {
        public int Id;
        public string[] AddressList;
    }

    partial class ConfigData
    {
        public int CurrentMonth;
        public string ImportPath;
        public string ExportPath;
        public WorkAddress[] AddressList;
        public MemberInfo[] MemberInfoList;
        public MemberWorkTime[] MemberWorkTimeList;
        public string[] HolidayTimeList;
        public int ClockThreshold;

        public MemberInfo GetMemberInfo(int id)
        {
            return MemberInfoList.FirstOrDefault(p => p.Id == id);
        }

        public WorkAddress GetWorkAddress(int addressId)
        {
            return AddressList.FirstOrDefault(p => p.Id == addressId);
        }

        public MemberWorkTime GetWorkTime(int timeId)
        {
            return MemberWorkTimeList.FirstOrDefault(p => p.Id == timeId);
        }
    }


    partial class GlobalDefine : Singleton<GlobalDefine>, IInit
    {
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
