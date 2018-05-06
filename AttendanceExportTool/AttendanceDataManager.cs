using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AttendanceExportTool.Util;
using OfficeOpenXml;

namespace AttendanceExportTool
{
    class AttendanceImportData
    {
        public int Id; //人员编号
        public string Name; //人员名称
        public string Type; //考勤类型
        public DateTime SignTime; //签到时间
        public float Longitude; //经度
        public float Latitude; //纬度
        public string Address; //定位地址
        public string ShopName; //门店名称
        public string ShopId; //门店编号
        public string Job; //职务名称

        public WorkerType GetWorkerType()
        {
            for (int i = 0; i < GlobalDefine.BUSINESS_WORK_TYPE_STRINGS.Length; i++)
            {

                var type = GlobalDefine.BUSINESS_WORK_TYPE_STRINGS[i];
                if (GlobalDefine.BUSINESS_WORK_TYPE_STRINGS[i].FirstOrDefault(p => Job.Contains(p)) != null)
                {
                    return (WorkerType)i;
                }
            }

            return WorkerType.None;
        }

        public override string ToString()
        {
            return
                $"人员编号:{Id} 人员名称:{Name} 考勤类型:{Type} 签到时间:{SignTime} 经度:{Longitude} 纬度:{Latitude} 定位地址:{Address} 门店名称:{ShopName} 门店编号:{ShopId} 职务名称:{Job}\n";
        }

        public string ToBusinessCommentString()
        {
            return
                $"时间:{SignTime}\n 地址:{Address}\n";
        }

        public string ToShoppingGuideCommentString()
        {
            return
                $"时间:{SignTime}\n";
        }

        public bool IsMorningSign()
        {
            return SignTime.Hour <= GlobalDefine.MIDDAY_HOUR;
        }
    }

    class WorkTypeInfo
    {
        public WorkTimeType WorkTimeType;
        public string Name = string.Empty;
        public AttendanceImportData ClockInTime;
        public AttendanceImportData ClockOffTime;

        public string GetComment()
        {
            string comment = Name + "\n";
            if (ClockInTime != null)
            {
                comment += string.Format("上班签到: \n{0}", ClockInTime.ToBusinessCommentString());
            }
            else
            {
                comment += "上班签到: 无\n";
            }

            if (ClockOffTime != null)
            {
                comment += string.Format("下班签到: \n{0}", ClockOffTime.ToBusinessCommentString());
            }
            else
            {
                comment += "下班签到: 无\n";
            }
            return comment;
        }
    }

    enum WorkTimeType
    {
        Normal = 0, //正常上班
        WorkUnClockIn, //上班未打卡
        WorkUnClockOff, //下班未打卡
        WorkUnClockInAndOff, //未打卡
        WorkLate, //迟到
        WorkLeaveEarly, //早退
        WorkLateAndLeveaEarly, //迟到早退
        WorkRest, //休息
        WorkDimission //离职
    }

    class AttendanceDataManager : ExcelReader<AttendanceDataManager>, IInit
    {
        private readonly Dictionary<int, Dictionary<int, AttendanceImportData>> mClockInDataList = new Dictionary<int, Dictionary<int, AttendanceImportData>>();
        private readonly Dictionary<int, Dictionary<int, AttendanceImportData>> mClockOffDataList = new Dictionary<int, Dictionary<int, AttendanceImportData>>();
        private Dictionary<int, string> mBusinessMemberNameList = new Dictionary<int, string>();
        private Dictionary<int, string> mAdministrativeMemberNameList = new Dictionary<int, string>();
        private Dictionary<string, int> mMemberNameList = new Dictionary<string, int>();

        public Dictionary<int, string> BusinessMemberNameList => mBusinessMemberNameList;
        public Dictionary<string, int> MemberNameList => mMemberNameList;

        public Dictionary<int, string> AdministrativeMemberNameList => mAdministrativeMemberNameList;



        public List<int> GetUnClockTimeList(string name)
        {
            List<int> unClockTimeList = new List<int>();
            int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year , GlobalDefine.Instance.Config.CurrentMonth);

            if (mMemberNameList.ContainsKey(name))
            {
                var signData = mClockInDataList[mMemberNameList[name]];
                for (int i = 1; i <= days; i++)
                {
                    if (!signData.ContainsKey(i))
                    {
                        unClockTimeList.Add(i);
                    }
                }
            }
            return unClockTimeList;
        }

        public List<int> GetSignTimeList(string name)
        {
            List<int> signDayList = new List<int>();
            if (mMemberNameList.ContainsKey(name))
            {
                var signData = mClockInDataList[mMemberNameList[name]];
                signDayList.AddRange(signData.Keys);
            }
            return signDayList;
        }

        public List<AttendanceImportData> GetUnClockInTimeList(string name)
        {
            List<AttendanceImportData> unClockTimeList = new List<AttendanceImportData>();
            int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth);

            if (mMemberNameList.ContainsKey(name))
            {
                var signClockInData = mClockInDataList[mMemberNameList[name]];
                var signClockOffData = mClockOffDataList[mMemberNameList[name]];
                for (int i = 1; i <= days; i++)
                {
                    if (signClockInData.ContainsKey(i))
                    {
                        AttendanceImportData clockInData = signClockInData[i];
                        AttendanceImportData clockOffData = signClockOffData[i];
                        if (clockInData != null && clockOffData == clockInData && !clockOffData.IsMorningSign())
                        {
                            unClockTimeList.Add(clockInData);
                        }
                    }
                }
            }

            return unClockTimeList;
        }

        public List<AttendanceImportData> GetUnClockOffTimeList(string name)
        {
            List<AttendanceImportData> unClockTimeList = new List<AttendanceImportData>();
            int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth);

            if (mMemberNameList.ContainsKey(name))
            {
                var signClockInData = mClockInDataList[mMemberNameList[name]];
                var signClockOffData = mClockOffDataList[mMemberNameList[name]];
                for (int i = 1; i <= days; i++)
                {
                    if (signClockInData.ContainsKey(i))
                    {
                        AttendanceImportData clockInData = signClockInData[i];
                        AttendanceImportData clockOffData = signClockOffData[i];
                        if (clockInData != null && clockOffData == clockInData && clockInData.IsMorningSign())
                        {
                            unClockTimeList.Add(clockInData);
                        }
                    }
                }
            }

            return unClockTimeList;
        }

        public int GetBusinessWorkCount(int id, WorkTimeType workTimeType)
        {
            int count = 0;
            int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year , GlobalDefine.Instance.Config.CurrentMonth);
            for (int i = 0; i < days; i++)
            {
                if (GetBusinessWorkType(id, i + 1).WorkTimeType == workTimeType)
                {
                    count++;
                }
            }

            return count;
        }

        public WorkTypeInfo GetBusinessWorkType(int id, int day)
        {

            WorkTypeInfo workTypeInfo = new WorkTypeInfo
            {
                Name = mBusinessMemberNameList[id],
                ClockInTime = GetClockInData(id, day),
                ClockOffTime = GetClockOffData(id, day)
            };

            if (workTypeInfo.ClockInTime == null && workTypeInfo.ClockOffTime == null)
            {
                var workAdress =
                    GlobalDefine.Instance.Config.GetWorkAddress(WorkerType.Business, WorkAddressType.InCompany);
                workTypeInfo.WorkTimeType = workAdress.IsWeekendTime(day) ? WorkTimeType.WorkRest : WorkTimeType.WorkUnClockInAndOff;
            }
            else if (workTypeInfo.ClockInTime == workTypeInfo.ClockOffTime)
            {
                workTypeInfo.WorkTimeType = workTypeInfo.ClockInTime.IsMorningSign() ? WorkTimeType.WorkUnClockOff : WorkTimeType.WorkUnClockIn;
                if (workTypeInfo.WorkTimeType == WorkTimeType.WorkUnClockIn)
                {
                    workTypeInfo.ClockInTime = null;
                }
                else
                {
                    workTypeInfo.ClockOffTime = null;
                }
            }
            else if (workTypeInfo.ClockOffTime != workTypeInfo.ClockInTime)
            {
                WorkAddress workAddress =
                    AttendanceDataManager.Instance.GetBusinessWorkAddress(workTypeInfo.ClockInTime, workTypeInfo.ClockOffTime);
                if (workAddress.IsClockInDelay(workTypeInfo.ClockInTime.SignTime))
                {
                    workTypeInfo.WorkTimeType = WorkTimeType.WorkLate;
                }

                if (workAddress.IsClockOffDelay(workTypeInfo.ClockOffTime.SignTime))
                {
                    workTypeInfo.WorkTimeType = workTypeInfo.WorkTimeType == WorkTimeType.WorkLate ? WorkTimeType.WorkLateAndLeveaEarly : WorkTimeType.WorkLeaveEarly;
                }
            }
            return workTypeInfo;
        }

        public WorkTypeInfo GetAdministrationWorkType(int id, int day)
        {
            WorkTypeInfo workTypeInfo = new WorkTypeInfo
            {
                Name = mAdministrativeMemberNameList[id],
                ClockInTime = GetClockInData(id, day),
                ClockOffTime = GetClockOffData(id, day)
            };

            var workAddress =
                GlobalDefine.Instance.Config.GetWorkAddress(WorkerType.Administration, WorkAddressType.None);

            if (workTypeInfo.ClockInTime != null && !workAddress.IsInAddressRange(workTypeInfo.ClockInTime.Address))
            {
                workTypeInfo.ClockInTime = null;
            }

            if (workTypeInfo.ClockOffTime != null && !workAddress.IsInAddressRange(workTypeInfo.ClockOffTime.Address))
            {
                workTypeInfo.ClockOffTime = null;
            }

            if (workTypeInfo.ClockInTime == null && workTypeInfo.ClockOffTime == null)
            {
                workTypeInfo.WorkTimeType = workAddress.IsWeekendTime(day) ? WorkTimeType.WorkRest : WorkTimeType.WorkUnClockInAndOff;
            }
            else if (workTypeInfo.ClockInTime == workTypeInfo.ClockOffTime)
            {
                workTypeInfo.WorkTimeType = workTypeInfo.ClockInTime.IsMorningSign() ? WorkTimeType.WorkUnClockOff : WorkTimeType.WorkUnClockIn;
                if (workTypeInfo.WorkTimeType == WorkTimeType.WorkUnClockIn)
                {
                    workTypeInfo.ClockInTime = null;
                }
                else
                {
                    workTypeInfo.ClockOffTime = null;
                }
            }
            else if (workTypeInfo.ClockInTime == null)
            {
                workTypeInfo.WorkTimeType = WorkTimeType.WorkUnClockIn;
            }
            else if (workTypeInfo.ClockOffTime == null)
            {
                workTypeInfo.WorkTimeType = WorkTimeType.WorkUnClockOff;
            }
            else if (workTypeInfo.ClockOffTime != workTypeInfo.ClockInTime)
            {
                if (workAddress.IsClockInDelay(workTypeInfo.ClockInTime.SignTime))
                {
                    workTypeInfo.WorkTimeType = WorkTimeType.WorkLate;
                }
                if (workAddress.IsClockOffDelay(workTypeInfo.ClockOffTime.SignTime))
                {
                    workTypeInfo.WorkTimeType = workTypeInfo.WorkTimeType == WorkTimeType.WorkLate ? WorkTimeType.WorkLateAndLeveaEarly : WorkTimeType.WorkLeaveEarly;
                }
            }
            return workTypeInfo;
        }

        public WorkAddress GetBusinessWorkAddress(AttendanceImportData clockInTime,
            AttendanceImportData clockOffTime)
        {
            var workAdress =
                GlobalDefine.Instance.Config.GetWorkAddress(WorkerType.Business, WorkAddressType.InCompany);
            if (workAdress.IsInAddressRange(clockInTime.Address) && workAdress.IsInAddressRange(clockOffTime.Address))
            {
                return workAdress;
            }

            if (clockInTime.Address.Substring(0, 9) == clockOffTime.Address.Substring(0, 9))
            {
                workAdress =
                    GlobalDefine.Instance.Config.GetWorkAddress(WorkerType.Business, WorkAddressType.InShop);
            }
            else
            {
                workAdress =
                    GlobalDefine.Instance.Config.GetWorkAddress(WorkerType.Business, WorkAddressType.ShopPatrol);
            }

            return workAdress;
        }

        private AttendanceImportData GetClockInData(int id, int day)
        {
            AttendanceImportData data = null;
            if (mClockInDataList.TryGetValue(id, out var dic))
            {
                dic.TryGetValue(day, out data);
            }
            return data;
        }

        private AttendanceImportData GetClockOffData(int id, int day)
        {
            AttendanceImportData data = null;
            if (mClockOffDataList.TryGetValue(id, out var dic))
            {
                dic.TryGetValue(day, out data);
            }
            return data;
        }


        protected override string GetPath()
        {
            return GlobalDefine.Instance.Config.ImportSignPath;
        }

        protected override void Load(ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
            for (int i = 2; i < worksheet.Dimension.End.Row; i++)
            {
                AttendanceImportData rowData = new AttendanceImportData();
                int col = 1;
                rowData.Type = worksheet.GetValue<string>(i, col++);
                rowData.SignTime = worksheet.GetValue<DateTime>(i, col++);
                rowData.Longitude = worksheet.GetValue<float>(i, col++);
                rowData.Latitude = worksheet.GetValue<float>(i, col++);
                rowData.Address = worksheet.GetValue<string>(i, col++);
                rowData.Id = worksheet.GetValue<int>(i, col++);
                rowData.Name = worksheet.GetValue<string>(i, col++);
                rowData.ShopName = worksheet.GetValue<string>(i, col++);
                rowData.ShopId = worksheet.GetValue<string>(i, col++);
                rowData.Job = worksheet.GetValue<string>(i, col);

                if (string.IsNullOrEmpty(rowData.Job))
                {
                    continue;
                }

                if (rowData.SignTime.Month != GlobalDefine.Instance.Config.CurrentMonth)
                {
                    continue;
                }

                var specialMember = GlobalDefine.Instance.Config.FindSpeicalMember(rowData.Id);
                if (!mBusinessMemberNameList.ContainsKey(rowData.Id) && (rowData.GetWorkerType() == WorkerType.Business || specialMember != null && specialMember.GetWorkerType() == WorkerType.Business))
                {
                    mBusinessMemberNameList.Add(rowData.Id, rowData.Name);
                }

                if (!mAdministrativeMemberNameList.ContainsKey(rowData.Id) && (rowData.GetWorkerType() == WorkerType.Administration || specialMember != null && specialMember.GetWorkerType() == WorkerType.Administration))
                {
                    mAdministrativeMemberNameList.Add(rowData.Id, rowData.Name);
                }

                if (!mMemberNameList.ContainsKey(rowData.Name))
                {
                    mMemberNameList.Add(rowData.Name, rowData.Id);
                }

                LogController.Log(rowData.ToString());

                if (mClockInDataList.TryGetValue(rowData.Id, out var lastData))
                {
                    if (lastData.TryGetValue(rowData.SignTime.Day, out var lastAttendanceImportData))
                    {
                        if (lastAttendanceImportData.SignTime >= rowData.SignTime)
                        {
                            lastData[rowData.SignTime.Day] = rowData;
                        }
                    }
                    else
                    {
                        lastData.Add(rowData.SignTime.Day, rowData);
                    }
                }
                else
                {
                    var dic = new Dictionary<int, AttendanceImportData>();
                    dic.Add(rowData.SignTime.Day, rowData);
                    mClockInDataList.Add(rowData.Id, dic);
                }

                if (mClockOffDataList.TryGetValue(rowData.Id, out var lastClockOffData))
                {
                    if (lastClockOffData.TryGetValue(rowData.SignTime.Day, out var lastAttendanceImportData))
                    {
                        if (lastAttendanceImportData.SignTime <= rowData.SignTime)
                        {
                            lastClockOffData[rowData.SignTime.Day] = rowData;
                        }
                    }
                    else
                    {
                        lastClockOffData.Add(rowData.SignTime.Day, rowData);
                    }
                }
                else
                {
                    var dic = new Dictionary<int, AttendanceImportData>();
                    dic.Add(rowData.SignTime.Day, rowData);
                    mClockOffDataList.Add(rowData.Id, dic);
                }
            }

            mBusinessMemberNameList = mBusinessMemberNameList.OrderBy(p => p.Key).ToDictionary(p => p.Key, o => o.Value);
        }
    }
}
