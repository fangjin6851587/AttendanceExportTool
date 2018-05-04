using System;
using System.Collections.Generic;
using System.IO;
using AttendanceExportTool.Util;
using OfficeOpenXml;
using OfficeOpenXml.Style;

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

        public override string ToString()
        {
            return
                $"人员编号:{Id} 人员名称:{Name} 考勤类型:{Type} 签到时间:{SignTime} 经度:{Longitude} 纬度:{Latitude} 定位地址:{Address} 门店名称:{ShopName} 门店编号:{ShopId} 职务名称:{Job}\n";
        }

        public string ToCommentString()
        {
            return
                $"人员编号:{Id} 人员名称:{Name} 职务名称:{Job} 签到时间:{SignTime} 定位地址:{Address} 门店名称:{ShopName}\n";
        }
    }

    class WorkTypeInfo
    {
        public WorkType WorkType;
        public AttendanceImportData ColorIn;
        public AttendanceImportData ColorOut;

        public string GetComment()
        {
            string comment = string.Empty;
            if (ColorIn != null)
            {
                comment += string.Format("上班签到: {0}", ColorIn.ToCommentString());
            }
            else
            {
                comment += "上班签到: 无\n";
            }

            if (ColorOut != null)
            {
                comment += string.Format("下班签到: {0}", ColorOut.ToCommentString());
            }
            else
            {
                comment += "下班签到: 无\n";
            }
            return comment;
        }
    }

    enum WorkType
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

    class AttendanceDataManager : Singleton<AttendanceDataManager>, IInit
    {
        private readonly Dictionary<string, Dictionary<int, AttendanceImportData>> mClockInDataList = new Dictionary<string, Dictionary<int, AttendanceImportData>>();
        private readonly Dictionary<string, Dictionary<int, AttendanceImportData>> mClockOffDataList = new Dictionary<string, Dictionary<int, AttendanceImportData>>();

        public void Export()
        {
            LogController.Log(GlobalDefine.Instance.Config.ExportPath + " exporting...");
            if (File.Exists(GlobalDefine.Instance.Config.ExportPath))
            {
                File.Delete(GlobalDefine.Instance.Config.ExportPath);
            }

            FileInfo newFile = new FileInfo(GlobalDefine.Instance.Config.ExportPath);
            using (ExcelPackage package = new ExcelPackage())
            {
                new SignSheet().Create(package, "2");

                var excelSetting = GlobalDefine.Instance.Config.ExportExcelSetting;
                package.Workbook.Properties.Title = excelSetting.Title;
                package.Workbook.Properties.Author = excelSetting.Author;
                package.Workbook.Properties.Company = excelSetting.Company;
                package.SaveAs(newFile);
            }
            LogController.Log(GlobalDefine.Instance.Config.ExportPath + " export finished.");
        }

        public int GetWorkCount(string name, WorkType workType)
        {
            int count = 0;
            int days = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year , GlobalDefine.Instance.Config.CurrentMonth);
            for (int i = 0; i < days; i++)
            {
                if (GetWorkType(name, i + 1).WorkType == workType)
                {
                    count++;
                }
            }

            return count;
        }

        public WorkTypeInfo GetWorkType(string name, int day)
        {
            WorkTypeInfo workTypeInfo = new WorkTypeInfo();

            MemberInfo member = GlobalDefine.Instance.Config.GetMemberInfo(name);

            DateTime dt = new DateTime(DateTime.Now.Year, GlobalDefine.Instance.Config.CurrentMonth, day);

            if (member == null || member.DimissionTime <= dt)
            {
                workTypeInfo.WorkType = WorkType.WorkDimission;
                return workTypeInfo;
            }

            MemberWorkTime workTime = GlobalDefine.Instance.Config.GetWorkTime(member.WorkTimeType);
            if (workTime == null)
            {
                workTypeInfo.WorkType = WorkType.WorkRest;
                return workTypeInfo;
            }

            AttendanceImportData clockInData = GetClockInData(name, day);
            if (clockInData != null && !IsInSignAddressRange(name, clockInData.Address))
            {
                clockInData = null;
            }

            AttendanceImportData clockOffData = GetClockOffData(name, day);
            if (clockOffData != null && !IsInSignAddressRange(name, clockOffData.Address))
            {
                clockOffData = null;
            }


            workTypeInfo.ColorIn = clockInData;
            workTypeInfo.ColorOut = clockOffData;

            if (clockInData != null && clockOffData != null && clockInData != clockOffData)
            {
                if (workTime.IsClockInDelay(clockInData.SignTime) && workTime.IsClockOffDelay(clockOffData.SignTime))
                {
                    workTypeInfo.WorkType = WorkType.WorkLateAndLeveaEarly;
                    return workTypeInfo;
                }

                if (workTime.IsClockInDelay(clockInData.SignTime))
                {
                    workTypeInfo.WorkType = WorkType.WorkLate;
                    return workTypeInfo;
                }

                if (workTime.IsClockOffDelay(clockOffData.SignTime))
                {
                    workTypeInfo.WorkType = WorkType.WorkLeaveEarly;
                    return workTypeInfo;
                }

                workTypeInfo.WorkType = WorkType.Normal;
                return workTypeInfo;
            }

            if(clockInData != null && clockOffData != null && clockInData == clockOffData)
            {
                if (clockInData.SignTime.Hour < GlobalDefine.MIDDAY)
                {
                    workTypeInfo.WorkType = WorkType.WorkUnClockOff;
                    return workTypeInfo;
                }

                workTypeInfo.WorkType = WorkType.WorkUnClockIn;
                return workTypeInfo;
            }

            if (IsHolidayTime(name, day))
            {
                workTypeInfo.WorkType = WorkType.WorkRest;
                return workTypeInfo;
            }

            workTypeInfo.WorkType = WorkType.WorkUnClockInAndOff;
            return workTypeInfo;
        }

        private bool IsHolidayTime(string name, int day)
        {
            MemberInfo member = GlobalDefine.Instance.Config.GetMemberInfo(name);
            if (member == null)
            {
                return true;
            }
            MemberWorkTime workTime = GlobalDefine.Instance.Config.GetWorkTime(member.WorkTimeType);
            if (workTime == null)
            {
                return true;
            }

            return workTime.IsHolidayTime(day);
        }

        private bool IsInSignAddressRange(string name, string address)
        {
            MemberInfo member = GlobalDefine.Instance.Config.GetMemberInfo(name);
            if (member == null)
            {
                return true;
            }

            WorkAddress workAddress = GlobalDefine.Instance.Config.GetWorkAddress(member.WorkAddress);
            if (workAddress == null)
            {
                return true;
            }

            return workAddress.IsInAddressRange(address);
        }

        private AttendanceImportData GetClockInData(string name, int day)
        {
            AttendanceImportData data = null;
            if (mClockInDataList.TryGetValue(name, out var dic))
            {
                dic.TryGetValue(day, out data);
            }
            return data;
        }

        private AttendanceImportData GetClockOffData(string name, int day)
        {
            AttendanceImportData data = null;
            if (mClockOffDataList.TryGetValue(name, out var dic))
            {
                dic.TryGetValue(day, out data);
            }
            return data;
        }


        public InitCode Init()
        {
            try
            {
                Load();
            }
            catch (Exception e)
            {
                LogController.Log(e.ToString());
                return InitCode.AttendanceDataLoadFailed;
            }

            return InitCode.Ok;
        }

        private void Load()
        {
            FileInfo importFileInfo = new FileInfo(GlobalDefine.Instance.Config.ImportPath);
            if (!importFileInfo.Exists)
            {
                LogController.Log(GlobalDefine.Instance.Config.ImportPath + "no exists.");
                return;
            }

            LogController.Log(GlobalDefine.Instance.Config.ImportPath + "importing...");
            using (var package = new ExcelPackage(importFileInfo))
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

                    if (rowData.SignTime.Month != GlobalDefine.Instance.Config.CurrentMonth)
                    {
                        continue;
                    }

                    LogController.Log(rowData.ToString());

                    if (mClockInDataList.TryGetValue(rowData.Name, out var lastData))
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
                        mClockInDataList.Add(rowData.Name, dic);
                    }

                    if (mClockOffDataList.TryGetValue(rowData.Name, out var lastClockOffData))
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
                        mClockOffDataList.Add(rowData.Name, dic);
                    }
                }
            }
            LogController.Log(GlobalDefine.Instance.Config.ImportPath + "import finished.");
        }
    }
}
