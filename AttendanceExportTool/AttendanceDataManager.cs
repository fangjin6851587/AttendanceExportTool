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

    partial class AttendanceImportData
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
                $"人员编号:{Id} 人员名称:{Name} 考勤类型:{Type} 签到时间:{SignTime} 经度:{Longitude} 纬度:{Latitude} 定位地址:{Address} 门店名称:{ShopName} 门店编号:{ShopId} 职务名称:{Job}";
        }
    }

    enum WorkType
    {
        Normal, //正常上班
        WorkUnClockIn, //上班未打卡
        WorkUnClockOff, //下班未打卡
        WorkLate, //上班迟到
        WorkLeaveEarly, //早退
        WorkRest //休息
    }

    partial class AttendanceDataManager : Singleton<AttendanceDataManager>, IInit
    {
        private Dictionary<int, Dictionary<int, AttendanceImportData>> mClockInDataList = new Dictionary<int, Dictionary<int, AttendanceImportData>>();
        private Dictionary<int, Dictionary<int, AttendanceImportData>> mClockOffDataList = new Dictionary<int, Dictionary<int, AttendanceImportData>>();


        public WorkType GetWorkType(string id, int day)
        {
            WorkType workType = WorkType.Normal;
            return workType;
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

                    if (rowData.SignTime.Hour >= 12)
                    {
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
                    }
                    else
                    {
                        if (mClockOffDataList.TryGetValue(rowData.Id, out var lastData))
                        {
                            if (lastData.TryGetValue(rowData.SignTime.Day, out var lastAttendanceImportData))
                            {
                                if (lastAttendanceImportData.SignTime <= rowData.SignTime)
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
                            mClockOffDataList.Add(rowData.Id, dic);
                        }
                    }
                }
            }
            LogController.Log(GlobalDefine.Instance.Config.ImportPath + "import finished.");
        }
    }
}
