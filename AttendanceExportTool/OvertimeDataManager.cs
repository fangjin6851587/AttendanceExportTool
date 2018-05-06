using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace AttendanceExportTool
{
    class OvertimeData
    {
        public string ShopType;
        public string ShoppingName;
        public string Name;
        public DateTime OverTime;
        public float OverTimeMoney;
    }

    class MemberOvertimeData
    {
        public string ShopType;
        public string ShoppingName;
        public List<DateTime> OverTime = new List<DateTime>();
        public float OverTimeMoney;
    }

    class OvertimeDataManager : ExcelReader<OvertimeDataManager>, IInit
    {
        private readonly List<OvertimeData> mOvertimeDataList = new List<OvertimeData>();

        public List<OvertimeData> GetOvertimeDatas(string name)
        {
            return mOvertimeDataList.Where(p => p.Name == name).ToList();
        }

        public List<OvertimeData> GetOvertimeDataByShoppingType(string shoppingType)
        {
            return mOvertimeDataList.Where(p => p.ShopType == shoppingType).ToList();
        }

        public Dictionary<string, MemberOvertimeData> GetOvertimeShoppingNameListByShoppingType(string shoppingType)
        {
            Dictionary<string, MemberOvertimeData> list = new Dictionary<string, MemberOvertimeData>();

            foreach (var overtimeData in GetOvertimeDataByShoppingType(shoppingType))
            {
                if (!list.TryGetValue(overtimeData.Name, out var memberOvertime))
                {
                    memberOvertime = new MemberOvertimeData();
                    list.Add(overtimeData.Name, memberOvertime);
                }

                memberOvertime.ShopType = overtimeData.ShopType;
                memberOvertime.ShoppingName = overtimeData.ShoppingName;
                if (overtimeData.OverTime > DateTime.MinValue)
                {
                    memberOvertime.OverTime.Add(overtimeData.OverTime);
                }
                memberOvertime.OverTimeMoney += overtimeData.OverTimeMoney;
            }

            return list;
        }

        public List<string> GetShopTypeList()
        {
            List<string> typeList = new List<string>();
            foreach (var overtimeData in mOvertimeDataList)
            {
                if (!typeList.Contains(overtimeData.ShopType))
                {
                    typeList.Add(overtimeData.ShopType);
                }
            }

            return typeList;
        }




        protected override string GetPath()
        {
            return GlobalDefine.Instance.Config.OverTimePath;
        }

        protected override void Load(ExcelPackage package)
        {
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                int rowStart = 1;
                for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    if (worksheet.Cells[i, 1].Value.ToString() == "系统")
                    {
                        rowStart = i + 1;
                        break;
                    }
                }

                for (int i = rowStart; i < worksheet.Dimension.End.Row; i++)
                {
                    int col = 1;
                    OvertimeData overtimeData = new OvertimeData();
                    overtimeData.ShopType = worksheet.GetValue<string>(i, col++);
                    overtimeData.ShoppingName = worksheet.GetValue<string>(i, col++);
                    overtimeData.Name = worksheet.GetValue<string>(i, col++);
                    try
                    {
                        overtimeData.OverTime = worksheet.GetValue<DateTime>(i, col++);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        throw;
                    }

                    try
                    {
                        overtimeData.OverTimeMoney = worksheet.GetValue<float>(i, col);
                    }
                    catch (Exception)
                    {
                        overtimeData.OverTimeMoney = 0;
                    }
                    mOvertimeDataList.Add(overtimeData);
                }
            }
        }
    }
}
