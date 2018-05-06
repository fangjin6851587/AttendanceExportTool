using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AttendanceExportTool.Util;
using OfficeOpenXml;

namespace AttendanceExportTool
{
    class MemberInfo
    {
        public int Id; //员工编码
        public string Name; //员工名称
        public string Sex; //性别
        public string IdNumber; //身份证号
        public string PhoneNumber; //联系电话
        public string JoinTime; //就职日期
        public string ExitTime; //离职日期
        public string Job; //备注
        public string ShoppingName; //所属门店

        public override string ToString()
        {
            return
                $"员工编码:{Id} 员工名称:{Name} 性别:{Sex} 身份证号:{IdNumber} 联系电话:{PhoneNumber} 就职日期:{JoinTime} 离职日期:{ExitTime} 备注:{Job} 所属门店:{ShoppingName}\n";
        }
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
    }


    class MemberDataManager : ExcelReader<MemberDataManager>, IInit
    {

        private Dictionary<int, MemberInfo> mMemberInfos = new Dictionary<int, MemberInfo>();

        private Dictionary<int, MemberInfo> mShoppingGuideList = new Dictionary<int, MemberInfo>();

        public MemberInfo GetShoppingGuideMemberInfo(string name)
        {
            foreach (var member in mShoppingGuideList.Values)
            {
                if (member.Name == name)
                {
                    return member;
                }
            }

            LogController.Log(name + " can not find in shopping guide member list.");
            return null;
        }

        protected override string GetPath()
        {
            return GlobalDefine.Instance.Config.ImportMemberPath;
        }

        protected override void Load(ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
            for (int i = 2; i < worksheet.Dimension.End.Row; i++)
            {
                int col = 2;
                MemberInfo member = new MemberInfo();
                member.Id = worksheet.GetValue<int>(i, col++);
                member.Name = worksheet.GetValue<string>(i, col++);

                if (string.IsNullOrEmpty(member.Name))
                {
                    continue;
                }

                col++;
                member.Sex = worksheet.GetValue<string>(i, col++);
                member.IdNumber = worksheet.GetValue<string>(i, col++);
                col++;
                member.PhoneNumber = worksheet.GetValue<string>(i, col++);
                col += 2;
                member.JoinTime = worksheet.GetValue<string>(i, col++);
                member.ExitTime = worksheet.GetValue<string>(i, col++);
                member.Job = worksheet.GetValue<string>(i, col++);
                member.ShoppingName = worksheet.GetValue<string>(i, col);

                LogController.Log(member.ToString());

                if (!mMemberInfos.ContainsKey(member.Id))
                {
                    mMemberInfos.Add(member.Id, member);
                }

                if (member.GetWorkerType() == WorkerType.ShoppingGuide)
                {
                    if (!mShoppingGuideList.ContainsKey(member.Id))
                    {
                        mShoppingGuideList.Add(member.Id, member);
                    }
                }
            }
        }
    }
}
