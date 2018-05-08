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
    class MemberPayData
    {
        public string Name;
        public string WorkType;
    }

    class MemberPayDataManager : ExcelReader<MemberPayDataManager>, IInit
    {
        private List<MemberPayData> mMemberPayDataList = new List<MemberPayData>();

        public List<MemberPayData> GetMemberPayDataList()
        {
            return mMemberPayDataList;
        }

        public MemberPayData GetPayData(string name)
        {
            return mMemberPayDataList.Find(p => p.Name == name);
        }

        protected override string GetPath()
        {
            return String.Join("|", GlobalDefine.Instance.Config.PayPathList);
        }

        public override InitCode Init()
        {
            try
            {
                foreach (var path in GetPath().Split('|'))
                {
                    LogController.Log(path + "importing...");
                    FileInfo importFileInfo = new FileInfo(path);
                    if (!importFileInfo.Exists)
                    {
                        LogController.Log(path + "no exists.");
                        return InitCode.ExcelInitPathNoExist;
                    }
                    using (var package = new ExcelPackage(importFileInfo))
                    {
                        Load(package);
                    }
                    LogController.Log(path + "import finished.");
                }
            }
            catch (Exception e)
            {
                LogController.Log(e.ToString());
                return InitCode.ExcelInitFailed;
            }

            return InitCode.Ok;
        }


        protected override void Load(ExcelPackage package)
        {

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                int rowStart = 0;
                int nameColumn = 0;
                int workTypeColumn = 0;
                for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                    {
                        var value = worksheet.Cells[i, j].Value;
                        if (value == null)
                        {
                             continue;
                        }

                        if (worksheet.Cells[i, j].Value.ToString() == "上班类型")
                        {
                            rowStart = i + 1;
                            workTypeColumn = j;
                        }
                        else if(worksheet.Cells[i, j].Value.ToString() == "姓名")
                        {
                            nameColumn = j;
                        }
                    }

                    if (rowStart > 0)
                    {
                        break;
                    }
                }

                if (rowStart > 0)
                {
                    for (int i = rowStart; i <= worksheet.Dimension.End.Row; i++)
                    {
                        var cell = worksheet.Cells[i, nameColumn];

                        if (cell == null)
                        {
                            continue;
                        }

                        MemberPayData memberPay = new MemberPayData();
                        memberPay.Name = cell.GetValue<string>();

                        if (memberPay.Name == "姓名")
                        {
                            continue;
                        }

                        cell = worksheet.Cells[i, workTypeColumn];
                        memberPay.WorkType = cell.GetValue<string>();
                        mMemberPayDataList.Add(memberPay);
                    }
                }
            }
        }
    }
}
