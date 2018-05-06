本工具用于导出丽洁考勤数据, 仅供个人使用

注意事项:
1. 每次导出当月的表格需要修改配置数据Data.json文件以下数据

	
	/当前月份
    "CurrentMonth": 4,
    //万信达导出数据路径
    "ImportSignPath": "导入表格/2018.04万信达导出数据.xlsx",
    //人事档案
    "ImportMemberPath": "导入表格/人事档案-万信达.xlsx",
    //加班记录
    "OverTimePath": "导入表格/加班记录2018.04.xlsx",
    //工资表
    "PayPathList": 
    [
        "导入表格/大润发4月促销员工资信息表.xlsx",
        "导入表格/家乐福4月促销员工资信息表.xlsx",
        "导入表格/外区家乐福4月促销员工资信息表.xlsx"
    ],
    //业务人员导出数据表格路径
    "BusinessExportPath": "导出表格/2018年销售部门考勤明细表.xlsx",
    //导购人员导出数据表格路径
    "ShoppingGuideExportPath": "导出表格/2018导购考勤明细表.xlsx",
    //行政人员导出数据表格路径
    "AdministrativeExportPath":"导出表格/2018年行政部门考勤明细表.xlsx",
	
