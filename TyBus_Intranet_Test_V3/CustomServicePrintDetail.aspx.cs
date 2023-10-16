using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;

namespace TyBus_Intranet_Test_V3
{
    public partial class CustomServicePrintDetail : System.Web.UI.Page
    {
        PublicFunction PF = new PublicFunction(); //加入公用程式碼參考
        private string vLoginID = "";
        private string vLoginName = "";
        private string vLoginDepNo = "";
        private string vLoginDepName = "";
        private string vLoginTitle = "";
        private string vLoginTitleName = "";
        private string vLoginEmpType = "";
        private string vComputerName = ""; //2021.09.27 新增

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                vLoginID = (Session["LoginID"] != null) ? Session["LoginID"].ToString().Trim() : "";
                vLoginName = (Session["LoginName"] != null) ? Session["LoginName"].ToString().Trim() : "";
                vLoginDepNo = (Session["LoginDepNo"] != null) ? Session["LoginDepNo"].ToString().Trim() : "";
                vLoginDepName = (Session["LoginDepName"] != null) ? Session["LoginDepName"].ToString().Trim() : "";
                vLoginTitle = (Session["LoginTitle"] != null) ? Session["LoginTitle"].ToString().Trim() : "";
                vLoginTitleName = (Session["LoginTitleName"] != null) ? Session["LoginTitleName"].ToString().Trim() : "";
                vLoginEmpType = (Session["LoginEmpType"] != null) ? Session["LoginEmpType"].ToString().Trim() : "";
                //vComputerName = Environment.MachineName; //2021.09.27 新增取得電腦名稱
                vComputerName = Page.Request.UserHostName; //2021.10.08 改變取得電腦名稱的方法

                if (vLoginID != "")
                {
                    if (!IsPostBack)
                    {
                        PrintDataBinding();
                    }
                }
                else
                {
                    Response.Redirect("~/default.aspx");
                }
            }
            else
            {
                Response.Redirect("~/default.aspx");
            }
        }

        private void PrintDataBinding()
        {
            if (vLoginID != "")
            {
                ReportDataSource rdsPrint = new ReportDataSource();

                string vSelectStr = "";
                string vPrintMode = (Session["ServicePrintMode"].ToString().Trim() != "") ? Session["ServicePrintMode"].ToString().Trim() : "2";
                switch (vPrintMode)
                {
                    case "1":
                        vSelectStr = "SELECT ServiceNo, BuildDate, BuildTime, BuildMan, " + Environment.NewLine +
                                     "       (SELECT NAME FROM Employee WHERE(EMPNO = a.BuildMan)) AS BuildManName, " + Environment.NewLine +
                                     "       ServiceType, " + Environment.NewLine +
                                     "       (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) ServiceType_C, " + Environment.NewLine +
                                     "       ServiceTypeB, " + Environment.NewLine +
                                     "       (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) ServiceTypeB_C, " + Environment.NewLine +
                                     "       ServiceTypeC, " + Environment.NewLine +
                                     "       (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) ServiceTypeC_C, " + Environment.NewLine +
                                     "       LinesNo, Car_ID, Driver, DriverName, BoardTime, BoardStation, GetoffTime, GetoffStation, ServiceNote, IsNoContect, " + Environment.NewLine +
                                     "       CivicName, CivicTelNo, CivicTelExtNo, CivicCellPhone, CivicAddress, CivicEMail, AthorityDepNo, AthorityDepNote, " + Environment.NewLine +
                                     "       (SELECT NAME FROM Department WHERE(DEPNO = a.AthorityDepNo)) AS AthorityDepName, " + Environment.NewLine +
                                     "       LinesNo2, Car_ID2, Driver2, DriverName2, AthorityDepNo2, AthorityDepNote2, " + Environment.NewLine +
                                     "       (SELECT NAME FROM Department WHERE(DEPNO = a.AthorityDepNo2)) AS AthorityDepName2, " + Environment.NewLine +
                                     "       IsReplied, Remark, IsPending, AssignDate, AssignMan," + Environment.NewLine +
                                     "       (SELECT NAME FROM Employee AS Employee_2 WHERE(EMPNO = a.AssignMan)) AS AssignManName, " + Environment.NewLine +
                                     "       IsClosed, CloseDate, CloseMan," + Environment.NewLine +
                                     "       (SELECT NAME FROM Employee AS Employee_1 WHERE(EMPNO = a.CloseMan)) AS CloseManName, " + Environment.NewLine +
                                     "       CaseSource, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '客服專線記錄表  fmCustomService CaseSource') AND (CLASSNO = a.CaseSource)) AS CaseSource_C, " + Environment.NewLine +
                                     "       ServiceDate, IsTrue " + Environment.NewLine +
                                     "  FROM CustomService a " + Environment.NewLine +
                                     " WHERE ServiceNo in (" + Session["CustomServicePrint"].ToString().Trim() + ")";
                        break;
                    case "2":
                        vSelectStr = "SELECT ServiceNo, BuildDate, BuildTime, BuildMan, " + Environment.NewLine +
                                     "       (SELECT NAME FROM Employee WHERE(EMPNO = a.BuildMan)) AS BuildManName, " + Environment.NewLine +
                                     "       ServiceType, " + Environment.NewLine +
                                     "       (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) ServiceType_C, " + Environment.NewLine +
                                     "       ServiceTypeB, " + Environment.NewLine +
                                     "       (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) ServiceTypeB_C, " + Environment.NewLine +
                                     "       ServiceTypeC, " + Environment.NewLine +
                                     "       (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) ServiceTypeC_C, " + Environment.NewLine +
                                     "       LinesNo, Car_ID, Driver, DriverName, BoardTime, BoardStation, GetoffTime, GetoffStation, ServiceNote, IsNoContect, " + Environment.NewLine +
                                     "       CivicName, CivicTelNo, CivicTelExtNo, CivicCellPhone, CivicAddress, CivicEMail, AthorityDepNo, AthorityDepNote, " + Environment.NewLine +
                                     "       (SELECT NAME FROM Department WHERE(DEPNO = a.AthorityDepNo)) AS AthorityDepName, " + Environment.NewLine +
                                     "       LinesNo2, Car_ID2, Driver2, DriverName2, AthorityDepNo2, AthorityDepNote2, " + Environment.NewLine +
                                     "       (SELECT NAME FROM Department WHERE(DEPNO = a.AthorityDepNo2)) AS AthorityDepName2, " + Environment.NewLine +
                                     "       IsReplied, Remark, IsPending, AssignDate, AssignMan," + Environment.NewLine +
                                     "       (SELECT NAME FROM Employee AS Employee_2 WHERE(EMPNO = a.AssignMan)) AS AssignManName, " + Environment.NewLine +
                                     "       IsClosed, CloseDate, CloseMan," + Environment.NewLine +
                                     "       (SELECT NAME FROM Employee AS Employee_1 WHERE(EMPNO = a.CloseMan)) AS CloseManName, " + Environment.NewLine +
                                     "       CaseSource, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '客服專線記錄表  fmCustomService CaseSource') AND (CLASSNO = a.CaseSource)) AS CaseSource_C, " + Environment.NewLine +
                                     "       ServiceDate, IsTrue " + Environment.NewLine +
                                     "  FROM CustomService a " + Environment.NewLine +
                                     " WHERE isnull(ServiceNo, '') <> '' " + Session["PriineAllSelected"].ToString().Trim();
                        break;
                }
                sdsCustomServiceDetailP.SelectCommand = "";
                sdsCustomServiceDetailP.SelectCommand = vSelectStr;
                sdsCustomServiceDetailP.DataBind();
                rdsPrint.DataSourceId = sdsCustomServiceDetailP.ID;
                //預覽畫面設定
                rdsPrint.Name = "CustomServiceMain";

                // 把 rvPrint 狀態重設
                rvPrint.LocalReport.DataSources.Clear(); //將 rvPrint 的DataSources集合清除
                rvPrint.Reset(); //將 rvPrint 重置為初始狀態
                rvPrint.LocalReport.Refresh();
                //給予 rvPrint 新的設定值
                rvPrint.LocalReport.ReportPath = @"Report\CustomServiceDetail.rdlc";  // 用 rbGroupBy 的 Value 設定 ReportPath
                rvPrint.LocalReport.DataSources.Add(rdsPrint); //指定新的資料來源
                rvPrint.LocalReport.Refresh(); //重新整理預覽畫面
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}