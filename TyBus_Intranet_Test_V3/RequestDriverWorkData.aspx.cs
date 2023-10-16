using Amaterasu_Function;
using System;
using System.Data;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class RequestDriverWorkData : System.Web.UI.Page
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
        private string vConnStr = "";

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

        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDriverNo = eDriver_Search.Text.Trim();
            string vDriverName = "";
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vDriverNo + "' and LeaveDay is null and Type = '20' ";
            vDriverName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriverNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null and Type = '20' ";
                vDriverNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                if (vDriverNo == "")
                {
                    vDriverName = "";
                }
            }
            eDriver_Search.Text = vDriverNo;
            eDriverName_Search.Text = vDriverName;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            string vSQLStr = "";
            string vDepNo_Driver = "";
            string vDriverNo = eDriver_Search.Text.Trim();
            string vCalYM = eDriveYM_Search.Text.Trim();
            if (vCalYM != "")
            {
                if (vDriverNo != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (vLoginDepNo == "")
                    {
                        vSQLStr = "select DepNo from Employee where EmpNo = '" + vLoginID + "' ";
                        vLoginDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
                    }
                    vSQLStr = "select DepNo from Employee where EmpNo = '" + vDriverNo + "' ";
                    vDepNo_Driver = PF.GetValue(vConnStr, vSQLStr, "DepNo");
                    //登入人員是總公司 (01) 總務課 (02) 主計課 (03) 營運課 (05) 電腦課 (09) 的話就不分站別都可以查，否則只能查自己站上的駕駛
                    if (((vLoginDepNo != "01") && (vLoginDepNo != "02") && (vLoginDepNo != "03") && (vLoginDepNo != "05") && (vLoginDepNo != "09")) && (vLoginDepNo != vDepNo_Driver))
                    {
                        string vRecordNote = "查詢資料_駕駛員公里數及津貼查詢" + Environment.NewLine +
                                             "RequestDriverWorkData.aspx" + Environment.NewLine +
                                             "行車年月：" + vCalYM + Environment.NewLine +
                                             "駕駛員編號：" + vDriverNo + Environment.NewLine +
                                             "查詢結果：錯誤--指定駕駛員非本站駕駛";
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('指定駕駛員非本站駕駛，請重新查詢！')");
                        Response.Write("</" + "Script>");
                        eDriver_Search.Text = "";
                        eDriverName_Search.Text = "";
                        eDriver_Search.Focus();
                    }
                    else
                    {
                        //班車投現載客獎金
                        string vCashBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CashBoundsR'", "Content");
                        vCashBoundsR = (vCashBoundsR != "") ? vCashBoundsR : "0.0";
                        //公車投現載客獎金
                        string vCashBoundsB = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CashBoundsB'", "Content");
                        vCashBoundsB = (vCashBoundsB != "") ? vCashBoundsB : "0.0";
                        //班車普卡載客獎金
                        string vNormalBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_NormalBoundsR'", "Content");
                        vNormalBoundsR = (vNormalBoundsR != "") ? vNormalBoundsR : "0.0";
                        //公車普卡載客獎金
                        string vNormalBoundsB = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_NormalBoundsB'", "Content");
                        vNormalBoundsB = (vNormalBoundsB != "") ? vNormalBoundsB : "0.0";
                        //交通車載客獎金
                        string vContractBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_ContractBoundsR'", "Content");
                        vContractBoundsR = (vContractBoundsR != "") ? vContractBoundsR : "0.0";
                        //國道投現獎金
                        string vHighwayCashR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_HighwayCashR'", "Content");
                        vHighwayCashR = (vHighwayCashR != "") ? vHighwayCashR : "0.0";
                        //班車學生卡載客獎金
                        string vStudentBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_StudentBoundsR'", "Content");
                        vStudentBoundsR = (vStudentBoundsR != "") ? vStudentBoundsR : "0.0";
                        //公車學生卡載客獎金
                        string vStudentBoundsB = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_StudentBoundsB'", "Content");
                        vStudentBoundsB = (vStudentBoundsB != "") ? vStudentBoundsB : "0.0";
                        //班車敬老愛心卡獎金
                        string vOlderBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_OlderBoundsR'", "Content");
                        vOlderBoundsR = (vOlderBoundsR != "") ? vOlderBoundsR : "0.0";
                        //公車敬老愛心卡獎金
                        string vOlderBoundsB = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_OlderBoundsB'", "Content");
                        vOlderBoundsB = (vOlderBoundsR != "") ? vOlderBoundsB : "0.0";
                        //班車偏遠路線獎金
                        string vRuleLinesBounds = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_RuleLinesBounds'", "Content");
                        vRuleLinesBounds = (vRuleLinesBounds != "") ? vRuleLinesBounds : "0.0";
                        //公車偏遠路線獎金
                        string vBusLinesBounds = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_BusLinesBounds'", "Content");
                        vBusLinesBounds = (vBusLinesBounds != "") ? vBusLinesBounds : "0.0";
                        //班車公里獎金
                        string vRuleBounds = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_RuleBounds'", "Content");
                        vRuleBounds = (vRuleBounds != "") ? vRuleBounds : "0.0";
                        //專車公里獎金
                        string vSPECBoundsKM = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_SPECBoundsKM'", "Content");
                        vSPECBoundsKM = (vSPECBoundsKM != "") ? vSPECBoundsKM : "0.0";
                        //公車公里獎金
                        string vBusBounds = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_BusBounds'", "Content");
                        vBusBounds = (vBusBounds != "") ? vBusBounds : "0.0";
                        //聯營公里獎金
                        string vUnionKMBounds = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_UnionKMBounds'", "Content");
                        vUnionKMBounds = (vUnionKMBounds != "") ? vUnionKMBounds : "0.0";
                        //交通車公里獎金
                        string vTransBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TransBoundsR'", "Content");
                        vTransBoundsR = (vTransBoundsR != "") ? vTransBoundsR : "0.0";
                        //區間租車公里獎金
                        string vRentBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_RentBoundsR'", "Content");
                        vRentBoundsR = (vRentBoundsR != "") ? vRentBoundsR : "0.0";
                        //遊覽車公里獎金
                        string vTourBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_TourBoundsR'", "Content");
                        vTourBoundsR = (vTourBoundsR != "") ? vTourBoundsR : "0.0";
                        //非營運公里獎金
                        string vNoneBusiBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_NoneBusiBoundsR'", "Content");
                        vNoneBusiBoundsR = (vNoneBusiBoundsR != "") ? vNoneBusiBoundsR : "0.0";
                        //班車借中巴
                        string vMiniBusBoundsR = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_MiniBusBoundsR'", "Content");
                        vMiniBusBoundsR = (vMiniBusBoundsR != "") ? vMiniBusBoundsR : "0.0";
                        //公車借中巴
                        string vMiniBusBoundsB = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_MiniBusBoundsB'", "Content");
                        vMiniBusBoundsB = (vMiniBusBoundsB != "") ? vMiniBusBoundsB : "0.0";
                        //租車趟次津貼
                        string vRentTimesBounds = PF.GetValue(vConnStr, "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_RentTimesBounds'", "Content");
                        vRentTimesBounds = (vRentTimesBounds != "") ? vRentTimesBounds : "0.0";
                        //起迄日期
                        DateTime vStartDate;
                        vStartDate = DateTime.Parse(vCalYM.Substring(0, 4) + "/" + vCalYM.Substring(4, 2) + "/01");
                        DateTime vEndDate;
                        vEndDate = DateTime.Parse(PF.GetMonthLastDay(vStartDate, "B"));
                        try
                        {
                            sdsDriverData.SelectParameters.Clear();
                            sdsDriverData.SelectParameters.Add(new Parameter("CashBoundsR", DbType.String, vCashBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("CashBoundsB", DbType.String, vCashBoundsB));
                            sdsDriverData.SelectParameters.Add(new Parameter("NormalBoundsR", DbType.String, vNormalBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("NormalBoundsB", DbType.String, vNormalBoundsB));
                            sdsDriverData.SelectParameters.Add(new Parameter("ContractBoundsR", DbType.String, vContractBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("HighwayCashR", DbType.String, vHighwayCashR));
                            sdsDriverData.SelectParameters.Add(new Parameter("StudentBoundsR", DbType.String, vStudentBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("StudentBoundsB", DbType.String, vStudentBoundsB));
                            sdsDriverData.SelectParameters.Add(new Parameter("OlderBoundsR", DbType.String, vOlderBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("OlderBoundsB", DbType.String, vOlderBoundsB));
                            sdsDriverData.SelectParameters.Add(new Parameter("RuleLinesBounds", DbType.String, vRuleLinesBounds));
                            sdsDriverData.SelectParameters.Add(new Parameter("BusLinesBounds", DbType.String, vBusLinesBounds));
                            sdsDriverData.SelectParameters.Add(new Parameter("RuleBounds", DbType.String, vRuleBounds));
                            sdsDriverData.SelectParameters.Add(new Parameter("SPECBoundsKM", DbType.String, vSPECBoundsKM));
                            sdsDriverData.SelectParameters.Add(new Parameter("BusBounds", DbType.String, vBusBounds));
                            sdsDriverData.SelectParameters.Add(new Parameter("UnionKMBounds", DbType.String, vUnionKMBounds));
                            sdsDriverData.SelectParameters.Add(new Parameter("TransBoundsR", DbType.String, vTransBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("RentBoundsR", DbType.String, vRentBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("TourBoundsR", DbType.String, vTourBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("NoneBusiBoundsR", DbType.String, vNoneBusiBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("MiniBusBoundsR", DbType.String, vMiniBusBoundsR));
                            sdsDriverData.SelectParameters.Add(new Parameter("MiniBusBoundsB", DbType.String, vMiniBusBoundsB));
                            sdsDriverData.SelectParameters.Add(new Parameter("RentTimesBounds", DbType.String, vRentTimesBounds));
                            sdsDriverData.SelectParameters.Add(new Parameter("EmpNo", DbType.String, vDriverNo));
                            sdsDriverData.SelectParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_Driver));
                            sdsDriverData.SelectParameters.Add(new Parameter("BuDate", DbType.Date, vStartDate.ToString("yyyy/MM/dd")));
                            sdsDriverData.SelectParameters.Add(new Parameter("EndDate", DbType.Date, vEndDate.ToString("yyyy/MM/dd")));

                            string vRecordNote = "查詢資料_駕駛員公里數及津貼查詢" + Environment.NewLine +
                                                 "RequestDriverWorkData.aspx" + Environment.NewLine +
                                                 "行車年月：" + vCalYM + Environment.NewLine +
                                                 "駕駛員編號：" + vDriverNo + Environment.NewLine +
                                                 "查詢結果：正常";
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                        }
                        catch (Exception eMessage)
                        {
                            string vRecordNote = "查詢資料_駕駛員公里數及津貼查詢" + Environment.NewLine +
                                                 "RequestDriverWorkData.aspx" + Environment.NewLine +
                                                 "行車年月：" + vCalYM + Environment.NewLine +
                                                 "駕駛員編號：" + vDriverNo + Environment.NewLine +
                                                 "查詢結果：錯誤--" + eMessage.Message.Trim() + eMessage.ToString();
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                            Response.Write("</" + "Script>");
                            //throw;
                        }
                    }
                }
                else
                {
                    string vRecordNote = "查詢資料_駕駛員公里數及津貼查詢" + Environment.NewLine +
                                         "RequestDriverWorkData.aspx" + Environment.NewLine +
                                         "行車年月：" + vCalYM + Environment.NewLine +
                                         "駕駛員編號：" + vDriverNo + Environment.NewLine +
                                         "查詢結果：錯誤--未指定駕駛員工號";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('請輸入要查詢的駕駛員編號！')");
                    Response.Write("</" + "Script>");
                    eDriver_Search.Focus();
                }
            }
            else
            {
                string vRecordNote = "查詢資料_駕駛員公里數及津貼查詢" + Environment.NewLine +
                                     "RequestDriverWorkData.aspx" + Environment.NewLine +
                                     "行車年月：" + vCalYM + Environment.NewLine +
                                     "駕駛員編號：" + vDriverNo + Environment.NewLine +
                                     "查詢結果：錯誤--未指定行車年月";
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請輸入行車年月！')");
                Response.Write("</" + "Script>");
                eDriveYM_Search.Focus();
            }
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}