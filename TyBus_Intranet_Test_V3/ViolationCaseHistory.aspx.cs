using Amaterasu_Function;
using System;
using System.Globalization;

namespace TyBus_Intranet_Test_V3
{
    public partial class ViolationCaseHistory : System.Web.UI.Page
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
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                System.Threading.Thread.CurrentThread.CurrentCulture = Cal;
                //
                DateTime vToday = DateTime.Today;
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
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    string vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_Start.ClientID;
                    string vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuildDate_Start.Attributes["onClick"] = vTempDateScript;

                    vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_End.ClientID;
                    vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuildDate_End.Attributes["onClick"] = vTempDateScript;

                    vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eViolationDate_Start.ClientID;
                    vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eViolationDate_Start.Attributes["onClick"] = vTempDateScript;

                    vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eViolationDate_End.ClientID;
                    vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eViolationDate_End.Attributes["onClick"] = vTempDateScript;

                    CaseDataBind();
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

        private void CaseDataBind()
        {
            string vSQLStr = "";
            string vWStr_BuildDate = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ?
                                     "   and a.BuildDate between '" + DateTime.Parse(eBuildDate_Start.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBuildDate_Start.Text.Trim()).ToString("MM/dd") + " and '" + DateTime.Parse(eBuildDate_End.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBuildDate_End.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ?
                                     "   and a.BuildDate = '" + DateTime.Parse(eBuildDate_Start.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBuildDate_Start.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ?
                                     "   and a.BuildDate = '" + DateTime.Parse(eBuildDate_End.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBuildDate_End.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_ViolationDate = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ?
                                         "   and a.ViolationDate between '" + DateTime.Parse(eViolationDate_Start.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_Start.Text.Trim()).ToString("MM/dd") + " and '" + DateTime.Parse(eViolationDate_End.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_End.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ?
                                         "   and a.ViolationDate = '" + DateTime.Parse(eViolationDate_Start.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_Start.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ?
                                         "   and a.ViolationDate = '" + DateTime.Parse(eViolationDate_End.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_End.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepNo_Start.Text.Trim() + "' and '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? "   and a.DepNo = '" + eDepNo_Start.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? "   and a.DepNo = '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver.Text.Trim() != "") ? "   and a.Driver = '" + eDriver.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CarID = (eCar_ID.Text.Trim() != "") ? "   and a.Car_ID = '" + eCar_ID.Text.Trim() + "' " + Environment.NewLine : "";
            vSQLStr = "SELECT HistoryNo, CaseNo, CaseType, " + Environment.NewLine +
                      "       (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.CaseType) AND (FKEY = CAST('違規記錄檔' AS char(16)) + CAST('ViolationCase' AS char(16)) + 'CaseType')) AS CaseType_C, " + Environment.NewLine +
                      "       BuildDate, BuildMan, BuildManName, LinesNo, DepNo, DepName, Car_ID, Driver, DriverName, ViolationDate, TicketNo, TicketTitle, " + Environment.NewLine +
                      "       PenaltyDep, Undertaker, ViolationLocation, ViolationRule, ViolationNote, FineAmount, PenaltyReason, ViolationPoint, " + Environment.NewLine +
                      "       PaymentDeadline, PaidDate, Remark, ModifyType, " + Environment.NewLine +
                      "       (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (CLASSNO = a.ModifyType) AND (FKEY = CAST('客訴單歷史' AS char(16)) + CAST('fmServiceHistory' AS char(16)) + 'ModifyType')) AS ModifyType_C, " + Environment.NewLine +
                      "       ModifyDate, ModifyMan, " + Environment.NewLine +
                      "       (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.ModifyMan)) AS ModifyMan_C " + Environment.NewLine +
                      "  FROM ViolationCaseHistory AS a" + Environment.NewLine +
                      " where 1 = 1 " + Environment.NewLine +
                      vWStr_BuildDate +
                      vWStr_ViolationDate +
                      vWStr_DepNo +
                      vWStr_Driver +
                      vWStr_CarID +
                      " order by a.BuildDate DESC";
            sdsViolationCaseHistoryList.SelectCommand = "";
            sdsViolationCaseHistoryList.SelectCommand = vSQLStr;
            gridViolationCaseHistory.DataBind();
        }

        protected void eDepNo_Start_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDepNo_Start.Text.Trim();
            string vDepName_Temp = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Start.Text = vDepNo_Temp;
            eDepName_Start.Text = vDepName_Temp;
        }

        protected void eDepNo_End_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDepNo_End.Text.Trim();
            string vDepName_Temp = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_End.Text = vDepNo_Temp;
            eDepName_End.Text = vDepName_Temp;
        }

        protected void eDriver_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo = eDriver.Text.Trim();
            string vEmpName = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from EMployee where EmpNo = '" + vEmpNo + "' and LeaveDay is null";
            vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select EmpNo from EMployee where [Name] = '" + vEmpName + "' and LeaveDay is null ";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eDriver.Text = vEmpNo;
            eDriverName.Text = vEmpName;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            CaseDataBind();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}