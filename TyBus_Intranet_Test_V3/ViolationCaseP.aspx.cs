using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class ViolationCaseP : System.Web.UI.Page
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

                    if (!IsPostBack)
                    {
                        plReport.Visible = false;
                    }
                    else
                    {
                        ListDataBind();
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

        private void ListDataBind()
        {
            string vSelectStr = "";
            switch (rbPrintMode.SelectedValue)
            {
                case "11":
                    vSelectStr = CaseDataBind_DepNo("V01");
                    gvShowData_DepNo.Visible = (gvShowData_DepNo.Rows.Count > 0);
                    bbPrint.Visible = (gvShowData_DepNo.Rows.Count > 0);
                    bbExcel.Visible = (gvShowData_DepNo.Rows.Count > 0);
                    break;
                case "12":
                    vSelectStr = CaseDataBind_Driver();
                    gvShowData_Driver.Visible = (gvShowData_Driver.Rows.Count > 0);
                    bbPrint.Visible = (gvShowData_Driver.Rows.Count > 0);
                    bbExcel.Visible = (gvShowData_Driver.Rows.Count > 0);
                    break;
                case "13":
                    vSelectStr = CaseDataBind_DepNo("V02");
                    gvShowData_DepNo.Visible = (gvShowData_DepNo.Rows.Count > 0);
                    bbPrint.Visible = (gvShowData_DepNo.Rows.Count > 0);
                    bbExcel.Visible = (gvShowData_DepNo.Rows.Count > 0);
                    break;
                case "21":
                    vSelectStr = CaseDataBind_Point();
                    gvShowData_Point.Visible = (gvShowData_Point.Rows.Count > 0);
                    bbPrint.Visible = (gvShowData_Point.Rows.Count > 0);
                    bbExcel.Visible = (gvShowData_Point.Rows.Count > 0);
                    break;
                case "99":
                    vSelectStr = CaseDataBind_99();
                    gvShowData_99.Visible = (gvShowData_99.Rows.Count > 0);
                    bbPrint.Visible = (gvShowData_99.Rows.Count > 0);
                    bbExcel.Visible = (gvShowData_99.Rows.Count > 0);
                    break;
            }
        }

        private string CaseDataBind_DepNo(string fCaseType)
        {
            string vSQLStr = "";
            string vSelectStr = "";
            sdsViolationCaseP_DepNo.SelectCommand = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vWStr_BuildDate = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ? "   and v.BuildDate between '" + PF.TransDateString(eBuildDate_Start.Text.Trim(), "B") + "' and '" + PF.TransDateString(eBuildDate_End.Text.Trim(), "B") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ? "   and v.BuildDate = '" + PF.TransDateString(eBuildDate_Start.Text.Trim(), "B") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ? "   and v.BuildDate = '" + PF.TransDateString(eBuildDate_End.Text.Trim(), "B") + "' " + Environment.NewLine : "";
            string vWStr_ViolationDate = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ? "   and v.ViolationDate between '" + PF.TransDateString(eViolationDate_Start.Text.Trim(), "B") + "' and '" + PF.TransDateString(eViolationDate_End.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ? "   and v.ViolationDate = '" + PF.TransDateString(eViolationDate_Start.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ? "   and v.ViolationDate = '" + PF.TransDateString(eViolationDate_End.Text.Trim(), "B") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "   and v.DepNo between '" + eDepNo_Start.Text.Trim() + "' and '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? "   and v.DepNo = '" + eDepNo_Start.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? "   and v.DepNo = '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver.Text.Trim() != "") ? "   and v.Driver = '" + eDriver.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Car_ID = (eCar_ID.Text.Trim() != "") ? "   and v.Car_ID = '" + eCar_ID.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ViolationRule = (eViolationRule_Search.Text.Trim() != "") ? "   and v.ViolationRule like '%" + eViolationRule_Search.Text.Trim() + "%'" + Environment.NewLine : "";

            vSQLStr = "SELECT CaseNo, CaseType, " + Environment.NewLine +
                      "       (SELECT CLASSTXT FROM DBDICB WHERE(CLASSNO = v.CaseType) AND(FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, " + Environment.NewLine +
                      "       DepNo, DepName, ViolationDate, Driver, DriverName, ViolationRule, ViolationPoint, ViolationNote " + Environment.NewLine +
                      "  FROM ViolationCase AS v" + Environment.NewLine +
                      " where CaseType = '" + fCaseType.Trim() + "' ";

            vSelectStr = vSQLStr +
                         vWStr_BuildDate +
                         vWStr_ViolationDate +
                         vWStr_DepNo +
                         vWStr_Driver +
                         vWStr_Car_ID +
                         vWStr_ViolationRule +
                         " order by v.DepNo, v.ViolationDate ";
            sdsViolationCaseP_DepNo.SelectCommand = vSelectStr;
            gvShowData_DepNo.DataBind();
            return vSelectStr;
        }

        private string CaseDataBind_Driver()
        {
            string vSelectStr = "";
            string vSQLStr = "";
            sdsViolationCaseP_Driver.SelectCommand = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vWStr_BuildDate = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ? "   and v.BuildDate between '" + PF.TransDateString(eBuildDate_Start.Text.Trim(), "B") + "' and '" + PF.TransDateString(eBuildDate_End.Text.Trim(), "B") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ? "   and v.BuildDate = '" + PF.TransDateString(eBuildDate_Start.Text.Trim(), "B") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ? "   and v.BuildDate = '" + PF.TransDateString(eBuildDate_End.Text.Trim(), "B") + "' " + Environment.NewLine : "";
            string vWStr_ViolationDate = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ? "   and v.ViolationDate between '" + PF.TransDateString(eViolationDate_Start.Text.Trim(), "B") + "' and '" + PF.TransDateString(eViolationDate_End.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ? "   and v.ViolationDate = '" + PF.TransDateString(eViolationDate_Start.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ? "   and v.ViolationDate = '" + PF.TransDateString(eViolationDate_End.Text.Trim(), "B") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "   and v.DepNo between '" + eDepNo_Start.Text.Trim() + "' and '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? "   and v.DepNo = '" + eDepNo_Start.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? "   and v.DepNo = '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver.Text.Trim() != "") ? "   and v.Driver = '" + eDriver.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Car_ID = (eCar_ID.Text.Trim() != "") ? "   and v.Car_ID = '" + eCar_ID.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ViolationRule = (eViolationRule_Search.Text.Trim() != "") ? "   and v.ViolationRule like '%" + eViolationRule_Search.Text.Trim() + "%'" + Environment.NewLine : "";

            vSQLStr = "SELECT v.CaseNo, v.CaseType, v.DepNo, v.DepName, v.Car_ID, " + Environment.NewLine +
                      "       (SELECT CLASSTXT FROM DBDICB WHERE(CLASSNO = v.CaseType) AND(FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, " + Environment.NewLine +
                      "       ViolationDate, Driver, DriverName, ViolationRule, ViolationPoint, ViolationNote " + Environment.NewLine +
                      "  FROM ViolationCase AS v" + Environment.NewLine +
                      " where 1 = 1 ";

            vSelectStr = vSQLStr +
                         vWStr_BuildDate +
                         vWStr_ViolationDate +
                         vWStr_DepNo +
                         vWStr_Driver +
                         vWStr_Car_ID +
                         vWStr_ViolationRule +
                         " order by v.Driver, v.ViolationDate ";
            sdsViolationCaseP_Driver.SelectCommand = vSelectStr;
            gvShowData_Driver.DataBind();
            return vSelectStr;
        }

        private string CaseDataBind_Point()
        {
            string vSelectStr = "";
            string vSQLStr = "";
            string vCalDate = PF.GetMonthFirstDay(DateTime.Today, "B");
            sdsViolationCaseP_Point.SelectCommand = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vWStr_BuildDate = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ? "   and BuildDate between '" + PF.TransDateString(eBuildDate_Start.Text.Trim(), "B") + "' and '" + PF.TransDateString(eBuildDate_End.Text.Trim(), "B") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ? "   and BuildDate = '" + PF.TransDateString(eBuildDate_Start.Text.Trim(), "B") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ? "   and BuildDate = '" + PF.TransDateString(eBuildDate_End.Text.Trim(), "B") + "' " + Environment.NewLine : "";
            string vWStr_ViolationDate = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ? "   and ViolationDate between '" + PF.TransDateString(eViolationDate_Start.Text.Trim(), "B") + "' and '" + PF.TransDateString(eViolationDate_End.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ? "   and ViolationDate = '" + PF.TransDateString(eViolationDate_Start.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ? "   and ViolationDate = '" + PF.TransDateString(eViolationDate_End.Text.Trim(), "B") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "   and DepNo between '" + eDepNo_Start.Text.Trim() + "' and '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? "   and DepNo = '" + eDepNo_Start.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? "   and DepNo = '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver.Text.Trim() != "") ? "   and Driver = '" + eDriver.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Car_ID = (eCar_ID.Text.Trim() != "") ? "   and Car_ID = '" + eCar_ID.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ViolationRule = (eViolationRule_Search.Text.Trim() != "") ? "   and ViolationRule like '%" + eViolationRule_Search.Text.Trim() + "%'" : "";

            vSQLStr = "select CaseNo, CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.CaseType) AND (FKEY = '違規記錄檔      ViolationCase   CaseType')) CaseType_C, " + Environment.NewLine +
                      "       DepNo, DepName, ViolationDate, DriverName, ViolationRule, ViolationPoint, ViolationNote, " + Environment.NewLine +
                      "       case when DateDiff(month, ViolationDate, '" + vCalDate + "') <= 6 then 'V' else '' end IsAffectivePoint " + Environment.NewLine +
                      "  from ViolationCase a " + Environment.NewLine +
                      " where CaseType = 'V01' ";

            vSelectStr = vSQLStr +
                        vWStr_BuildDate +
                        vWStr_ViolationDate +
                        vWStr_DepNo +
                        vWStr_Driver +
                        vWStr_Car_ID +
                        vWStr_ViolationRule +
                        " order by DepNo, ViolationDate ";
            sdsViolationCaseP_Point.SelectCommand = vSelectStr;
            gvShowData_Point.DataBind();
            return vSelectStr;
        }

        private string CaseDataBind_99()
        {
            string vSelectStr = "";
            string vSQLStr = "";
            sdsViolationCaseP_99.SelectCommand = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vWStr_BuildDate = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ? "   and BuildDate between '" + PF.TransDateString(eBuildDate_Start.Text.Trim(), "B") + "' and '" + PF.TransDateString(eBuildDate_End.Text.Trim(), "B") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ? "   and BuildDate = '" + PF.TransDateString(eBuildDate_Start.Text.Trim(), "B") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ? "   and BuildDate = '" + PF.TransDateString(eBuildDate_End.Text.Trim(), "B") + "' " + Environment.NewLine : "";
            string vWStr_ViolationDate = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ? "   and ViolationDate between '" + PF.TransDateString(eViolationDate_Start.Text.Trim(), "B") + "' and '" + PF.TransDateString(eViolationDate_End.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ? "   and ViolationDate = '" + PF.TransDateString(eViolationDate_Start.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ? "   and ViolationDate = '" + PF.TransDateString(eViolationDate_End.Text.Trim(), "B") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "   and DepNo between '" + eDepNo_Start.Text.Trim() + "' and '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? "   and DepNo = '" + eDepNo_Start.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? "   and DepNo = '" + eDepNo_End.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver.Text.Trim() != "") ? "   and Driver = '" + eDriver.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Car_ID = (eCar_ID.Text.Trim() != "") ? "   and Car_ID = '" + eCar_ID.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ViolationRule = (eViolationRule_Search.Text.Trim() != "") ? "   and ViolationRule like '%" + eViolationRule_Search.Text.Trim() + "%'" : "";

            vSQLStr = "select CaseNo, " + Environment.NewLine +
                      "       (select count(CaseNo) from ViolationCase where CaseType = a.CaseType " + vWStr_BuildDate + vWStr_Car_ID + vWStr_DepNo + vWStr_Driver + vWStr_ViolationDate + ") as CountByType, " + Environment.NewLine +
                      "       (select count(CaseNo) from ViolationCase where DepNo = a.DepNo and CaseType = a.CaseType " + vWStr_BuildDate + vWStr_Car_ID + vWStr_DepNo + vWStr_Driver + vWStr_ViolationDate + ") as CountByDep, " + Environment.NewLine +
                      "       CaseType, " + Environment.NewLine +
                      "       (select ClassTxt from DBDICB where ClassNo = a.CaseType and FKey = '違規記錄檔      ViolationCase   CaseType') as CaseType_C, " + Environment.NewLine +
                      "       DepNo, DepName, LinesNo, Driver, DriverName, ViolationDate, ViolationRule, FineAmount, ViolationPoint, ViolationNote, TicketTitle " + Environment.NewLine +
                      "  from ViolationCase a " + Environment.NewLine +
                      " where 1 = 1";

            vSelectStr = vSQLStr +
                        vWStr_BuildDate +
                        vWStr_ViolationDate +
                        vWStr_DepNo +
                        vWStr_Driver +
                        vWStr_Car_ID +
                        vWStr_ViolationRule +
                        " order by DepNo, ViolationDate ";
            sdsViolationCaseP_99.SelectCommand = vSelectStr;
            gvShowData_99.DataBind();
            return vSelectStr;
        }

        private void PrintReport_DepNo(string fSelStr)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintDepNo = new SqlDataAdapter(fSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintDepNo = new DataTable("ViolationCaseP");
                daPrintDepNo.Fill(dtPrintDepNo);
                if (dtPrintDepNo.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("ViolationCaseP", dtPrintDepNo);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\ViolationCaseP_DepNo.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("StartDate", eViolationDate_Start.Text.Trim()));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("EndDate", eViolationDate_End.Text.Trim()));
                    rvPrint.LocalReport.Refresh();
                    plReport.Visible = true;
                    plShowData.Visible = false;

                    string vPrintModeStr = rbPrintMode.SelectedItem.Text.Trim();
                    string vBuildDateStr = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ? "從 " + eBuildDate_Start.Text.Trim() + " 起至 " + eBuildDate_End.Text.Trim() + " 止" :
                                           ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ? eBuildDate_Start.Text.Trim() :
                                           ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ? eBuildDate_End.Text.Trim() : "";
                    string vViolationDateStr = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ? "從 " + eViolationDate_Start.Text.Trim() + " 起至 " + eViolationDate_End.Text.Trim() + " 止" :
                                               ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ? eViolationDate_Start.Text.Trim() :
                                               ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ? eViolationDate_End.Text.Trim() : "";
                    string vDepNoStr = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "從 " + eDepNo_Start.Text.Trim() + " 起至 " + eDepNo_End.Text.Trim() + " 止" :
                                       ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? eDepNo_Start.Text.Trim() :
                                       ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? eDepNo_End.Text.Trim() : "全部";
                    string vDriverStr = (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : "";
                    string vCarIDStr = (eCar_ID.Text.Trim() != "") ? eCar_ID.Text.Trim() : "";
                    string vViolationRuleStr = (eViolationRule_Search.Text.Trim() != "") ? eViolationRule_Search.Text.Trim() : "";
                    string vRecordNote = "預覽報表_違規案件統計表" + Environment.NewLine +
                                         "ViolationCaseP.aspx" + Environment.NewLine +
                                         "報表種類：" + vPrintModeStr + Environment.NewLine +
                                         "建檔日期：" + vBuildDateStr + Environment.NewLine +
                                         "違規日期："+ vViolationDateStr+Environment.NewLine+
                                         "站別："+ vDepNoStr+Environment.NewLine+
                                         "駕駛員：" + vDriverStr + Environment.NewLine +
                                         "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                         "違規條文：" + vViolationRuleStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        private void PrintReport_Driver(string fSelStr)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintDepNo = new SqlDataAdapter(fSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintDepNo = new DataTable("ViolationCaseP");
                daPrintDepNo.Fill(dtPrintDepNo);
                if (dtPrintDepNo.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("ViolationCaseP", dtPrintDepNo);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\ViolationCaseP_Driver.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("StartDate", eViolationDate_Start.Text.Trim()));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("EndDate", eViolationDate_End.Text.Trim()));
                    rvPrint.LocalReport.Refresh();
                    plReport.Visible = true;
                    plShowData.Visible = false;

                    string vPrintModeStr = rbPrintMode.SelectedItem.Text.Trim();
                    string vBuildDateStr = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ? "從 " + eBuildDate_Start.Text.Trim() + " 起至 " + eBuildDate_End.Text.Trim() + " 止" :
                                           ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ? eBuildDate_Start.Text.Trim() :
                                           ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ? eBuildDate_End.Text.Trim() : "";
                    string vViolationDateStr = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ? "從 " + eViolationDate_Start.Text.Trim() + " 起至 " + eViolationDate_End.Text.Trim() + " 止" :
                                               ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ? eViolationDate_Start.Text.Trim() :
                                               ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ? eViolationDate_End.Text.Trim() : "";
                    string vDepNoStr = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "從 " + eDepNo_Start.Text.Trim() + " 起至 " + eDepNo_End.Text.Trim() + " 止" :
                                       ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? eDepNo_Start.Text.Trim() :
                                       ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? eDepNo_End.Text.Trim() : "全部";
                    string vDriverStr = (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : "";
                    string vCarIDStr = (eCar_ID.Text.Trim() != "") ? eCar_ID.Text.Trim() : "";
                    string vViolationRuleStr = (eViolationRule_Search.Text.Trim() != "") ? eViolationRule_Search.Text.Trim() : "";
                    string vRecordNote = "預覽報表_違規案件統計表" + Environment.NewLine +
                                         "ViolationCaseP.aspx" + Environment.NewLine +
                                         "報表種類：" + vPrintModeStr + Environment.NewLine +
                                         "建檔日期：" + vBuildDateStr + Environment.NewLine +
                                         "違規日期：" + vViolationDateStr + Environment.NewLine +
                                         "站別：" + vDepNoStr + Environment.NewLine +
                                         "駕駛員：" + vDriverStr + Environment.NewLine +
                                         "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                         "違規條文：" + vViolationRuleStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        private void PrintReport_Point(string fSelStr)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintPoint = new SqlDataAdapter(fSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintPoint = new DataTable("ViolationCaseP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("ViolationCaseP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\ViolationCaseP_Point.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("StartDate", eViolationDate_Start.Text.Trim()));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("EndDate", eViolationDate_End.Text.Trim()));
                    rvPrint.LocalReport.Refresh();
                    plReport.Visible = true;
                    plShowData.Visible = false;

                    string vPrintModeStr = rbPrintMode.SelectedItem.Text.Trim();
                    string vBuildDateStr = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ? "從 " + eBuildDate_Start.Text.Trim() + " 起至 " + eBuildDate_End.Text.Trim() + " 止" :
                                           ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ? eBuildDate_Start.Text.Trim() :
                                           ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ? eBuildDate_End.Text.Trim() : "";
                    string vViolationDateStr = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ? "從 " + eViolationDate_Start.Text.Trim() + " 起至 " + eViolationDate_End.Text.Trim() + " 止" :
                                               ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ? eViolationDate_Start.Text.Trim() :
                                               ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ? eViolationDate_End.Text.Trim() : "";
                    string vDepNoStr = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "從 " + eDepNo_Start.Text.Trim() + " 起至 " + eDepNo_End.Text.Trim() + " 止" :
                                       ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? eDepNo_Start.Text.Trim() :
                                       ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? eDepNo_End.Text.Trim() : "全部";
                    string vDriverStr = (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : "";
                    string vCarIDStr = (eCar_ID.Text.Trim() != "") ? eCar_ID.Text.Trim() : "";
                    string vViolationRuleStr = (eViolationRule_Search.Text.Trim() != "") ? eViolationRule_Search.Text.Trim() : "";
                    string vRecordNote = "預覽報表_違規案件統計表" + Environment.NewLine +
                                         "ViolationCaseP.aspx" + Environment.NewLine +
                                         "報表種類：" + vPrintModeStr + Environment.NewLine +
                                         "建檔日期：" + vBuildDateStr + Environment.NewLine +
                                         "違規日期：" + vViolationDateStr + Environment.NewLine +
                                         "站別：" + vDepNoStr + Environment.NewLine +
                                         "駕駛員：" + vDriverStr + Environment.NewLine +
                                         "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                         "違規條文：" + vViolationRuleStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        private void PrintReport_99(string fSelStr)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintDepNo = new SqlDataAdapter(fSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintDepNo = new DataTable("ViolationCaseP");
                daPrintDepNo.Fill(dtPrintDepNo);
                if (dtPrintDepNo.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("ViolationCaseP", dtPrintDepNo);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\ViolationCaseP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("StartDate", eViolationDate_Start.Text.Trim()));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("EndDate", eViolationDate_End.Text.Trim()));
                    rvPrint.LocalReport.Refresh();
                    plReport.Visible = true;
                    plShowData.Visible = false;

                    string vPrintModeStr = rbPrintMode.SelectedItem.Text.Trim();
                    string vBuildDateStr = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ? "從 " + eBuildDate_Start.Text.Trim() + " 起至 " + eBuildDate_End.Text.Trim() + " 止" :
                                           ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ? eBuildDate_Start.Text.Trim() :
                                           ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ? eBuildDate_End.Text.Trim() : "";
                    string vViolationDateStr = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ? "從 " + eViolationDate_Start.Text.Trim() + " 起至 " + eViolationDate_End.Text.Trim() + " 止" :
                                               ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ? eViolationDate_Start.Text.Trim() :
                                               ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ? eViolationDate_End.Text.Trim() : "";
                    string vDepNoStr = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "從 " + eDepNo_Start.Text.Trim() + " 起至 " + eDepNo_End.Text.Trim() + " 止" :
                                       ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? eDepNo_Start.Text.Trim() :
                                       ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? eDepNo_End.Text.Trim() : "全部";
                    string vDriverStr = (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : "";
                    string vCarIDStr = (eCar_ID.Text.Trim() != "") ? eCar_ID.Text.Trim() : "";
                    string vViolationRuleStr = (eViolationRule_Search.Text.Trim() != "") ? eViolationRule_Search.Text.Trim() : "";
                    string vRecordNote = "預覽報表_違規案件統計表" + Environment.NewLine +
                                         "ViolationCaseP.aspx" + Environment.NewLine +
                                         "報表種類：" + vPrintModeStr + Environment.NewLine +
                                         "建檔日期：" + vBuildDateStr + Environment.NewLine +
                                         "違規日期：" + vViolationDateStr + Environment.NewLine +
                                         "站別：" + vDepNoStr + Environment.NewLine +
                                         "駕駛員：" + vDriverStr + Environment.NewLine +
                                         "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                         "違規條文：" + vViolationRuleStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        private void SaveExcel(string fFileName, string fSelectStr)
        {
            DateTime vBuDate;
            HSSFWorkbook wbExcel = new HSSFWorkbook();
            //Excel 工作表
            HSSFSheet wsExcel;
            //設定標題欄位的格式
            HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中

            //設定字體格式
            HSSFFont fontTitle = (HSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 12; //字體大小
            csTitle.SetFont(fontTitle);

            //設定資料內容欄位的格式
            HSSFCellStyle csData = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
            fontData.FontHeightInPoints = 12;
            csData.SetFont(fontData);

            HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
            fontData_Red.FontHeightInPoints = 12; //字體大小
            fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
            csData_Red.SetFont(fontData_Red);

            HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
            csData_Int.DataFormat = format.GetFormat("###,##0");

            string vHeaderText = "";
            int vLinesNo = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            SqlConnection connExcel = new SqlConnection(vConnStr);
            SqlCommand cmdExcel = new SqlCommand(fSelectStr, connExcel);
            connExcel.Open();
            SqlDataReader drExcel = cmdExcel.ExecuteReader();
            if (drExcel.HasRows)
            {
                //查詢結果有資料的時候才執行
                //新增一個工作表
                wsExcel = (HSSFSheet)wbExcel.CreateSheet(fFileName);
                //寫入標題列
                vLinesNo = 0;
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < drExcel.FieldCount; i++)
                {
                    vHeaderText = (drExcel.GetName(i) == "CaseNo") ? "案件編號" :
                                  (drExcel.GetName(i) == "CaseType") ? "案件類別代號" :
                                  (drExcel.GetName(i) == "CaseType_C") ? "案件類別" :
                                  (drExcel.GetName(i) == "DepNo") ? "站別代號" :
                                  (drExcel.GetName(i) == "DepName") ? "站別" :
                                  (drExcel.GetName(i) == "Car_ID") ? "車號" :
                                  (drExcel.GetName(i) == "ViolationDate") ? "違規日期" :
                                  (drExcel.GetName(i) == "Driver") ? "駕駛員代號" :
                                  (drExcel.GetName(i) == "DriverName") ? "駕駛員姓名" :
                                  (drExcel.GetName(i) == "ViolationRule") ? "法條" :
                                  (drExcel.GetName(i) == "ViolationPoint") ? "記 (扣) 點" :
                                  (drExcel.GetName(i) == "ViolationNote") ? "違規事項" :
                                  (drExcel.GetName(i) == "CountByType") ? "案件總數 (依類別)" :
                                  (drExcel.GetName(i) == "CounyByDep") ? "各站總數 (依類別)" :
                                  (drExcel.GetName(i) == "LinesNo") ? "路線" :
                                  (drExcel.GetName(i) == "FineAmount") ? "裁罰金額" :
                                  (drExcel.GetName(i) == "TicketTitle") ? "裁罰主旨" :
                                  (drExcel.GetName(i) == "IsAffectivePoint") ? "有效扣點" : drExcel.GetName(i);
                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                }
                vLinesNo++;
                while (drExcel.Read())
                {
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i);
                        if ((drExcel.GetName(i) == "ViolationPoint") ||
                            (drExcel.GetName(i) == "CountByType") ||
                            (drExcel.GetName(i) == "CounyByDep") ||
                            (drExcel.GetName(i) == "FineAmount"))
                        {
                            if (drExcel[i].ToString() != "")
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(0);
                            }
                        }
                        else if ((drExcel.GetName(i) == "ViolationDate") && (drExcel[i].ToString() != ""))
                        {
                            string vTempStr = drExcel[i].ToString();
                            vBuDate = DateTime.Parse(drExcel[i].ToString());
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 1911).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                        }
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                    }
                    vLinesNo++;
                }
                try
                {
                    var msTarget = new NPOIMemoryStream();
                    msTarget.AllowClose = false;
                    wbExcel.Write(msTarget);
                    msTarget.Flush();
                    msTarget.Seek(0, SeekOrigin.Begin);
                    msTarget.AllowClose = true;

                    if (msTarget.Length > 0)
                    {
                        string vPrintModeStr = rbPrintMode.SelectedItem.Text.Trim();
                        string vBuildDateStr = ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() != "")) ? "從 " + eBuildDate_Start.Text.Trim() + " 起至 " + eBuildDate_End.Text.Trim() + " 止" :
                                               ((eBuildDate_Start.Text.Trim() != "") && (eBuildDate_End.Text.Trim() == "")) ? eBuildDate_Start.Text.Trim() :
                                               ((eBuildDate_Start.Text.Trim() == "") && (eBuildDate_End.Text.Trim() != "")) ? eBuildDate_End.Text.Trim() : "";
                        string vViolationDateStr = ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() != "")) ? "從 " + eViolationDate_Start.Text.Trim() + " 起至 " + eViolationDate_End.Text.Trim() + " 止" :
                                                   ((eViolationDate_Start.Text.Trim() != "") && (eViolationDate_End.Text.Trim() == "")) ? eViolationDate_Start.Text.Trim() :
                                                   ((eViolationDate_Start.Text.Trim() == "") && (eViolationDate_End.Text.Trim() != "")) ? eViolationDate_End.Text.Trim() : "";
                        string vDepNoStr = ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() != "")) ? "從 " + eDepNo_Start.Text.Trim() + " 起至 " + eDepNo_End.Text.Trim() + " 止" :
                                           ((eDepNo_Start.Text.Trim() != "") && (eDepNo_End.Text.Trim() == "")) ? eDepNo_Start.Text.Trim() :
                                           ((eDepNo_Start.Text.Trim() == "") && (eDepNo_End.Text.Trim() != "")) ? eDepNo_End.Text.Trim() : "全部";
                        string vDriverStr = (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : "";
                        string vCarIDStr = (eCar_ID.Text.Trim() != "") ? eCar_ID.Text.Trim() : "";
                        string vViolationRuleStr = (eViolationRule_Search.Text.Trim() != "") ? eViolationRule_Search.Text.Trim() : "";
                        string vRecordNote = "匯出檔案_違規案件統計表" + Environment.NewLine +
                                             "ViolationCaseP.aspx" + Environment.NewLine +
                                             "報表種類：" + vPrintModeStr + Environment.NewLine +
                                             "建檔日期：" + vBuildDateStr + Environment.NewLine +
                                             "違規日期：" + vViolationDateStr + Environment.NewLine +
                                             "站別：" + vDepNoStr + Environment.NewLine +
                                             "駕駛員：" + vDriverStr + Environment.NewLine +
                                             "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                             "違規條文：" + vViolationRuleStr;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                        //先判斷是不是IE
                        HttpContext.Current.Response.ContentType = "application/octet-stream";
                        HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                        string TourVerision = brObject.Type;
                        if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                        {
                            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + fFileName + ".xls", System.Text.Encoding.UTF8));
                        }
                        else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                        {
                            // 設定強制下載標頭
                            Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fFileName + ".xls"));
                        }
                        // 輸出檔案
                        Response.BinaryWrite(msTarget.ToArray());
                        msTarget.Close();

                        Response.End();
                    }
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void rbPrintMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            gvShowData_DepNo.Visible = false;
            gvShowData_Driver.Visible = false;
            gvShowData_Point.Visible = false;
            gvShowData_99.Visible = false;
            switch (rbPrintMode.SelectedValue)
            {
                case "11":
                case "13":
                    gvShowData_DepNo.Visible = (gvShowData_DepNo.Rows.Count > 0);
                    break;
                case "12":
                    gvShowData_Driver.Visible = (gvShowData_Driver.Rows.Count > 0);
                    break;
                case "21":
                    gvShowData_Point.Visible = (gvShowData_Point.Rows.Count > 0);
                    break;
                case "99":
                    gvShowData_99.Visible = (gvShowData_99.Rows.Count > 0);
                    break;
            }
        }

        protected void bbPreviewData_Click(object sender, EventArgs e)
        {
            ListDataBind();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vSelectStr = "";
            plShowData.Visible = true;
            plReport.Visible = false;
            switch (rbPrintMode.SelectedValue)
            {
                case "11":
                    vSelectStr = CaseDataBind_DepNo("V01");
                    PrintReport_DepNo(vSelectStr);
                    break;
                case "12":
                    vSelectStr = CaseDataBind_Driver();
                    PrintReport_Driver(vSelectStr);
                    break;
                case "13":
                    vSelectStr = CaseDataBind_DepNo("V02");
                    PrintReport_DepNo(vSelectStr);
                    break;
                case "21":
                    vSelectStr = CaseDataBind_Point();
                    PrintReport_Point(vSelectStr);
                    break;
                case "99":
                    vSelectStr = CaseDataBind_99();
                    PrintReport_99(vSelectStr);
                    break;
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelectStr = "";
            switch (rbPrintMode.SelectedValue)
            {
                case "11":
                    vSelectStr = CaseDataBind_DepNo("V01");
                    SaveExcel(rbPrintMode.SelectedItem.Text.Trim(), vSelectStr);
                    break;
                case "12":
                    vSelectStr = CaseDataBind_Driver();
                    SaveExcel(rbPrintMode.SelectedItem.Text.Trim(), vSelectStr);
                    break;
                case "13":
                    vSelectStr = CaseDataBind_DepNo("V02");
                    SaveExcel(rbPrintMode.SelectedItem.Text.Trim(), vSelectStr);
                    break;
                case "21":
                    vSelectStr = CaseDataBind_Point(); ;
                    SaveExcel(rbPrintMode.SelectedItem.Text.Trim(), vSelectStr);
                    break;
                case "99":
                    vSelectStr = CaseDataBind_99();
                    SaveExcel(rbPrintMode.SelectedItem.Text.Trim(), vSelectStr);
                    break;
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            rvPrint.LocalReport.DataSources.Clear();
            plReport.Visible = false;
        }
    }
}