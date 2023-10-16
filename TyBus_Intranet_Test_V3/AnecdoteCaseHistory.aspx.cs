using Amaterasu_Function;
using System;
using System.Globalization;

namespace TyBus_Intranet_Test_V3
{
    public partial class AnecdoteCaseHistory : System.Web.UI.Page
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
                        eBuildDate_Start_Search.Text = (eBuildDate_Start_Search.Text.Trim() != "") ? eBuildDate_Start_Search.Text.Trim() : PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C");
                        eBuildDate_End_Search.Text = (eBuildDate_End_Search.Text.Trim() != "") ? eBuildDate_End_Search.Text.Trim() : PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C");

                        plShowData.Visible = true;
                    }

                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    string vBuildDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_Start_Search.ClientID;
                    string vBuildDateScript = "window.open('" + vBuildDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuildDate_Start_Search.Attributes["onClick"] = vBuildDateScript;

                    vBuildDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_End_Search.ClientID;
                    vBuildDateScript = "window.open('" + vBuildDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuildDate_End_Search.Attributes["onClick"] = vBuildDateScript;
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
            if (vLoginID != "")
            {
                string vWStr_BuildDate = ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                         "   and a.BuildDate between '" + PF.TransDateString(eBuildDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eBuildDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() == "")) ?
                                         "   and a.BuildDate = '" + PF.TransDateString(eBuildDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eBuildDate_Start_Search.Text.Trim() == "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                         "   and a.BuildDate = '" + PF.TransDateString(eBuildDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";
                string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? "   and a.DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and a.DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
                vSQLStr = "SELECT HistoryNo, CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, " + Environment.NewLine +
                          "       (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuildMan)) AS BuildMan_C, Car_ID, Driver, DriverName, InsuMan, " + Environment.NewLine +
                          "       AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, ModifyType, " + Environment.NewLine +
                          "       (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = a.ModifyType) AND (FKEY = CAST('客訴單歷史' AS char(16)) + CAST('fmServiceHistory' AS char(16)) + 'ModifyType')) AS ModifyType_C, " + Environment.NewLine +
                          "       ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyMan_C, CaseOccurrence " + Environment.NewLine +
                          "  FROM AnecdoteCaseHistory AS a " + Environment.NewLine +
                          " where 1 = 1 " + Environment.NewLine +
                          vWStr_BuildDate +
                          vWStr_DepNo +
                          " order by HistoryNo DESC";
                sdsAnecdoteCaseHistoryList.SelectCommand = "";
                sdsAnecdoteCaseHistoryList.SelectCommand = vSQLStr;
                gridAnecdoteCaseHistory.DataBind();
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            CaseDataBind();
            string vRecordDateStr = ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() != "")) ? eBuildDate_Start_Search.Text.Trim() + "~" + eBuildDate_End_Search.Text.Trim() :
                                    ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() == "")) ? eBuildDate_Start_Search.Text.Trim() :
                                    ((eBuildDate_Start_Search.Text.Trim() == "") && (eBuildDate_End_Search.Text.Trim() != "")) ? eBuildDate_End_Search.Text.Trim() : "不指定日期";
            string vRecordDepStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_Start_Search.Text.Trim() + "~" + eDepNo_End_Search.Text.Trim() :
                                   ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? eDepNo_Start_Search.Text.Trim() :
                                   ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_End_Search.Text.Trim() : "全部";
            string vRecordNote = "查詢資料_肇事案件異動記錄" + Environment.NewLine +
                                 "AnecdoteCaseHistory.aspx" + Environment.NewLine +
                                 "站別：" + vRecordDepStr + Environment.NewLine +
                                 "發生日期：" + vRecordDateStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void eDepNo_Start_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Search = eDepNo_Start_Search.Text.Trim();
            string vDepName_Search = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Search + "' ";
            vDepName_Search = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Search == "")
            {
                vDepName_Search = vDepNo_Search;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Search + "' ";
                vDepNo_Search = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Start_Search.Text = vDepNo_Search;
            eDepName_Start_Search.Text = vDepName_Search;
        }

        protected void eDepNo_End_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Search = eDepNo_End_Search.Text.Trim();
            string vDepName_Search = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Search + "' ";
            vDepName_Search = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Search == "")
            {
                vDepName_Search = vDepNo_Search;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Search + "' ";
                vDepNo_Search = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_End_Search.Text = vDepNo_Search;
            eDepName_End_Search.Text = vDepName_Search;
        }
    }
}