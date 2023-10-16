using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverWorkHourAnalysisMain : System.Web.UI.Page
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
                    if (!IsPostBack)
                    {
                        plPrint.Visible = false;
                        plSearch.Visible = true;
                        eDriveYear_Search.Text = (DateTime.Today.AddMonths(-1).Year - 1911).ToString();
                        eDriveMonth_Search.SelectedIndex = DateTime.Today.AddMonths(-1).Month - 1;
                        if (Int32.Parse(vLoginDepNo) >= 11)
                        {
                            eDepNoS_Search.Text = PF.GetValue(vConnStr, "select DepNo from EmployeeDepNo where EmpNo = '" + vLoginID + "' and IsDefault = 1", "DepNo");
                            eDepNoS_Search.Enabled = false;
                            eDepNoE_Search.Text = "";
                            eDepNoE_Search.Enabled = false;
                        }
                        else
                        {
                            eDepNoS_Search.Text = "";
                            eDepNoS_Search.Enabled = true;
                            eDepNoE_Search.Text = "";
                            eDepNoE_Search.Enabled = true;
                        }
                    }
                    else
                    {
                        //OpenData();
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

        private string GetSelectStr_PrintMain()
        {
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ? DateTime.Today.AddMonths(-1) : DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vWStr_BuDate = "          and ra.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine;
            string vWStr_DepNo = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo between '" + eDepNoS_Search.Text.Trim() + "' and '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? "   and e.DepNo = '" + eDepNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo = '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eDriver_Search.Text.Trim() != "") ? "          and ra.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "select e.DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, cast(null as varchar(2)) RiskLevel, t.Driver, e.[Name], " + Environment.NewLine +
                                "       sum(t.WorkTime01) WorkTime01, sum(t.WorkTime02) WorkTime02, sum(t.WorkTime03) WorkTime03, sum(t.WorkTime04) WorkTime04, sum(t.WorkTime05) WorkTime05, " + Environment.NewLine +
                                "       sum(t.WorkTime06) WorkTime06, sum(t.WorkTime07) WorkTime07, sum(t.WorkTime08) WorkTime08, sum(t.WorkTime09) WorkTime09, sum(t.WorkTime10) WorkTime10, " + Environment.NewLine +
                                "       sum(t.WorkTime11) WorkTime11, sum(t.WorkTime12) WorkTime12, sum(t.WorkTime13) WorkTime13, sum(t.WorkTime14) WorkTime14, sum(t.WorkTime15) WorkTime15, " + Environment.NewLine +
                                "       sum(t.WorkTime16) WorkTime16, sum(t.WorkTime17) WorkTime17, sum(t.WorkTime18) WorkTime18, sum(t.WorkTime19) WorkTime19, sum(t.WorkTime20) WorkTime20, " + Environment.NewLine +
                                "       sum(t.WorkTime21) WorkTime21, sum(t.WorkTime22) WorkTime22, sum(t.WorkTime23) WorkTime23, sum(t.WorkTime24) WorkTime24, sum(t.WorkTime25) WorkTime25, " + Environment.NewLine +
                                "       sum(t.WorkTime26) WorkTime26, sum(t.WorkTime27) WorkTime27, sum(t.WorkTime28) WorkTime28, sum(t.WorkTime29) WorkTime29, sum(t.WorkTime30) WorkTime30, sum(t.WorkTime31) WorkTime31, " + Environment.NewLine +
                                "       sum(t.WorkTotalMin) TMin, round(cast(sum(t.WorkTotalMin) as float) / cast(sum(t.WorkDays) as float) / 60, 2) AvgHours, " + Environment.NewLine +
                                "       cast(NULL as varchar) WorkState01, cast(NULL as varchar) WorkState02, cast(NULL as varchar) WorkState03, cast(NULL as varchar) WorkState04, cast(NULL as varchar) WorkState05, " + Environment.NewLine +
                                "       cast(NULL as varchar) WorkState06, cast(NULL as varchar) WorkState07, cast(NULL as varchar) WorkState08, cast(NULL as varchar) WorkState09, cast(NULL as varchar) WorkState10, " + Environment.NewLine +
                                "       cast(NULL as varchar) WorkState11, cast(NULL as varchar) WorkState12, cast(NULL as varchar) WorkState13, cast(NULL as varchar) WorkState14, cast(NULL as varchar) WorkState15, " + Environment.NewLine +
                                "       cast(NULL as varchar) WorkState16, cast(NULL as varchar) WorkState17, cast(NULL as varchar) WorkState18, cast(NULL as varchar) WorkState19, cast(NULL as varchar) WorkState20, " + Environment.NewLine +
                                "       cast(NULL as varchar) WorkState21, cast(NULL as varchar) WorkState22, cast(NULL as varchar) WorkState23, cast(NULL as varchar) WorkState24, cast(NULL as varchar) WorkState25, " + Environment.NewLine +
                                "       cast(NULL as varchar) WorkState26, cast(NULL as varchar) WorkState27, cast(NULL as varchar) WorkState28, cast(NULL as varchar) WorkState29, cast(NULL as varchar) WorkState30, cast(NULL as varchar) WorkState31, " + Environment.NewLine +
                                "       cast(NULL as varchar) ESC02Hour01, cast(NULL as varchar) ESC02Hour02, cast(NULL as varchar) ESC02Hour03, cast(NULL as varchar) ESC02Hour04, cast(NULL as varchar) ESC02Hour05, " + Environment.NewLine +
                                "       cast(NULL as varchar) ESC02Hour06, cast(NULL as varchar) ESC02Hour07, cast(NULL as varchar) ESC02Hour08, cast(NULL as varchar) ESC02Hour09, cast(NULL as varchar) ESC02Hour10, " + Environment.NewLine +
                                "       cast(NULL as varchar) ESC02Hour11, cast(NULL as varchar) ESC02Hour12, cast(NULL as varchar) ESC02Hour13, cast(NULL as varchar) ESC02Hour14, cast(NULL as varchar) ESC02Hour15, " + Environment.NewLine +
                                "       cast(NULL as varchar) ESC02Hour16, cast(NULL as varchar) ESC02Hour17, cast(NULL as varchar) ESC02Hour18, cast(NULL as varchar) ESC02Hour19, cast(NULL as varchar) ESC02Hour20, " + Environment.NewLine +
                                "       cast(NULL as varchar) ESC02Hour21, cast(NULL as varchar) ESC02Hour22, cast(NULL as varchar) ESC02Hour23, cast(NULL as varchar) ESC02Hour24, cast(NULL as varchar) ESC02Hour25, " + Environment.NewLine +
                                "       cast(NULL as varchar) ESC02Hour26, cast(NULL as varchar) ESC02Hour27, cast(NULL as varchar) ESC02Hour28, cast(NULL as varchar) ESC02Hour29, cast(NULL as varchar) ESC02Hour30, cast(NULL as varchar) ESC02Hour31, " + Environment.NewLine +
                                "       cast(NULL as varchar) Line99908Time01, cast(NULL as varchar) Line99908Time02, cast(NULL as varchar) Line99908Time03, cast(NULL as varchar) Line99908Time04, cast(NULL as varchar) Line99908Time05, " + Environment.NewLine +
                                "       cast(NULL as varchar) Line99908Time06, cast(NULL as varchar) Line99908Time07, cast(NULL as varchar) Line99908Time08, cast(NULL as varchar) Line99908Time09, cast(NULL as varchar) Line99908Time10, " + Environment.NewLine +
                                "       cast(NULL as varchar) Line99908Time11, cast(NULL as varchar) Line99908Time12, cast(NULL as varchar) Line99908Time13, cast(NULL as varchar) Line99908Time14, cast(NULL as varchar) Line99908Time15, " + Environment.NewLine +
                                "       cast(NULL as varchar) Line99908Time16, cast(NULL as varchar) Line99908Time17, cast(NULL as varchar) Line99908Time18, cast(NULL as varchar) Line99908Time19, cast(NULL as varchar) Line99908Time20, " + Environment.NewLine +
                                "       cast(NULL as varchar) Line99908Time21, cast(NULL as varchar) Line99908Time22, cast(NULL as varchar) Line99908Time23, cast(NULL as varchar) Line99908Time24, cast(NULL as varchar) Line99908Time25, " + Environment.NewLine +
                                "       cast(NULL as varchar) Line99908Time26, cast(NULL as varchar) Line99908Time27, cast(NULL as varchar) Line99908Time28, cast(NULL as varchar) Line99908Time29, cast(NULL as varchar) Line99908Time30, cast(NULL as varchar) Line99908Time31 " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "        select ra.Driver, (ra.WorkHR * 60 + ra.WorkMin) WorkTotalMin, case when isnull(ra.Driver, '') <> '' then 1 else 0 end WorkDays, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 1 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime01, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 2 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime02, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 3 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime03, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 4 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime04, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 5 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime05, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 6 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime06, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 7 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime07, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 8 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime08, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 9 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime09, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 10 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime10, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 11 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime11, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 12 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime12, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 13 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime13, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 14 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime14, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 15 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime15, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 16 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime16, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 17 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime17, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 18 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime18, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 19 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime19, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 20 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime20, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 21 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime21, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 22 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime22, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 23 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime23, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 24 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime24, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 25 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime25, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 26 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime26, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 27 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime27, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 28 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime28, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 29 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime29, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 30 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime30, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 31 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime31 " + Environment.NewLine +
                                "          from RunSheetA ra " + Environment.NewLine +
                                "         where isnull(ra.AssignNo, '') <> '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_EmpNo +
                                "         group by ra.Driver, ra.BuDate, ra.WorkHR, ra.WorkMin " + Environment.NewLine +
                                ") t left join Employee e on e.EmpNo = t.Driver " + Environment.NewLine +
                                " where isnull(t.Driver, '') <> '' " + Environment.NewLine +
                                vWStr_DepNo +
                                " group by e.DepNo, t.Driver, e.[Name] " + Environment.NewLine +
                                " order by e.DepNo, t.Driver";
            return vResultStr;
        }

        private string GetSelectStr_Detail()
        {
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ? DateTime.Today.AddMonths(-1) : DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vWStr_BuDate = "          and ra.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine;
            string vWStr_DepNo = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo between '" + eDepNoS_Search.Text.Trim() + "' and '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? "   and e.DepNo = '" + eDepNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo = '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eDriver_Search.Text.Trim() != "") ? "          and ra.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "select e.DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, cast(null as varchar(2)) RiskLevel, t.Driver, e.[Name], " + Environment.NewLine +
                                "       sum(t.WorkTime01) WorkTime01, sum(t.WorkTime02) WorkTime02, sum(t.WorkTime03) WorkTime03, sum(t.WorkTime04) WorkTime04, sum(t.WorkTime05) WorkTime05, " + Environment.NewLine +
                                "       sum(t.WorkTime06) WorkTime06, sum(t.WorkTime07) WorkTime07, sum(t.WorkTime08) WorkTime08, sum(t.WorkTime09) WorkTime09, sum(t.WorkTime10) WorkTime10, " + Environment.NewLine +
                                "       sum(t.WorkTime11) WorkTime11, sum(t.WorkTime12) WorkTime12, sum(t.WorkTime13) WorkTime13, sum(t.WorkTime14) WorkTime14, sum(t.WorkTime15) WorkTime15, " + Environment.NewLine +
                                "       sum(t.WorkTime16) WorkTime16, sum(t.WorkTime17) WorkTime17, sum(t.WorkTime18) WorkTime18, sum(t.WorkTime19) WorkTime19, sum(t.WorkTime20) WorkTime20, " + Environment.NewLine +
                                "       sum(t.WorkTime21) WorkTime21, sum(t.WorkTime22) WorkTime22, sum(t.WorkTime23) WorkTime23, sum(t.WorkTime24) WorkTime24, sum(t.WorkTime25) WorkTime25, " + Environment.NewLine +
                                "       sum(t.WorkTime26) WorkTime26, sum(t.WorkTime27) WorkTime27, sum(t.WorkTime28) WorkTime28, sum(t.WorkTime29) WorkTime29, sum(t.WorkTime30) WorkTime30, sum(t.WorkTime31) WorkTime31, " + Environment.NewLine +
                                "       sum(t.WorkTotalMin) TMin, round(cast(sum(t.WorkTotalMin) as float) / cast(sum(t.WorkDays) as float) / 60, 2) AvgHours, sum(t.WorkDays) TotalWorkDays " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "        select ra.Driver, (ra.WorkHR * 60 + ra.WorkMin) WorkTotalMin, case when isnull(ra.Driver, '') <> '' then 1 else 0 end WorkDays, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 1 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime01, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 2 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime02, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 3 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime03, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 4 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime04, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 5 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime05, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 6 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime06, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 7 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime07, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 8 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime08, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 9 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime09, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 10 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime10, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 11 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime11, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 12 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime12, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 13 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime13, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 14 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime14, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 15 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime15, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 16 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime16, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 17 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime17, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 18 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime18, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 19 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime19, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 20 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime20, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 21 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime21, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 22 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime22, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 23 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime23, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 24 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime24, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 25 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime25, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 26 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime26, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 27 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime27, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 28 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime28, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 29 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime29, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 30 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime30, " + Environment.NewLine +
                                "               case when Day(ra.BuDate) = 31 then ra.WorkHR * 60 + ra.WorkMin else null end WorkTime31 " + Environment.NewLine +
                                "          from RunSheetA ra " + Environment.NewLine +
                                "         where isnull(ra.AssignNo, '') <> '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_EmpNo +
                                "         group by ra.Driver, ra.BuDate, ra.WorkHR, ra.WorkMin " + Environment.NewLine +
                                ") t left join Employee e on e.EmpNo = t.Driver " + Environment.NewLine +
                                " where isnull(t.Driver, '') <> '' " + Environment.NewLine +
                                vWStr_DepNo +
                                " group by e.DepNo, t.Driver, e.[Name] " + Environment.NewLine +
                                " order by e.DepNo, t.Driver";
            return vResultStr;
        }

        private string GetSelectStr_Month()
        {
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ? DateTime.Today.AddMonths(-1) : DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vWStr_BuDate = "   and ra.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine;
            string vWStr_DepNo = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo between '" + eDepNoS_Search.Text.Trim() + "' and '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? "   and e.DepNo = '" + eDepNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo = '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eDriver_Search.Text.Trim() != "") ? "   and ra.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "select e.DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                "       e.EmpNo, e.[Name], cast(NULL as varchar(2)) RiskLevel, " + Environment.NewLine +
                                "       sum((ra.WorkHR * 60 + ra.WorkMin)) WorkTotalMin, sum(case when isnull(ra.Driver, '') <> '' then 1 else 0 end) WorkDays, " + Environment.NewLine +
                                "       round(cast(sum((ra.WorkHR * 60 + ra.WorkMin)) as float) / cast(sum(case when isnull(ra.Driver, '') <> '' then 1 else 0 end) as float) / 60, 2) AVGWorkHR " + Environment.NewLine +
                                "  from RunSheetA ra left join Employee e on e.EmpNo = ra.Driver " + Environment.NewLine +
                                " where isnull(ra.Driver, '') <> '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo + vWStr_EmpNo +
                                " group by e.DepNo, e.EmpNo, e.[Name] " + Environment.NewLine +
                                " order by e.DepNo, e.EmpNo";
            return vResultStr;
        }

        private string GetSelectStr_Status()
        {
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ? DateTime.Today.AddMonths(-1) : DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vCalYM = (vCalDate.Year - 1911).ToString() + vCalDate.Month.ToString("D2");
            string vWStr_BuDate = "           and ra.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine;
            string vWStr_DepNo = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? "           and ra.DepNo between '" + eDepNoS_Search.Text.Trim() + "' and '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? "           and ra.DepNo = '" + eDepNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? "           and ra.DepNo = '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eDriver_Search.Text.Trim() != "") ? "   and ra.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_RiskLevel = (cbHasRiskLevel.Checked) ? " where (select case when isnull(DoctorLevel, '') <> '' then DoctorLevel else CompanyLevel end RiskLevel from DriverRiskLevel where IDCardNo = totalData.IDCardNo and CalYM = '" + vCalYM + "') >= '3'" + Environment.NewLine : "";
            string vResultStr = "select totalData.DepName, totalData.Driver, totalData.DriverName, " + Environment.NewLine +
                                "       totalData.LowOverduty, totalData.MiddleOverduty, totalData.HighOverduty, " + Environment.NewLine +
                                "       cast(totalData.TotalMin / 60 as nvarchar) + ':' + Right('00' + cast(totalData.TotalMin % 60 as nvarchar), 2) as MonthWorkHR, " + Environment.NewLine +
                                "       case when totalData.TotalMin < 13320 then 'V' else '' end as HRRange_1, " + Environment.NewLine +
                                "       case when totalData.TotalMin >= 13320 and totalData.TotalMin <= 18000 then 'V' else '' end as HRRange_2, " + Environment.NewLine +
                                "       case when totalData.TotalMin > 18000 and totalData.TotalMin <= 21000 then 'V' else '' end as HRRange_3, " + Environment.NewLine +
                                "       case when totalData.TotalMin > 21000 and totalData.TotalMin <= 24000 then 'V' else '' end as HRRange_4, " + Environment.NewLine +
                                "       case when totalData.TotalMin > 24000 then 'V' else '' end as HRRange_5, " + Environment.NewLine +
                                "       totalData.WorkDays, round(cast(totalData.TotalMin as float) / 60 / cast(totalData.WorkDays as float), 2) as AVGHR, " + Environment.NewLine +
                                "       isnull((select case when isnull(DoctorLevel, '') <> '' then DoctorLevel else CompanyLevel end RiskLevel from DriverRiskLevel where IDCardNo = totalData.IDCardNo and CalYM = '" + vCalYM + "'), '') RiskLevel " + Environment.NewLine +
                                "  from (" + Environment.NewLine +
                                "        select DepNo, (select[Name] from Department where DepNo = ra.DepNo) DepName, " + Environment.NewLine +
                                "               Driver, (select[Name] from Employee where EmpNo = ra.Driver) DriverName, (select IDCardNo from Employee where EmpNo = ra.Driver) as IDCardNo, " + Environment.NewLine +
                                "               sum(case when((cast(ra.WorkHR as int) * 60) + cast(ra.WorkMin as int)) < 720 then 1 else 0 end) as LowOverduty, " + Environment.NewLine +
                                "               sum(case when((cast(ra.WorkHR as int) * 60) + cast(ra.WorkMin as int)) >= 720 and((cast(ra.WorkHR as int) * 60) + cast(ra.WorkMin as int)) < 900 then 1 else 0 end) as MiddleOverduty, " + Environment.NewLine +
                                "               sum(case when((cast(ra.WorkHR as int) * 60) + cast(ra.WorkMin as int)) >= 900 then 1 else 0 end) as HighOverduty, " + Environment.NewLine +
                                "               sum(((cast(ra.WorkHR as int) * 60) + cast(ra.WorkMin as int))) as TotalMin, " + Environment.NewLine +
                                "               count(ra.AssignNo) as WorkDays " + Environment.NewLine +
                                "          from RunSheetA ra " + Environment.NewLine +
                                "         where isnull(ra.Driver, '') <> '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo + vWStr_EmpNo +
                                "         group by ra.DepNo, ra.Driver " + Environment.NewLine +
                                "		) totalData " + Environment.NewLine +
                                vWStr_RiskLevel +
                                " order by totalData.DepNo, totalData.DRIVER";
            return vResultStr;
        }

        private string GetSelectStr_Overduty()
        {
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ? DateTime.Today.AddMonths(-1) : DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vCalYM = (vCalDate.Year - 1911).ToString() + vCalDate.Month.ToString("D2");
            string vWStr_BuDate = "           and ra.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine;
            string vWStr_DepNo = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? "           and ra.DepNo between '" + eDepNoS_Search.Text.Trim() + "' and '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? "           and ra.DepNo = '" + eDepNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? "           and ra.DepNo = '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eDriver_Search.Text.Trim() != "") ? "   and ra.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_RiskLevel = (cbHasRiskLevel.Checked) ? " where (select case when isnull(DoctorLevel, '') <> '' then DoctorLevel else CompanyLevel end RiskLevel from DriverRiskLevel where IDCardNo = totalData.IDCardNo and CalYM = '" + vCalYM + "') >= '3'" + Environment.NewLine : "";
            string vResultStr = "select totalData.DepNo, (select [Name] from Department where DepNo = totalData.DepNo) [DepName], " + Environment.NewLine +
                                "       totalData.Driver, (select[Name] from Employee where EmpNo = totalData.Driver) [DriverName], " + Environment.NewLine +
                                "       totalData.LiveAMT_T, totalData.RentOverAMT_T, " + Environment.NewLine +
                                "       cast(cast(totalData.NMin1 as int) / 60 as varchar) + ':' + right('00' + cast(cast(totalData.NMin1 as int) % 60 as varchar), 2) as [OverDuty_1], " + Environment.NewLine +
                                "       cast(cast(totalData.NMin2 as int) / 60 as varchar) + ':' + right('00' + cast(cast(totalData.NMin2 as int) % 60 as varchar), 2) as [OverDuty_2], " + Environment.NewLine +
                                "       cast(cast(totalData.RMin1 as int) / 60 as varchar) + ':' + right('00' + cast(cast(totalData.RMin1 as int) % 60 as varchar), 2) as [OverDuty_3], " + Environment.NewLine +
                                "       cast(cast(totalData.RMin2 as int) / 60 as varchar) + ':' + right('00' + cast(cast(totalData.RMin2 as int) % 60 as varchar), 2) as [OverDuty_4], " + Environment.NewLine +
                                "       cast(cast(totalData.RMin3 as int) / 60 as varchar) + ':' + right('00' + cast(cast(totalData.RMin3 as int) % 60 as varchar), 2) as [OverDuty_5], " + Environment.NewLine +
                                "       cast(cast(totalData.VMin1 as int) / 60 as varchar) + ':' + right('00' + cast(cast(totalData.VMin1 as int) % 60 as varchar), 2) as [OverDuty_6], " + Environment.NewLine +
                                "       cast(cast(totalData.VMin2 as int) / 60 as varchar) + ':' + right('00' + cast(cast(totalData.VMin2 as int) % 60 as varchar), 2) as [OverDuty_7], " + Environment.NewLine +
                                "       cast(cast(totalData.VMin3 as int) / 60 as varchar) + ':' + right('00' + cast(cast(totalData.VMin3 as int) % 60 as varchar), 2) as [OverDuty_8], " + Environment.NewLine +
                                "       cast(cast(totalData.VMin4 as int) / 60 as varchar) + ':' + right('00' + cast(cast(totalData.VMin4 as int) % 60 as varchar), 2) as [OverDuty_9], " + Environment.NewLine +
                                //2021.03.30 更新
                                //"       totalData.WState1, totalData.WState2, totalData.WState3 " + Environment.NewLine +
                                "       totalData.WState1, totalData.WState2, totalData.WState3, pr.CashNum01, pr.CashNum02, " + Environment.NewLine +
                                //2021.03.31 更新
                                //"       Round(((pr.Nowpay5 + pr.Nowpay6 + pr.Nowpay7 + pr.CashNum11 + pr.CashNum07 + pr.CashNum13 + pr.CashNum26 + pr.CashNum06 + pr.CashNum08 + pr.CashNum14 + pr.CashNum30 + pr.CashNum12 - totalData.V14) / (30.0 * 8.0)), 4) as TimePay "+Environment.NewLine+
                                "       (pr.Nowpay5 + pr.Nowpay6 + pr.Nowpay7 + pr.CashNum11 + pr.CashNum07 + pr.CashNum13 + pr.CashNum26 + pr.CashNum06 + pr.CashNum08 + pr.CashNum14 + pr.CashNum30 + pr.CashNum12 - totalData.V14) as TimePay " + Environment.NewLine +
                                //===========================================================================================================================================
                                "  from ( " + Environment.NewLine +
                                "        select DepNo, Driver, (select IDCardNo from Employee where EmpNo = ra.Driver) as IDCardNo, sum(LiveAMT) as LiveAMT_T, sum(RentOverAMT) as RentOverAMT_T, " + Environment.NewLine +
                                "               sum(case when WorkState = '0' and(WorkHR * 60 + WorkMin) > 480 and(WorkHR * 60 + WorkMin) <= 600 then(WorkHR * 60 + WorkMin) - 480 " + Environment.NewLine +
                                "                        when WorkState = '0' and(WorkHR * 60 + WorkMin) > 600 then 120 else 0 end) as NMin1, " + Environment.NewLine +
                                "               sum(case when WorkState = '0' and(WorkHR * 60 + WorkMin) > 600 then(WorkHR * 60 + WorkMin) - 600 else 0 end) as NMin2, " + Environment.NewLine +
                                "               sum(case when WorkState = '1' and(WorkHR * 60 + WorkMin) <= 120 then(WorkHR * 60 + WorkMin) " + Environment.NewLine +
                                "                        when WorkState = '1' and(WorkHR * 60 + WorkMin) > 120 then 120 else 0 end) as RMin1, " + Environment.NewLine +
                                "               sum(case when WorkState = '1' and(WorkHR * 60 + WorkMin) > 120 and(WorkHR * 60 + WorkMin) <= 480 then(WorkHR * 60 + WorkMin) - 120 " + Environment.NewLine +
                                "                        when WorkState = '1' and(WorkHR * 60 + WorkMin) > 480 then 360 else 0 end) as RMin2, " + Environment.NewLine +
                                "               sum(case when WorkState = '1' and(WorkHR * 60 + WorkMin) > 480 then(WorkHR * 60 + WorkMin) - 480  else 0 end) as RMin3, " + Environment.NewLine +
                                "               sum(case when WorkState = '2' and(WorkHR * 60 + WorkMin) > 480 and(WorkHR * 60 + WorkMin) <= 600 then(WorkHR * 60 + WorkMin) - 480 " + Environment.NewLine +
                                "                        when WorkState = '2' and(WorkHR * 60 + WorkMin) > 600 then 120 else 0 end) as VMin1, " + Environment.NewLine +
                                "               sum(case when WorkState = '2' and(WorkHR * 60 + WorkMin) > 600 then(WorkHR * 60 + WorkMin) - 600 else 0 end) as VMin2, " + Environment.NewLine +
                                "               sum(case when WorkState = '3' and(WorkHR * 60 + WorkMin) > 480 and(WorkHR * 60 + WorkMin) <= 600 then(WorkHR * 60 + WorkMin) - 480 " + Environment.NewLine +
                                "                        when WorkState = '3' and(WorkHR * 60 + WorkMin) > 600 then 120 else 0 end) as VMin3, " + Environment.NewLine +
                                "               sum(case when WorkState = '3' and(WorkHR * 60 + WorkMin) > 600 then(WorkHR * 60 + WorkMin) - 600 else 0 end) as VMin4, " + Environment.NewLine +
                                "               sum(case when WorkState = '1' then 1 else 0 end) WState1, " + Environment.NewLine +
                                "               sum(case when WorkState = '2' then 1 else 0 end) WState2, " + Environment.NewLine +
                                //2021.03.30 更新
                                //"               sum(case when WorkState = '3' then 1 else 0 end) WState3 " + Environment.NewLine +
                                "               sum(case when WorkState = '3' then 1 else 0 end) WState3, " + Environment.NewLine +
                                "               (select isnull(sum(expense),0) " + Environment.NewLine +
                                "                  from Mshz " + Environment.NewLine +
                                "                 where PayNo = '14' " + Environment.NewLine +
                                "                   and PaybNo in ('1405', '1407', '1408', '1409', '1410', '1412', '1414', '1415', '1416', '1418') " + Environment.NewLine +
                                "                   and PayDate = '" + vBuDate_S + "' " + Environment.NewLine +
                                "                   and PayDur = '1' " + Environment.NewLine +
                                "                   and EmpNo = ra.Driver) as V14 " + Environment.NewLine +
                                //==========================================================================================================================================
                                "          from RunSheetA ra " + Environment.NewLine +
                                "         where isnull(ra.Driver, '') <> '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo + vWStr_EmpNo +
                                "         group by DepNo, Driver " + Environment.NewLine +
                                ") totalData " + Environment.NewLine +
                                //2021.03.30 新增
                                "  left join PayRec pr on pr.EmpNo = totalData.Driver " + Environment.NewLine +
                                "   and pr.PayDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine +
                                "   and pr.PayDur = '1' " + Environment.NewLine +
                                //==========================================================================================================================================
                                vWStr_RiskLevel +
                                " order by totalData.DepNo, totalData.Driver";
            return vResultStr;
        }

        private DataTable CalWorkHourDataMain()
        {
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ?
                                DateTime.Today.AddMonths(-1) :
                                DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vRiskLevelYear = "";
            string vRiskLevelYM = "";
            int vRiskLevel = 0;
            string vRiskLevelStr = "";
            string vTempStr = "";
            string vEmpNo = "";
            string vIDCardNo = "";
            DateTime vCalDay;
            string vESCHourFieldName = "";
            string vLines99908FieldName = "";
            string vWorkStateFieldName = "";
            DataTable tbTemp = new DataTable();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                string vSQLStr = GetSelectStr_PrintMain();
                SqlDataAdapter daPrint = new SqlDataAdapter(vSQLStr, connPrint);
                connPrint.Open();
                daPrint.Fill(tbTemp);
            }
            for (int i = 0; i < tbTemp.Rows.Count; i++)
            {
                //先取回職醫評等
                vRiskLevelYear = (vCalDate.Year - 1911).ToString();
                vRiskLevelYM = vRiskLevelYear + vCalDate.Month.ToString("D2");

                vEmpNo = tbTemp.Rows[i]["Driver"].ToString().Trim();
                vTempStr = "select IDCardNo from Employee where EmpNo = '" + vEmpNo + "' ";
                vIDCardNo = PF.GetValue(vConnStr, vTempStr, "IDCardNo");
                vTempStr = "select top 1 DoctorLevel " + Environment.NewLine +
                           "  from DriverRiskLevel " + Environment.NewLine +
                           " where CalYM <= '" + vRiskLevelYM + "' " + Environment.NewLine +
                           "   and IDCardNo = '" + vIDCardNo + "' " + Environment.NewLine +
                           "   and isnull(DoctorLevel, '') <> '' " + Environment.NewLine +
                           " order by CalYM DESC ";
                vRiskLevelStr = PF.GetValue(vConnStr, vTempStr, "DoctorLevel");
                if (Int32.TryParse(vRiskLevelStr, out vRiskLevel))
                {
                    tbTemp.Rows[i]["RiskLevel"] = (vRiskLevel > 2) ? vRiskLevel.ToString() : "";
                }
                else
                {
                    tbTemp.Rows[i]["RiskLevel"] = "";
                }
                //再取回公假時數
                using (SqlConnection connESC = new SqlConnection(vConnStr))
                {
                    vTempStr = "select ApplyMan, RealDay, sum([Hours]) TotalHours " + Environment.NewLine +
                               "  from ESCDuty " + Environment.NewLine +
                               " where ESCType = '02' " + Environment.NewLine +
                               "   and ApplyMan = '" + vEmpNo + "' " + Environment.NewLine +
                               "   and RealDay between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine +
                               " group by ApplyMan, RealDay " + Environment.NewLine +
                               " order by RealDay";
                    SqlCommand cmdESC = new SqlCommand(vTempStr, connESC);
                    connESC.Open();
                    SqlDataReader drESC = cmdESC.ExecuteReader();
                    while (drESC.Read())
                    {
                        vCalDay = DateTime.Parse(drESC["RealDay"].ToString().Trim());
                        vESCHourFieldName = "ESC02Hour" + vCalDay.Day.ToString("D2");
                        tbTemp.Rows[i][vESCHourFieldName] = (drESC["TotalHours"].ToString().Trim() != "") ? "公：" + drESC["TotalHours"].ToString().Trim() : "";
                    }
                }
                //再取回公出時數
                using (SqlConnection connLines99908 = new SqlConnection(vConnStr))
                {
                    vTempStr = "select ra.BuDate, case when (LTrim(RTrim(isnull(rb.ToTime, ':'))) <> ':') and (LTrim(RTrim(isnull(rb.BackTime, ':'))) <> ':') " + Environment.NewLine +
                               "                       then DateDiff(hour, convert(varchar(10), ra.BuDate, 111) +' ' + rb.ToTime, convert(varchar(10), ra.BuDate, 111) + ' ' + rb.BackTime) " + Environment.NewLine +
                               "                       else 0 end TotalTime " + Environment.NewLine +
                               "  from RunSheetA ra left join RunSheetB rb on rb.AssignNo = ra.AssignNo " + Environment.NewLine +
                               " where rb.LinesNo = '99908' " + Environment.NewLine +
                               "   and isnull(rb.ReduceReason, '') = '' " + Environment.NewLine +
                               "   and ra.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine +
                               "   and ra.Driver = '" + vEmpNo + "' " + Environment.NewLine +
                               " order by ra.BuDate";
                    SqlCommand cmdLines99908 = new SqlCommand(vTempStr, connLines99908);
                    connLines99908.Open();
                    SqlDataReader drLines99908 = cmdLines99908.ExecuteReader();
                    while (drLines99908.Read())
                    {
                        vCalDay = DateTime.Parse(drLines99908["BuDate"].ToString().Trim());
                        vLines99908FieldName = "Line99908Time" + vCalDay.Day.ToString("D2");
                        tbTemp.Rows[i][vLines99908FieldName] = (drLines99908["TotalTime"].ToString().Trim() != "") ? "出：" + drLines99908["TotalTime"].ToString().Trim() : "";
                    }
                }
                //最後取回 WorkState
                using (SqlConnection connWorkState = new SqlConnection(vConnStr))
                {
                    vTempStr = "select BuDate, WorkState " + Environment.NewLine +
                               "  from RunSheetA " + Environment.NewLine +
                               " where BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine +
                               "   and Driver = '" + vEmpNo + "' " + Environment.NewLine +
                               "   and isnull(WorkState, '0') <> '0' " + Environment.NewLine +
                               " order by BuDate ";
                    SqlCommand cmdWorkState = new SqlCommand(vTempStr, connWorkState);
                    connWorkState.Open();
                    SqlDataReader drWorkState = cmdWorkState.ExecuteReader();
                    while (drWorkState.Read())
                    {
                        vCalDay = DateTime.Parse(drWorkState["BuDate"].ToString().Trim());
                        vWorkStateFieldName = "WorkState" + vCalDay.Day.ToString("D2");
                        tbTemp.Rows[i][vWorkStateFieldName] = (drWorkState["WorkState"].ToString().Trim() == "1") ? "休" :
                                                              (drWorkState["WorkState"].ToString().Trim() == "2") ? "例" :
                                                              (drWorkState["WorkState"].ToString().Trim() == "3") ? "國" : "";
                    }
                }
            }
            //如果有勾選 "只列出評級 3 以上" 才執行
            if (cbHasRiskLevel.Checked)
            {
                for (int i = 0; i < tbTemp.Rows.Count; i++)
                {
                    vRiskLevelStr = tbTemp.Rows[i]["RiskLevel"].ToString().Trim();
                    if (vRiskLevelStr == "")
                    {
                        tbTemp.Rows[i].Delete();
                    }
                }
                tbTemp.AcceptChanges();
            }
            return tbTemp;
        }

        private DataTable CalWorkHourData_Detail()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ?
                                DateTime.Today.AddMonths(-1) :
                                DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            int vMonthDays = PF.GetMonthDays(vCalDate);
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vRiskLevelYear = "";
            string vRiskLevelYM = "";
            string vRiskLevelStr = "";
            int vRiskLevel = 0;
            string vTempStr = "";
            string vEmpNo = "";
            string vIDCardNo = "";
            string vMonthWorkHR = "";
            int vWorkTotalMin = 0;
            int vWorkTotalHR = 0;
            int vWorkTimeMin = 0;
            int vWorkTimeHR = 0;
            string vColName = "";
            DataTable vResultTable = new DataTable();
            DataRow vNewRow;

            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                vTempStr = GetSelectStr_Detail();
                SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    vResultTable.Columns.Clear();
                    vResultTable.Columns.Add("DepNo", typeof(String));
                    vResultTable.Columns.Add("DepName", typeof(String));
                    vResultTable.Columns.Add("Driver", typeof(String));
                    vResultTable.Columns.Add("Name", typeof(String));
                    for (int i = 1; i <= 31; i++)
                    {
                        vColName = "WorkTime" + i.ToString("D2");
                        vResultTable.Columns.Add(vColName, typeof(String));
                    }
                    vResultTable.Columns.Add("TotalWorkHR", typeof(String));
                    vResultTable.Columns.Add("AVGWorkHR", typeof(Double));
                    vResultTable.Columns.Add("WorkDays", typeof(int));
                    vResultTable.Columns.Add("DoctorRiskLevel", typeof(String));
                    while (drTemp.Read())
                    {
                        vEmpNo = drTemp["Driver"].ToString().Trim();
                        vTempStr = "select IDCardNo from Employee where EmpNo = '" + vEmpNo + "' ";
                        vIDCardNo = PF.GetValue(vConnStr, vTempStr, "IDCardNo");
                        vRiskLevelYear = (vCalDate.Year - 1911).ToString();
                        vRiskLevelYM = vRiskLevelYear + vCalDate.Month.ToString("D2");
                        vTempStr = "select top 1 DoctorLevel " + Environment.NewLine +
                                   "  from DriverRiskLevel " + Environment.NewLine +
                                   " where CalYM <= '" + vRiskLevelYM + "' " + Environment.NewLine +
                                   "   and IDCardNo = '" + vIDCardNo + "' " + Environment.NewLine +
                                   "   and isnull(DoctorLevel, '') <> '' " + Environment.NewLine +
                                   " order by CalYM DESC ";
                        if (Int32.TryParse(PF.GetValue(vConnStr, vTempStr, "DoctorLevel"), out vRiskLevel))
                        {
                            vRiskLevelStr = vRiskLevel.ToString();
                        }
                        else
                        {
                            vRiskLevelStr = "0";
                        }
                        vWorkTotalMin = (drTemp["TMin"].ToString().Trim() != "") ? Int32.Parse(drTemp["TMin"].ToString().Trim()) : 0;
                        vWorkTotalHR = (vWorkTotalMin != 0) ? vWorkTotalMin / 60 : 0;
                        vMonthWorkHR = vWorkTotalHR.ToString() + ":" + (vWorkTotalMin - (vWorkTotalHR * 60)).ToString("D2");

                        if ((cbHasRiskLevel.Checked == false) || ((cbHasRiskLevel.Checked) && (Int32.Parse(vRiskLevelStr) > 2)))
                        {
                            vNewRow = vResultTable.NewRow();
                            vNewRow["DepNo"] = drTemp["DepNo"];
                            vNewRow["DepName"] = drTemp["DepName"];
                            vNewRow["Driver"] = vEmpNo;
                            vNewRow["Name"] = drTemp["Name"];
                            for (int i = 1; i <= 31; i++)
                            {
                                vColName = "WorkTime" + i.ToString("D2");
                                vWorkTimeMin = (drTemp[vColName].ToString().Trim() != "") ? Int32.Parse(drTemp[vColName].ToString().Trim()) : 0;
                                vWorkTimeHR = (vWorkTimeMin != 0) ? vWorkTimeMin / 60 : 0;
                                vNewRow[vColName] = (vWorkTimeMin != 0) ? vWorkTimeHR.ToString() + ":" + (vWorkTimeMin - (vWorkTimeHR * 60)).ToString("D2") :
                                                    (i > vMonthDays) ? "" : "X";
                            }
                            vNewRow["TotalWorkHR"] = vMonthWorkHR.Trim();
                            vNewRow["AVGWorkHR"] = drTemp["AvgHours"];
                            vNewRow["WorkDays"] = drTemp["TotalWorkDays"];
                            vNewRow["DoctorRiskLevel"] = (Int32.Parse(vRiskLevelStr.Trim()) > 0) ? vRiskLevelStr.Trim() : "";
                            vResultTable.Rows.Add(vNewRow);
                        }
                    }
                }
            }
            return vResultTable;
        }

        private DataTable CalWorkHourData_Month()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ?
                                DateTime.Today.AddMonths(-1) :
                                DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vRiskLevelYear = "";
            string vRiskLevelYM = "";
            string vRiskLevelStr = "";
            int vRiskLevel = 0;
            string vTempStr = "";
            string vEmpNo = "";
            string vIDCardNo = "";
            string vMonthWorkHR = "";
            int vWorkTotalMin = 0;
            int vWorkTotalHR = 0;
            DataTable vResultTable = new DataTable();
            DataRow vNewRow;

            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                vTempStr = GetSelectStr_Month();
                SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    vResultTable.Columns.Clear();
                    vResultTable.Columns.Add("DepNo", typeof(string));
                    vResultTable.Columns.Add("DepName", typeof(string));
                    vResultTable.Columns.Add("EmpNo", typeof(string));
                    vResultTable.Columns.Add("EmpName", typeof(string));
                    vResultTable.Columns.Add("MonthWorkHR", typeof(string));
                    vResultTable.Columns.Add("WorkDays", typeof(int));
                    vResultTable.Columns.Add("AVGWorkHour", typeof(Double));
                    vResultTable.Columns.Add("DoctorRiskLevel", typeof(string));
                    while (drTemp.Read())
                    {
                        vEmpNo = drTemp["EmpNo"].ToString().Trim();
                        vTempStr = "select IDCardNo from Employee where EmpNo = '" + vEmpNo + "' ";
                        vIDCardNo = PF.GetValue(vConnStr, vTempStr, "IDCardNo");
                        vRiskLevelYear = (vCalDate.Year - 1911).ToString();
                        vRiskLevelYM = vRiskLevelYear + vCalDate.Month.ToString("D2");
                        vTempStr = "select top 1 DoctorLevel " + Environment.NewLine +
                                   "  from DriverRiskLevel " + Environment.NewLine +
                                   " where CalYM <= '" + vRiskLevelYM + "' " + Environment.NewLine +
                                   "   and IDCardNo = '" + vIDCardNo + "' " + Environment.NewLine +
                                   "   and isnull(DoctorLevel, '') <> '' " + Environment.NewLine +
                                   " order by CalYM DESC ";
                        if (Int32.TryParse(PF.GetValue(vConnStr, vTempStr, "DoctorLevel"), out vRiskLevel))
                        {
                            vRiskLevelStr = vRiskLevel.ToString();
                        }
                        else
                        {
                            vRiskLevelStr = "0";
                        }
                        vWorkTotalMin = (drTemp["WorkTotalMin"].ToString().Trim() != "") ? Int32.Parse(drTemp["WorkTotalMin"].ToString().Trim()) : 0;
                        vWorkTotalHR = (vWorkTotalMin != 0) ? vWorkTotalMin / 60 : 0;
                        vMonthWorkHR = vWorkTotalHR.ToString() + ":" + (vWorkTotalMin - (vWorkTotalHR * 60)).ToString("D2");

                        if ((cbHasRiskLevel.Checked == false) || ((cbHasRiskLevel.Checked) && (Int32.Parse(vRiskLevelStr) > 2)))
                        {
                            vNewRow = vResultTable.NewRow();
                            vNewRow["DepNo"] = drTemp["DepNo"];
                            vNewRow["DepName"] = drTemp["DepName"];
                            vNewRow["EmpNo"] = vEmpNo;
                            vNewRow["EmpName"] = drTemp["Name"];
                            vNewRow["MonthWorkHR"] = vMonthWorkHR;
                            vNewRow["WorkDays"] = drTemp["WorkDays"];
                            vNewRow["AVGWorkHour"] = drTemp["AVGWorkHR"];
                            vNewRow["DoctorRiskLevel"] = (Int32.Parse(vRiskLevelStr) > 2) ? vRiskLevelStr : String.Empty;
                            vResultTable.Rows.Add(vNewRow);
                        }
                    }
                }
            }
            return vResultTable;
        }

        private DataTable CalWorkHourData_Status()
        {
            DataTable vResultTable = new DataTable();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                string vSelStr = GetSelectStr_Status();
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelStr, connTemp);
                connTemp.Open();
                daTemp.Fill(vResultTable);
            }
            return vResultTable;
        }

        private DataTable CalWorkHourData_Overduty()
        {
            DataTable vResultTable = new DataTable();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                //2021.03.31 新增
                int vWorkDays = 0;
                string vEmpNo = "";
                double vTimePay = 0;
                DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ? DateTime.Today.AddMonths(-1) : DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
                DateTime vBuDate_S = DateTime.Parse(PF.GetMonthFirstDay(vCalDate, "C"));
                DateTime vBuDate_E = DateTime.Parse(PF.GetMonthLastDay(vCalDate, "C"));
                //==================================================================================================
                string vSelStr = GetSelectStr_Overduty();
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelStr, connTemp);
                connTemp.Open();
                daTemp.Fill(vResultTable);
                //2021.03.31 新增
                if (vResultTable.Rows.Count > 0)
                {
                    for (int i = 0; i < vResultTable.Rows.Count; i++)
                    {
                        vEmpNo = vResultTable.Rows[i]["Driver"].ToString().Trim();
                        vWorkDays = PF.CalEmployeeWorkDays(vConnStr, vEmpNo, vBuDate_S, vBuDate_E);
                        vWorkDays = (vWorkDays >= 30) ? 30 * 8 :
                                    ((vCalDate.Month == 2) && (vWorkDays >= 28)) ? 30 * 8 : vWorkDays * 8;
                        vTimePay = double.Parse(vResultTable.Rows[i]["TimePay"].ToString().Trim()) / double.Parse(vWorkDays.ToString().Trim());
                        vResultTable.Rows[i]["TimePay"] = Math.Round(vTimePay, 4, MidpointRounding.AwayFromZero);
                        vResultTable.AcceptChanges();
                    }
                }
                //==================================================================================================
            }
            return vResultTable;
        }

        /// <summary>
        /// 列印報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_Click(object sender, EventArgs e)
        {
            switch (ddlReportType_Search.SelectedValue)
            {
                case "0":
                    PrintReport_Main();
                    break;
                case "1":
                    PrintReport_Detail();
                    break;
                case "2":
                    PrintReport_Detail2();
                    break;
                case "3":
                    PrintReport_Month();
                    break;
                case "4":
                    PrintReport_Status();
                    break;
                case "5":
                    PrintReport_Overduty();
                    break;
            }
        }

        /// <summary>
        /// 列印總表
        /// </summary>
        private void PrintReport_Main()
        {
            string vFileName = "駕駛員工時統計分析總表";
            DataTable tbPrint = CalWorkHourDataMain();
            if (tbPrint.Rows.Count > 0)
            {
                ReportDataSource rdsPrint = new ReportDataSource("DriverWorkHourAnalysisMainP", tbPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\DriverWorkHourAnalysisMainP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vFileName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("CalDate", "行車日期：" + eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plSearch.Visible = false;

                string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                string vRecordNote = "列印資料_" + vFileName + Environment.NewLine +
                                     "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                     "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                     "站別：" + vRecordDepNoStr + Environment.NewLine +
                                     "駕駛員：" + vRecordDriverStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到符合條件的資料！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 列印駕駛員風險等級追蹤統計明細表
        /// </summary>
        private void PrintReport_Detail()
        {
            string vFileName = eDriveYear_Search.Text.Trim() + "年" + Int32.Parse(eDriveMonth_Search.SelectedValue).ToString("D2") + "月職醫評估駕駛員風險等級追蹤統計明細表";
            DataTable tbPrint = CalWorkHourData_Detail();
            if (tbPrint.Rows.Count > 0)
            {
                ReportDataSource rdsPrint = new ReportDataSource("DriverWorkHourAnalysisDetailP", tbPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\DriverWorkHourAnalysisDetailP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vFileName));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plSearch.Visible = false;

                string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                string vRecordNote = "列印資料_" + vFileName + Environment.NewLine +
                                     "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                     "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                     "站別：" + vRecordDepNoStr + Environment.NewLine +
                                     "駕駛員：" + vRecordDriverStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到符合條件的資料！')");
                Response.Write("</" + "Script>");
            }
        }

        private void PrintReport_Detail2()
        {
            string vFileName = eDriveYear_Search.Text.Trim() + "年" + Int32.Parse(eDriveMonth_Search.SelectedValue).ToString("D2") + "月職醫評估駕駛員風險等級追蹤統計明細表(依單位)";
            DataTable tbPrint = CalWorkHourData_Detail();
            if (tbPrint.Rows.Count > 0)
            {
                ReportDataSource rdsPrint = new ReportDataSource("DriverWorkHourAnalysisDetailP", tbPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\DriverWorkHourAnalysisDetailP2.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vFileName));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plSearch.Visible = false;

                string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                string vRecordNote = "列印資料_" + vFileName + Environment.NewLine +
                                     "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                     "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                     "站別：" + vRecordDepNoStr + Environment.NewLine +
                                     "駕駛員：" + vRecordDriverStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到符合條件的資料！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 列印各單位駕駛員月工時統計表
        /// </summary>
        private void PrintReport_Month()
        {
            string vFileName = eDriveYear_Search.Text.Trim() + "年" + Int32.Parse(eDriveMonth_Search.SelectedValue).ToString("D2") + "月職醫評估駕駛員風險等級追蹤統計總表";
            DataTable tbPrint = CalWorkHourData_Month();
            if (tbPrint.Rows.Count > 0)
            {
                ReportDataSource rdsPrint = new ReportDataSource("DriverWorkHourAnalysisMonthP", tbPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\DriverWorkHourAnalysisMonthP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vFileName));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plSearch.Visible = false;

                string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                string vRecordNote = "列印資料_" + vFileName + Environment.NewLine +
                                     "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                     "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                     "站別：" + vRecordDepNoStr + Environment.NewLine +
                                     "駕駛員：" + vRecordDriverStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到符合條件的資料！')");
                Response.Write("</" + "Script>");
            }
        }

        private void PrintReport_Status()
        {
            /* 2021.03.24 先卡位，程式碼等真的要出報表再重寫
            string vFileName = eDriveYear_Search.Text.Trim() + "年" + Int32.Parse(eDriveMonth_Search.SelectedValue).ToString("D2") + "月各單位駕駛員工時情況";
            DataTable tbPrint = CalWorkHourData_Month();
            if (tbPrint.Rows.Count > 0)
            {
                ReportDataSource rdsPrint = new ReportDataSource("DriverWorkHourAnalysisStatusP", tbPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\DriverWorkHourAnalysisStatusP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vFileName));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plSearch.Visible = false;

                string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                string vRecordNote = "列印資料_" + vFileName + Environment.NewLine +
                                     "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                     "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                     "站別：" + vRecordDepNoStr + Environment.NewLine +
                                     "駕駛員：" + vRecordDriverStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到符合條件的資料！')");
                Response.Write("</" + "Script>");
            } //*/
        }

        private void PrintReport_Overduty()
        {
            /* 2021.03.24 先卡位，程式碼等真的要出報表再重寫
            string vFileName = eDriveYear_Search.Text.Trim() + "年" + Int32.Parse(eDriveMonth_Search.SelectedValue).ToString("D2") + "月各單位駕駛員工時情況";
            DataTable tbPrint = CalWorkHourData_Month();
            if (tbPrint.Rows.Count > 0)
            {
                ReportDataSource rdsPrint = new ReportDataSource("DriverWorkHourAnalysisStatusP", tbPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\DriverWorkHourAnalysisStatusP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vFileName));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plSearch.Visible = false;

                string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                         ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                string vRecordNote = "列印資料_" + vFileName + Environment.NewLine +
                                     "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                     "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                     "站別：" + vRecordDepNoStr + Environment.NewLine +
                                     "駕駛員：" + vRecordDriverStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到符合條件的資料！')");
                Response.Write("</" + "Script>");
            } //*/
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
        {
            switch (ddlReportType_Search.SelectedValue)
            {
                case "0":
                    ExportExcel_Main();
                    break;
                case "1":
                    ExportExcel_Detail();
                    break;
                case "2":
                    ExportExcel_Detail2();
                    break;
                case "3":
                    ExportExcel_Month();
                    break;
                case "4":
                    ExportExcel_Status();
                    break;
                case "5":
                    ExportExcel_Overduty();
                    break;
            }
        }

        /// <summary>
        /// 總表匯出 EXCEL
        /// </summary>
        private void ExportExcel_Main()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if ((eDriveYear_Search.Text.Trim() != "") && (eDriveMonth_Search.Text.Trim() != ""))
            {
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

                HSSFCellStyle csData_Right = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Right.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Right.Alignment = HorizontalAlignment.Right;

                HSSFFont fontData_Right = (HSSFFont)wbExcel.CreateFont();
                fontData_Right.FontHeightInPoints = 12;
                csData_Right.SetFont(fontData);

                HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
                fontData_Red.FontHeightInPoints = 12; //字體大小
                fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
                csData_Red.SetFont(fontData_Red);

                HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Int.DataFormat = format.GetFormat("###,##0");

                int vLinesNo = 0;
                string vFileName = eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月駕駛員工時統計分析總表";
                string vRiskDepNo = "";
                string vGroupText = "";
                string vCellData = "";
                int vColIndex = 0;

                DataTable dtExcel = CalWorkHourDataMain();
                if (dtExcel.Rows.Count > 0)
                {
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        if (dtExcel.Rows[i]["DepNo"].ToString().Trim() != vRiskDepNo)
                        {
                            vRiskDepNo = dtExcel.Rows[i]["DepNo"].ToString().Trim();
                            vLinesNo = 0; //列數歸零
                                          //依單位分別寫進不同的工作表
                            vGroupText = dtExcel.Rows[i]["DepNo"].ToString().Trim() + "_" + dtExcel.Rows[i]["DepName"].ToString().Trim();
                            //建立新的工作表
                            wsExcel = (HSSFSheet)wbExcel.CreateSheet(vGroupText);
                            //建立標題列
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 2));
                            wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;

                            wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("日期");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 3, 32));
                            wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csTitle;

                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("職醫" + Environment.NewLine + "評等" + Environment.NewLine + "風險");
                            wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;

                            wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("編號");
                            wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csTitle;

                            wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("姓名");
                            wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csTitle;

                            for (int j = 1; j < 32; j++)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(j + 2).SetCellValue(j.ToString() + Environment.NewLine + "出勤" + Environment.NewLine + "狀況");
                                wsExcel.GetRow(vLinesNo).GetCell(j + 2).CellStyle = csTitle;
                            }
                            wsExcel.GetRow(vLinesNo).CreateCell(34).SetCellValue("總計" + Environment.NewLine + "出勤" + Environment.NewLine + "狀況");
                            wsExcel.GetRow(vLinesNo).GetCell(34).CellStyle = csTitle;

                            wsExcel.GetRow(vLinesNo).CreateCell(35).SetCellValue("平均" + Environment.NewLine + "每日" + Environment.NewLine + "時數");
                            wsExcel.GetRow(vLinesNo).GetCell(35).CellStyle = csTitle;
                            vLinesNo++;
                        }
                        wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                        wsExcel.CreateRow(vLinesNo);
                        for (int ColumnsIndex = 2; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                        {
                            if ((dtExcel.Columns[ColumnsIndex].ColumnName.Trim() == "RiskLevel") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim() == "Driver") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim() == "Name"))
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellValue(vCellData);
                                //wsExcel.AutoSizeColumn(ColumnsIndex - 2);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csTitle;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim() == "AvgHours")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellValue(vCellData);
                                //wsExcel.AutoSizeColumn(ColumnsIndex - 2);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csData_Int;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim() == "TMin")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                vCellData = (Int32.Parse(vCellData) / 60).ToString() + ":" + (Int32.Parse(vCellData) - 60 * (Int32.Parse(vCellData) / 60)).ToString("D2");
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellValue(vCellData);
                                //wsExcel.AutoSizeColumn(ColumnsIndex - 2);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csData;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Substring(0, 8) == "WorkTime")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "X";
                                vCellData = (vCellData != "X") ? (Int32.Parse(vCellData) / 60).ToString() + ":" + (Int32.Parse(vCellData) - 60 * (Int32.Parse(vCellData) / 60)).ToString("D2") : "X";
                                vColIndex = Int32.Parse(dtExcel.Columns[ColumnsIndex].ColumnName.Replace("WorkTime", "")) + 2;
                                wsExcel.GetRow(vLinesNo).CreateCell(vColIndex).SetCellValue(vCellData);
                                //wsExcel.AutoSizeColumn(vColIndex);
                                wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csData;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Substring(0, 9) == "WorkState")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                vColIndex = Int32.Parse(dtExcel.Columns[ColumnsIndex].ColumnName.Replace("WorkState", "")) + 2;
                                if (vColIndex == 3)
                                {
                                    vLinesNo++;
                                    wsExcel.CreateRow(vLinesNo);
                                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("");
                                    //wsExcel.AutoSizeColumn(0);
                                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 2));
                                }
                                wsExcel.GetRow(vLinesNo).CreateCell(vColIndex).SetCellValue(vCellData);
                                //wsExcel.AutoSizeColumn(vColIndex);
                                wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csData;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Substring(0, 9) == "ESC02Hour")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                vColIndex = Int32.Parse(dtExcel.Columns[ColumnsIndex].ColumnName.Replace("ESC02Hour", "")) + 2;
                                if (vColIndex == 3)
                                {
                                    vLinesNo++;
                                    wsExcel.CreateRow(vLinesNo);
                                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("公假時數");
                                    //wsExcel.AutoSizeColumn(0);
                                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 2));
                                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData_Right;
                                }
                                wsExcel.GetRow(vLinesNo).CreateCell(vColIndex).SetCellValue(vCellData);
                                //wsExcel.AutoSizeColumn(vColIndex);
                                wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csData_Int;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Substring(0, 13) == "Line99908Time")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                vColIndex = Int32.Parse(dtExcel.Columns[ColumnsIndex].ColumnName.Replace("Line99908Time", "")) + 2;
                                if (vColIndex == 3)
                                {
                                    vLinesNo++;
                                    wsExcel.CreateRow(vLinesNo);
                                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("公出時數");
                                    //wsExcel.AutoSizeColumn(0);
                                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 2));
                                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData_Right;
                                }
                                wsExcel.GetRow(vLinesNo).CreateCell(vColIndex).SetCellValue(vCellData);
                                //wsExcel.AutoSizeColumn(vColIndex);
                                wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csData_Int;
                            }
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
                            string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                            string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                                 "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
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
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('找不到符合條件的資料！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('行車日期年月不可空白！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 匯出各單位駕駛員工時明細
        /// </summary>
        private void ExportExcel_Detail()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if ((eDriveYear_Search.Text.Trim() != "") && (eDriveMonth_Search.Text.Trim() != ""))
            {
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

                HSSFCellStyle csData_Right = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Right.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Right.Alignment = HorizontalAlignment.Right;

                HSSFFont fontData_Right = (HSSFFont)wbExcel.CreateFont();
                fontData_Right.FontHeightInPoints = 12;
                csData_Right.SetFont(fontData);

                HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
                fontData_Red.FontHeightInPoints = 12; //字體大小
                fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
                csData_Red.SetFont(fontData_Red);

                HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Int.DataFormat = format.GetFormat("###,##0");

                HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Float.DataFormat = format.GetFormat("##0.00");

                int vLinesNo = 0;
                string vFileName = eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月各單位駕駛員月工時統計表";
                string vHeaderText = "";
                string vCellData = "";

                DataTable dtExcel = CalWorkHourData_Detail();
                if (dtExcel.Rows.Count > 0)
                {
                    //建立新的工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //建立標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int ColumnsIndex = 2; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                    {
                        vHeaderText = (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DRIVER") ? "工號" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "NAME") ? "姓名" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "TOTALWORKHR") ? "總計" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WORKDAYS") ? "出勤天數" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "AVGWORKHR") ? "平均每日時數" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DOCTORRISKLEVEL") ? "職醫評估風險等級" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().IndexOf("WorkTime") >= 0) ? dtExcel.Columns[ColumnsIndex].ColumnName.Replace("WorkTime", "") :
                                      "";
                        wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csTitle;
                    }
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ColumnsIndex = 2; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                        {
                            if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WORKDAYS")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csData_Int;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "AVGWORKHOUR")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csData_Float;
                            }
                            else
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csData;
                            }
                        }
                    }
                    //寫入小計
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("合計：");
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData_Right;
                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellFormula("CountIF(B2:B" + vLinesNo.ToString() + ",\"<>\"\"\")");
                    wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csData_Int;
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
                            string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                            string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                                 "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
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
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('找不到符合條件的資料！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('行車日期年月不可空白！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 匯出各單位駕駛員工時明細 (依單位切工作表)
        /// </summary>
        private void ExportExcel_Detail2()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if ((eDriveYear_Search.Text.Trim() != "") && (eDriveMonth_Search.Text.Trim() != ""))
            {
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

                HSSFCellStyle csData_Right = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Right.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Right.Alignment = HorizontalAlignment.Right;

                HSSFFont fontData_Right = (HSSFFont)wbExcel.CreateFont();
                fontData_Right.FontHeightInPoints = 12;
                csData_Right.SetFont(fontData);

                HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
                fontData_Red.FontHeightInPoints = 12; //字體大小
                fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
                csData_Red.SetFont(fontData_Red);

                HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Int.DataFormat = format.GetFormat("###,##0");

                HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Float.DataFormat = format.GetFormat("##0.00");

                int vLinesNo = 0;
                string vFileName = eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月各單位駕駛員工時明細表(依單位)";
                string vRiskDepNo = "";
                string vGroupText = "";
                string vHeaderText = "";
                string vCellData = "";

                DataTable dtExcel = CalWorkHourData_Detail();
                if (dtExcel.Rows.Count > 0)
                {
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        if (dtExcel.Rows[i]["DepNo"].ToString().Trim() != vRiskDepNo)
                        {
                            if (vGroupText != "")
                            {
                                //單位有改變時先寫入前一個單位的小計
                                vLinesNo++;
                                wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("合計：");
                                wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData_Right;
                                wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellFormula("CountIF(B2:B" + vLinesNo.ToString() + ",\"<>\"\"\")");
                                wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csData_Int;
                            }
                            vRiskDepNo = dtExcel.Rows[i]["DepNo"].ToString().Trim();
                            vLinesNo = 0; //列數歸零
                                          //依單位分別寫進不同的工作表
                            vGroupText = dtExcel.Rows[i]["DepNo"].ToString().Trim() + "_" + dtExcel.Rows[i]["DepName"].ToString().Trim();
                            //建立新的工作表
                            wsExcel = (HSSFSheet)wbExcel.CreateSheet(vGroupText);
                            //建立標題列
                            wsExcel.CreateRow(vLinesNo);
                            for (int ColumnsIndex = 2; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                            {
                                vHeaderText = (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DRIVER") ? "工號" :
                                              (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "NAME") ? "姓名" :
                                              (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "TOTALWORKHR") ? "總計" :
                                              (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WORKDAYS") ? "出勤天數" :
                                              (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "AVGWORKHR") ? "平均每日時數" :
                                              (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DOCTORRISKLEVEL") ? "職醫評估風險等級" :
                                              (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().IndexOf("WorkTime") >= 0) ? dtExcel.Columns[ColumnsIndex].ColumnName.Replace("WorkTime", "") :
                                              "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csTitle;
                            }
                        }
                        vLinesNo++;
                        wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                        wsExcel.CreateRow(vLinesNo);
                        for (int ColumnsIndex = 2; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                        {
                            if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WORKDAYS")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csData_Int;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "AVGWORKHOUR")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csData_Float;
                            }
                            else
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex - 2).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex - 2).CellStyle = csData;
                            }
                        }
                    }
                    //寫入最後一個單位的小計
                    vLinesNo++;
                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("合計：");
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData_Right;
                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellFormula("CountIF(B2:B" + vLinesNo.ToString() + ",\"<>\"\"\")");
                    wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csData_Int;
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
                            string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                            string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                                 "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
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
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('找不到符合條件的資料！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('行車日期年月不可空白！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 匯出各單位駕駛員月工時統計表
        /// </summary>
        private void ExportExcel_Month()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if ((eDriveYear_Search.Text.Trim() != "") && (eDriveMonth_Search.Text.Trim() != ""))
            {
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

                HSSFCellStyle csData_Right = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Right.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Right.Alignment = HorizontalAlignment.Right;

                HSSFFont fontData_Right = (HSSFFont)wbExcel.CreateFont();
                fontData_Right.FontHeightInPoints = 12;
                csData_Right.SetFont(fontData);

                HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
                fontData_Red.FontHeightInPoints = 12; //字體大小
                fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
                csData_Red.SetFont(fontData_Red);

                HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Int.DataFormat = format.GetFormat("###,##0");

                HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Float.DataFormat = format.GetFormat("##0.00");

                int vLinesNo = 0;
                string vFileName = eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月各單位駕駛員月工時統計表";
                string vHeaderText = "";
                string vCellData = "";

                DataTable dtExcel = CalWorkHourData_Month();
                if (dtExcel.Rows.Count > 0)
                {
                    //建立新的工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //建立標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                    {
                        vHeaderText = (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DEPNO") ? "單位代碼" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DEPNAME") ? "單位" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "EMPNO") ? "工號" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "EMPNAME") ? "姓名" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "MONTHWORKHR") ? "月工時" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WORKDAYS") ? "出勤天數" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "AVGWORKHOUR") ? "平均每日時數" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DOCTORRISKLEVEL") ? "職醫評估風險等級" : "";
                        wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csTitle;
                    }
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                        {
                            if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WORKDAYS")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Int;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "AVGWORKHOUR")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Float;
                            }
                            else
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData;
                            }
                        }
                    }
                    //寫入小計
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("合計：");
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 2));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData_Right;
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellFormula("CountIF(D2:D" + vLinesNo.ToString() + ",\"<>\"\"\")");
                    wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csData_Int;
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
                            string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                            string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                                 "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
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
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('找不到符合條件的資料！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('行車日期年月不可空白！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 按月匯出駕駛員工時情況
        /// </summary>
        private void ExportExcel_Status()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if ((eDriveYear_Search.Text.Trim() != "") && (eDriveMonth_Search.Text.Trim() != ""))
            {
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

                HSSFCellStyle csData_Right = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Right.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Right.Alignment = HorizontalAlignment.Right;

                HSSFFont fontData_Right = (HSSFFont)wbExcel.CreateFont();
                fontData_Right.FontHeightInPoints = 12;
                csData_Right.SetFont(fontData);

                HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
                fontData_Red.FontHeightInPoints = 12; //字體大小
                fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
                csData_Red.SetFont(fontData_Red);

                HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Int.DataFormat = format.GetFormat("###,##0");

                HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Float.DataFormat = format.GetFormat("##0.00");

                int vLinesNo = 0;
                string vFileName = eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月各單位駕駛員工時情況";
                string vHeaderText = "";
                string vCellData = "";

                DataTable dtExcel = CalWorkHourData_Status();
                if (dtExcel.Rows.Count > 0)
                {
                    //建立新的工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //建立標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                    {
                        vHeaderText = (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DEPNAME") ? "單位" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DRIVER") ? "工號" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DRIVERNAME") ? "姓名" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "LOWOVERDUTY") ? "未滿12小時" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "MIDDLEOVERDUTY") ? "12至15小時" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "HIGHOVERDUTY") ? "超過15小時" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "MONTHWORKHR") ? "月工時" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "HRRANGE_1") ? "0-222" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "HRRANGE_2") ? "223-300" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "HRRANGE_3") ? "301-350" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "HRRANGE_4") ? "351-400" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "HRRANGE_5") ? "401以上" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WORKDAYS") ? "出勤天數" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "AVGHR") ? "平均每日時數" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "RISKLEVEL") ? "職醫評估風險等級" : "";
                        wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csTitle;
                    }
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                        {
                            if ((dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "LOWOVERDUTY") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "MIDDLEOVERDUTY") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "HIGHOVERDUTY") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WORKDAYS"))
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Int;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "AVGHR")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Float;
                            }
                            else
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData;
                            }
                        }
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
                            string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                            string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                                 "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
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
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('找不到符合條件的資料！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('行車日期年月不可空白！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 駕駛員加班時數計算
        /// </summary>
        private void ExportExcel_Overduty()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if ((eDriveYear_Search.Text.Trim() != "") && (eDriveMonth_Search.Text.Trim() != ""))
            {
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

                HSSFCellStyle csData_Right = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Right.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Right.Alignment = HorizontalAlignment.Right;

                HSSFFont fontData_Right = (HSSFFont)wbExcel.CreateFont();
                fontData_Right.FontHeightInPoints = 12;
                csData_Right.SetFont(fontData);

                HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
                fontData_Red.FontHeightInPoints = 12; //字體大小
                fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
                csData_Red.SetFont(fontData_Red);

                HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Int.DataFormat = format.GetFormat("###,##0");

                HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Float.DataFormat = format.GetFormat("##0.0000");

                int vLinesNo = 0;
                string vFileName = eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月各單位駕駛員加班時數情況";
                string vHeaderText = "";
                string vCellData = "";

                DataTable dtExcel = CalWorkHourData_Overduty();
                if (dtExcel.Rows.Count > 0)
                {
                    //建立新的工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //建立標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                    {
                        vHeaderText = (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DEPNO") ? "單位代碼" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DEPNAME") ? "站別" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DRIVER") ? "工號" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "DRIVERNAME") ? "姓名" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "LIVEAMT_T") ? "駐在宿泊" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "RENTOVERAMT_T") ? "租車加班費" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "OVERDUTY_1") ? "平日1.33" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "OVERDUTY_2") ? "平日1.66" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "OVERDUTY_3") ? "休息日1.33" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "OVERDUTY_4") ? "休息日1.66" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "OVERDUTY_5") ? "休息日2.66" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "OVERDUTY_6") ? "例假日1.33" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "OVERDUTY_7") ? "例假日1.66" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "OVERDUTY_8") ? "國定假日1.33" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "OVERDUTY_9") ? "國定假日1.66" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WSTATE1") ? "休息日出勤" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WSTATE2") ? "例假日出勤" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WSTATE3") ? "國定假日出勤" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "CASHNUM01") ? "免稅加班" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "CASHNUM02") ? "應稅加班" :
                                      (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "TIMEPAY") ? "時薪" : "";
                        wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csTitle;
                    }
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                        {
                            if ((dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "LIVEAMT_T") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "RENTOVERAMT_T") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WSTATE1") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WSTATE2") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "WSTATE3") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "CASHNUM01") ||
                                (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "CASHNUM02"))
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Int;
                            }
                            else if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "TIMEPAY")
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Float;
                            }
                            else
                            {
                                vCellData = (dtExcel.Rows[i][ColumnsIndex].ToString().Trim() != "") ? dtExcel.Rows[i][ColumnsIndex].ToString().Trim() : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).SetCellValue(vCellData);
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData;
                            }
                        }
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
                            string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                            string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "DriverWorkHourAnalysisMain.aspx" + Environment.NewLine +
                                                 "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
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
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('找不到符合條件的資料！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('行車日期年月不可空白！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCLoseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plSearch.Visible = true;
        }

        protected void ddlReportType_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (ddlReportType_Search.SelectedIndex)
            {
                case 4:
                case 5:
                    bbPrint.Enabled = false;
                    break;
                default:
                    bbPrint.Enabled = true;
                    break;
            }
        }
    }
}
