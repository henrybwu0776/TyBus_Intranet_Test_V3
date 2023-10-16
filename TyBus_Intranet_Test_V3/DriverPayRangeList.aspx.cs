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
using System.Web.UI;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverPayRangeList : System.Web.UI.Page
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
                        DateTime vCalDate = DateTime.Today.AddMonths(-1);
                        ePayYear_Search.Text = (vCalDate.Year - 1911).ToString();
                        ePayMonth_Search.SelectedIndex = vCalDate.Month - 1;
                        plShowData.Visible = true;
                        plReport.Visible = false;
                    }
                    else
                    {
                        OpenDataList();
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

        private string GetSelStr()
        {
            string vPayDate_Start = (Int32.Parse(ePayYear_Search.Text.Trim()) + 1911).ToString("D4") + "/" +
                                         (Int32.Parse(ePayMonth_Search.Text.Trim())).ToString("D2") + "/01";
            string vPayDate_End = PF.GetMonthLastDay(DateTime.Parse(vPayDate_Start), "B");
            string vPayDur = ePayDur_Search.SelectedValue.Trim();
            string vResultStr = "select z.DepNo, (select [Name] from Department where DepNo = z.DepNo) DepName , " + Environment.NewLine +
                                "       sum(z.DepCount) DepCount, " + Environment.NewLine +
                                "       sum(z.RCount_1) RCount_1, sum(z.RCount_2) RCount_2, sum(z.RCount_3) RCount_3, sum(z.RCount_4) RCount_4, sum(z.RCount_5) RCount_5, " + Environment.NewLine +
                                "       sum(z.RCount_6) RCount_6, sum(z.RCount_7) RCount_7, sum(z.RCount_8) RCount_8, sum(z.RCount_9) RCount_9 " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "       select p.DepNo, count(p.DepNo) DepCount, " + Environment.NewLine +
                                "              cast(0 as int) RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              count(EmpNo) RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash <= 24000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, count(EmpNo) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 24001 and 40000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, count(EmpNo) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 40001 and 50000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, count(EmpNo) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 50001 and 60000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, count(EmpNo) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 60001 and 70000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              count(EmpNo) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 70001 and 80000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, count(EmpNo) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 80001 and 90000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, count(EmpNo) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 90001 and 100000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, count(EmpNo) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash >= 100001 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                " ) z " + Environment.NewLine +
                                " group by z.DepNo " + Environment.NewLine +
                                " order by z.DepNo";
            return vResultStr;
        }

        private string GetSelStr_Excel()
        {
            string vPayDate_Start = (Int32.Parse(ePayYear_Search.Text.Trim()) + 1911).ToString("D4") + "/" +
                                    (Int32.Parse(ePayMonth_Search.Text.Trim())).ToString("D2") + "/01";
            string vPayDate_End = PF.GetMonthLastDay(DateTime.Parse(vPayDate_Start), "B");
            string vPayDur = ePayDur_Search.SelectedValue.Trim();
            string vResultStr = "select z.DepNo, (select [Name] from Department where DepNo = z.DepNo) DepName , " + Environment.NewLine +
                                "       sum(z.DepCount) DepCount, " + Environment.NewLine +
                                "       sum(z.RCount_1) RCount_1, (sum(z.RCount_1) / cast(sum(z.DepCount) as float)) Ratio_1, " + Environment.NewLine +
                                "       sum(z.RCount_2) RCount_2, (sum(z.RCount_2) / cast(sum(z.DepCount) as float)) Ratio_2, " + Environment.NewLine +
                                "       sum(z.RCount_3) RCount_3, (sum(z.RCount_3) / cast(sum(z.DepCount) as float)) Ratio_3, " + Environment.NewLine +
                                "       sum(z.RCount_4) RCount_4, (sum(z.RCount_4) / cast(sum(z.DepCount) as float)) Ratio_4, " + Environment.NewLine +
                                "       sum(z.RCount_5) RCount_5, (sum(z.RCount_5) / cast(sum(z.DepCount) as float)) Ratio_5, " + Environment.NewLine +
                                "       sum(z.RCount_6) RCount_6, (sum(z.RCount_6) / cast(sum(z.DepCount) as float)) Ratio_6, " + Environment.NewLine +
                                "       sum(z.RCount_7) RCount_7, (sum(z.RCount_7) / cast(sum(z.DepCount) as float)) Ratio_7, " + Environment.NewLine +
                                "       sum(z.RCount_8) RCount_8, (sum(z.RCount_8) / cast(sum(z.DepCount) as float)) Ratio_8, " + Environment.NewLine +
                                "       sum(z.RCount_9) RCount_9, (sum(z.RCount_9) / cast(sum(z.DepCount) as float)) Ratio_9 " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "       select p.DepNo, count(p.DepNo) DepCount, " + Environment.NewLine +
                                "              cast(0 as int) RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              count(EmpNo) RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash <= 24000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, count(EmpNo) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 24001 and 40000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, count(EmpNo) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 40001 and 50000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, count(EmpNo) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 50001 and 60000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, count(EmpNo) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 60001 and 70000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              count(EmpNo) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 70001 and 80000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, count(EmpNo) RCount_7, cast(0 as int) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 80001 and 90000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, count(EmpNo) RCount_8, cast(0 as int) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash between 90001 and 100000 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select p.DepNo, cast(0 as int) DepCount, " + Environment.NewLine +
                                "              cast(0 as int)RCount_1, cast(0 as int) RCount_2, cast(0 as int) RCount_3, cast(0 as int) RCount_4, cast(0 as int) RCount_5, " + Environment.NewLine +
                                "              cast(0 as int) RCount_6, cast(0 as int) RCount_7, cast(0 as int) RCount_8, count(EmpNo) RCount_9 " + Environment.NewLine +
                                "         from PAYREC p " + Environment.NewLine +
                                "        where p.Title = '300' " + Environment.NewLine +
                                "          and p.PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "          and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "          and p.GivCash >= 100001 " + Environment.NewLine +
                                "        group by p.DepNo " + Environment.NewLine +
                                " ) z " + Environment.NewLine +
                                " group by z.DepNo " + Environment.NewLine +
                                " order by z.DepNo";
            return vResultStr;
        }

        private void OpenDataList()
        {
            sdsDriverPayRangeList.SelectCommand = GetSelStr();
            gridDriverPayRangeList.DataBind();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenDataList();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vReportName = "各單位駕駛員薪資發放級距統計表";
            string vCalYM = "—" + ePayYear_Search.Text.Trim() + " 年 " + ePayMonth_Search.Text.Trim() + " 月第 " + ePayDur_Search.Text.Trim() + " 期";
            //統計報表
            string vSelectStr = GetSelStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("DriverPayRangeListP", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\DriverPayRangeListP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("CalYM", vCalYM));
                rvPrint.LocalReport.Refresh();
                plShowData.Visible = false;
                plReport.Visible = true;

                string vRecordNote = "列印資料_各單位駕駛員薪資發放級距統計表" + Environment.NewLine +
                                     "DriverPayRangeList.aspx" + Environment.NewLine +
                                     "統計年月：" + ePayYear_Search.Text.Trim() + "年" + ePayMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                     "期別：" + ePayDur_Search.Text.Trim();
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
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

            HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_Float.DataFormat = format.GetFormat("###,##0.00%");

            string vHeaderText = "";
            string vColumnName = "";
            string vFileName = "各單位駕駛員薪資發放級距統計表";
            int vLinesNo = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelStr_Excel();
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "DEPNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "所屬單位" :
                                      (drExcel.GetName(i).ToUpper() == "DEPCOUNT") ? "人數合計" :
                                      (drExcel.GetName(i).ToUpper() == "RCOUNT_1") ? "少於24000" :
                                      (drExcel.GetName(i).ToUpper() == "RATIO_1") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "RCOUNT_2") ? "24001-40000" :
                                      (drExcel.GetName(i).ToUpper() == "RATIO_2") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "RCOUNT_3") ? "40001-50000" :
                                      (drExcel.GetName(i).ToUpper() == "RATIO_3") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "RCOUNT_4") ? "50001-60000" :
                                      (drExcel.GetName(i).ToUpper() == "RATIO_4") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "RCOUNT_5") ? "60001-70000" :
                                      (drExcel.GetName(i).ToUpper() == "RATIO_5") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "RCOUNT_6") ? "70001-80000" :
                                      (drExcel.GetName(i).ToUpper() == "RATIO_6") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "RCOUNT_7") ? "8001-90000" :
                                      (drExcel.GetName(i).ToUpper() == "RATIO_7") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "RCOUNT_8") ? "9001-100000" :
                                      (drExcel.GetName(i).ToUpper() == "RATIO_8") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "RCOUNT_9") ? "100001以上" :
                                      (drExcel.GetName(i).ToUpper() == "RATIO_9") ? "" : drExcel.GetName(i);
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
                            vColumnName = drExcel.GetName(i);
                            if ((vColumnName.Length > 7) && (vColumnName.Substring(0, 7).ToUpper() == "RCOUNT_"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else if ((vColumnName.Length > 6) && (vColumnName.Substring(0, 6).ToUpper() == "RATIO_"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Float;
                            }
                            else if (drExcel.GetName(i).ToUpper() == "DEPCOUNT")
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
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
                            string vRecordNote = "匯出檔案_各單位駕駛員薪資發放級距統計表" + Environment.NewLine +
                                                 "DriverPayRangeList.aspx" + Environment.NewLine +
                                                 "統計年月：" + ePayYear_Search.Text.Trim() + "年" + ePayMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                 "期別：" + ePayDur_Search.Text.Trim();
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
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
            plShowData.Visible = true;
        }
    }
}