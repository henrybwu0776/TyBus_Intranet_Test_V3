using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace TyBus_Intranet_Test_V3
{
    public partial class PayCash65List : System.Web.UI.Page
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
        DateTime vToday = DateTime.Today;

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
                        plShowData.Visible = true;
                        eCalYear_Search.Text = (vToday.Month - 1 < 0) ? (vToday.Year - 1912).ToString() : (vToday.Year - 1911).ToString();
                        eCalMonth_Search.Text = (vToday.Month - 1 < 0) ? "12" : vToday.Month.ToString();
                    }
                    ListDataBind(0);
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

        private string ListDataBind(int CalMode)
        {
            string vSelectStr = "";
            DateTime vCalDate = DateTime.Parse(eCalYear_Search.Text.Trim() + "/" + eCalMonth_Search.Text.Trim() + "/01");
            string vStartDate = PF.TransDateString(vCalDate, "B");
            string vEndDate = PF.GetMonthLastDay(vCalDate, "B");

            vSelectStr = "select aa.DepNo, (select[Name] from department z where z.DepNo = aa.DepNo) as DepNo_L, " + Environment.NewLine +
                         "       aa.Driver, aa.EmpName, aa.EmpTitle, aa.Car_No, aa.Car_ID, aa.HapDate, aa.Content, " + Environment.NewLine +
                         "       aa.AMT_Duty, aa.AMT_FixAsk, aa.Pay_Day, (select cast(WorkType as varchar(4)) as WorkType from employee e5 where e5.EmpNo = aa.Driver) as WorkType " + Environment.NewLine +
                         "  from ( " + Environment.NewLine +
                         "        select a.DepNo, a.Driver, (select[Name] from employee e1 where e1.EmpNo = a.Driver) as EmpName, " + Environment.NewLine +
                         "               (select (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and CLASSNO = e2.TITLE) as EmpTitle from Employee e2 where e2.EmpNo = a.Driver) as EmpTitle, " + Environment.NewLine +
                         "               a.Car_No, a.Car_ID, a.Day1 as HapDate, (select Content from DutyCode dc where dc.DutyCode = a.DutyCode) as Content, " + Environment.NewLine +
                         "               a.AMT as AMT_Duty, cast(0 as float) as AMT_FixAsk, a.Day2 as Pay_Day " + Environment.NewLine +
                         "          from Duty a where a.Day2 between '" + vStartDate + "' and '" + vEndDate + "' and isnull(IsAbort,'X') = 'X' " + Environment.NewLine +
                         "         union all " + Environment.NewLine +
                         "        select c.DepNo, b.EmpNo as Driver, (select[Name] from employee e3 where e3.EmpNo = b.EmpNo) as EmpName, " + Environment.NewLine +
                         "               (select(select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and CLASSNO = e4.TITLE) as emptitle from employee e4 where e4.EmpNo = b.EmpNo) as EmpTitle, " + Environment.NewLine +
                         "               c.Car_No, c.Car_ID, c.DamDate as HapDate, c.DamPart as Content, cast(0 as float) as AMT_Duty, isnull(b.AMT, 0) as AMT_FixAsk, b.PayDate " + Environment.NewLine +
                         "          from FixAskB b left join FIXASK c on c.FixAskNo = b.FIxAskNo " + Environment.NewLine +
                         "         where PayDate between '" + vStartDate + "' and '" + vEndDate + "' " + Environment.NewLine +
                         "       ) aa Order by aa.DepNo, aa.Driver ";
            if (CalMode == 0)
            {
                sdsPayCash65List.SelectCommand = "";
                sdsPayCash65List.SelectCommand = vSelectStr;
                gridPayCash65List.DataBind();
                bbPrint.Visible = (gridPayCash65List.Rows.Count > 0);
            }
            return vSelectStr;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            ListDataBind(0);
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vSelectStr = ListDataBind(1);
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                DataTable dtPrint = new DataTable("PayCash65ListP");
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("PayCash65ListP", dtPrint);
                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\PayCash65ListP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plShowData.Visible = false;

                string vCalYMStr = eCalYear_Search.Text.Trim() + " 年 " + eCalMonth_Search.Text.Trim() + " 月";
                string vRecordNote = "預覽報表_精勤保安獎金減發名冊" + Environment.NewLine +
                                     "PayCash65List.aspx" + Environment.NewLine +
                                     "計算年月：" + vCalYMStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            rvPrint.LocalReport.DataSources.Clear();
            plPrint.Visible = false;
            plShowData.Visible = true;
        }
    }
}