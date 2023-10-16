using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Threading;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class EmpESCList : System.Web.UI.Page
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
        private string vSHReport = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                Thread.CurrentThread.CurrentCulture = Cal;

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
                        plReport.Visible = false;
                        plShow.Visible = true;
                        plSearch.Visible = true;
                        //if (vLoginDepNo != "09")
                        vSHReport = PF.GetValue(vConnStr, "select InSHReport from Department where DepNo = '" + vLoginDepNo.Trim() + "' ", "InSHReport");
                        if (vSHReport.Trim() == "V")
                        {
                            eDepNo_Search.Text = vLoginDepNo;
                            eDepNo_Search.Enabled = false;
                            eDepName_Search.Text = vLoginDepName;
                            eYearMonth_Search.Text = PF.GetMonthFirstDay(DateTime.Today, "C");
                        }
                    }
                    else
                    {

                    }
                    string vBuildDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eYearMonth_Search.ClientID;
                    string vBuildDateScript = "window.open('" + vBuildDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eYearMonth_Search.Attributes["onClick"] = vBuildDateScript;
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

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = eDepName_Search.Text.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '" + vDepName_Temp.Trim() + "%' order by DepNo";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Search.Text = vDepNo_Temp.Trim();
            eDepName_Search.Text = vDepName_Temp.Trim();
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vTargetMonth = DateTime.Parse(PF.GetMonthFirstDay(DateTime.Parse(eYearMonth_Search.Text.Trim()), "C")); //取回輸入日期當月的第一天
            calShowdata.VisibleDate = vTargetMonth;
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            if (eYearMonth_Search.Text.Trim() != "")
            {
                plReport.Visible = true;
                plShow.Visible = false;
                plSearch.Visible = false;
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSelectStr = GetListSelectStr();
                string vCompanyName = "桃園汽車客運股份有限公司";
                string vReportName = "員工請假狀況月報";
                string vDepName = (eDepNo_Search.Text.Trim() != "") ? eDepName_Search.Text.Trim() : "";
                using (SqlConnection connPrint = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                    connPrint.Open();
                    DataTable dtPrint = new DataTable();
                    daPrint.Fill(dtPrint);
                    ReportDataSource rdsPrint = new ReportDataSource("EmpESCListP", dtPrint);
                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\EmpESCListP.rdlc";
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("DepName", vDepName));
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.Refresh();
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇查詢月份')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
            plShow.Visible = true;
            plSearch.Visible = true;
        }

        protected void calShowdata_DayRender(object sender, DayRenderEventArgs e)
        {
            if (eYearMonth_Search.Text.Trim() != "")
            {
                if (e.Day.IsOtherMonth)
                {
                    e.Cell.ForeColor = System.Drawing.Color.White;
                }
                else
                {
                    string vRealDayStr = e.Day.Date.Year.ToString("D4") + "/" + e.Day.Date.ToString("MM/dd");
                    string vSQLStr_Temp = GetCalSelectStr(vRealDayStr);

                    //取回要顯示的資料
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        if (drTemp.HasRows)
                        {
                            ListBox lbTemp = new ListBox();
                            int vIndex = 0;
                            while (drTemp.Read())
                            {
                                lbTemp.Items.Insert(vIndex, drTemp["DepName"].ToString() + "_" + drTemp["ApplyMan"].ToString() + "_" + drTemp["Name"].ToString() + "_" + drTemp["ESCType_C"].ToString());
                                vIndex++;
                            }
                            lbTemp.Rows = vIndex;
                            lbTemp.Attributes.Remove("Style");
                            lbTemp.Attributes.Add("Style", "Width: 95%; Height:120px");

                            e.Cell.Controls.Add(new LiteralControl("<br />"));
                            e.Cell.Controls.Add(lbTemp);
                        }
                        else
                        {
                            ListBox lbTemp = new ListBox();
                            lbTemp.Items.Insert(0, "");
                            lbTemp.Rows = 1;
                            lbTemp.Attributes.Remove("Style");
                            lbTemp.Attributes.Add("Style", "Width: 95%; Height:120px");

                            e.Cell.Controls.Add(new LiteralControl("<br />"));
                            e.Cell.Controls.Add(lbTemp);
                        }
                    }
                }
            }
        }

        private string GetCalSelectStr(string fShowDate)
        {
            string vWStr_RealDay = " where a.RealDay = '" + fShowDate + "' " + Environment.NewLine;
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "select case when a.DepNo = '22' then '桃公' " + Environment.NewLine +
                                "            when a.DepNo = '23' then '中公' " + Environment.NewLine +
                                "            when a.DepNo = '07' then '桃車' " + Environment.NewLine +
                                "            when a.DepNo = '08' then '中車' " + Environment.NewLine +
                                "            when a.DepNo = '25' then '中壢' " + Environment.NewLine +
                                "            else left(c.[Name], 2) end DepName, " + Environment.NewLine +
                                "       a.ApplyMan, b.[Name], " + Environment.NewLine +
                                "       case when a.ESCType = '11' then '家照' " + Environment.NewLine +
                                "            when a.ESCType = '16' then '無薪防疫' " + Environment.NewLine +
                                "            when len(d.ClassTxt)> 2 then left(d.ClassTxt, 2) " + Environment.NewLine +
                                "            else left(d.ClassTxt, 1) end ESCType_C " + Environment.NewLine +
                                "  from ESCDuty a left join Employee b on b.EmpNo = a.ApplyMan " + Environment.NewLine +
                                "                 left join Department c on c.DepNo = a.DepNo " + Environment.NewLine +
                                "                 left join DBDICB d on d.ClassNo = a.ESCType and d.FKey = '請假資料檔      ESCDUTY         ESCTYPE'" + Environment.NewLine +
                                vWStr_RealDay +
                                vWStr_DepNo +
                                " order by a.DepNo, a.ApplyMan ";
            return vResultStr;
        }

        private string GetListSelectStr()
        {
            string vStartDate = PF.GetMonthFirstDay(DateTime.Parse(eYearMonth_Search.Text.Trim()), "B");
            string vEndDate = PF.GetMonthLastDay(DateTime.Parse(eYearMonth_Search.Text.Trim()), "B");
            string vWStr_RealDay = " where a.RealDay between '" + vStartDate + "' and '" + vEndDate + "' " + Environment.NewLine;
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "select c.[Name] DepName, a.RealDay, a.ApplyMan, b.[Name], d.ClassTxt ESCType_C " + Environment.NewLine +
                                "  from ESCDuty a left join Employee b on b.EmpNo = a.ApplyMan " + Environment.NewLine +
                                "                 left join Department c on c.DepNo = a.DepNo " + Environment.NewLine +
                                "                 left join DBDICB d on d.ClassNo = a.ESCType and d.FKey = '請假資料檔      ESCDUTY         ESCTYPE'" + Environment.NewLine +
                                vWStr_RealDay +
                                vWStr_DepNo +
                                " order by a.DepNo, a.RealDay, a.ApplyMan ";
            return vResultStr;
        }
    }
}