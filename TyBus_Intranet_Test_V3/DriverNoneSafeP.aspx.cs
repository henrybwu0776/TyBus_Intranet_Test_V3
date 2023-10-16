using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverNoneSafeP : System.Web.UI.Page
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
        private DateTime vToday = DateTime.Today;

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
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        eCalYear_Search.Text = (vToday.Year - 1911).ToString();
                        eCalMonth_Start_Search.Text = vToday.Month.ToString();
                        if ((vLoginDepNo != "09") && (vLoginDepNo != "03") && (vLoginDepNo != "06"))
                        {
                            eDepNo_Start_Search.Text = vLoginDepNo;
                            eDepNo_Start_Search.Enabled = false;
                            eDepNo_End_Search.Text = "";
                            eDepNo_End_Search.Enabled = false;
                        }
                        else
                        {
                            eDepNo_Start_Search.Text = "";
                            eDepNo_Start_Search.Enabled = true;
                            eDepNo_End_Search.Text = "";
                            eDepNo_End_Search.Enabled = true;
                        }
                        plShowData.Visible = true;
                        plPrint.Visible = false;
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

        private string GetSelectStr()
        {
            string vResultStr = "";
            string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? "           and DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? "           and DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? "           and DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
            DateTime vStartDate;
            DateTime vEndDate;
            int vDays = 1;
            int vCalYear = (eCalYear_Search.Text.Trim() != "") ? (Int32.Parse(eCalYear_Search.Text.Trim()) < 1911) ? Int32.Parse(eCalYear_Search.Text.Trim()) + 1911 : Int32.Parse(eCalYear_Search.Text.Trim()) : vToday.Year;
            int vMonth_S = (eCalMonth_Start_Search.Text.Trim() != "") ? Int32.Parse(eCalMonth_Start_Search.Text.Trim()) : vToday.Month;
            int vMonth_E = ((eCalMonth_Start_Search.Text.Trim() != "") && (eCalMonth_End_Search.Text.Trim() != "")) ? Int32.Parse(eCalMonth_End_Search.Text.Trim()) :
                           ((eCalMonth_Start_Search.Text.Trim() != "") && (eCalMonth_End_Search.Text.Trim() == "")) ? Int32.Parse(eCalMonth_Start_Search.Text.Trim()) :
                           vToday.Month;
            for (int i = vMonth_S; i <= vMonth_E; i++)
            {
                vStartDate = new DateTime(vCalYear, i, 1);
                vDays = vStartDate.AddMonths(1).AddDays(-1).Day;
                vEndDate = vStartDate.AddMonths(1).AddDays(-1);
                if (vResultStr == "")
                {
                    vResultStr = "select EmpNo, [Name], DepNo, DepName, DepGroup, CalMonth, WorkDays, TotalKMs, SplDays, UseDays, " + Environment.NewLine +
                                 "       (ESC01 / 8) ESCDay01, (ESC02 / 8) ESCDay02, (ESC03 / 8) ESCDay03, (ESC04 / 8) ESCDay04, (ESC05 / 8) ESCDay05, (OtherESC / 8) OtherESCDay " + Environment.NewLine +
                                 "  from ( " + Environment.NewLine +
                                 "        select EmpNo, [Name], DepNo, (select[Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                 "               cast(" + i.ToString() + " as int) CalMonth, " + Environment.NewLine +
                                 "               (select count(AssignNo) from RunSheetA where Driver = e.EmpNo and BuDate between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "') WorkDays, " + Environment.NewLine +
                                 "               isnull((select sum(ActualKM) from RunSheetA  where Driver = e.EmpNo and BuDate between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) TotalKMs, " + Environment.NewLine +
                                 "               (select DepGroup from Department where DepNo = e.DepNo) DepGroup, " + Environment.NewLine +
                                 "               isnull((select SplDays from YearHoliday where Years = " + vCalYear.ToString() + " and EmpNo = e.EmpNo), 0) SplDays, " + Environment.NewLine +
                                 "               isnull((select UseDays from YearHoliday where Years = " + vCalYear.ToString() + " and EmpNo = e.EmpNo), 0) UseDays, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '01' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC01, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '02' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC02, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '03' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC03, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '04' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC04, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '05' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC05, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType not in ('01', '02', '03', '04', '05') and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) OtherESC " + Environment.NewLine +
                                 "          from Employee e " + Environment.NewLine +
                                 "         where [Type] = '20' and (LeaveDay is null or LeaveDay > '" + vEndDate.ToString("yyyy/MM/dd") + "') " + Environment.NewLine +
                                 "           and AssumeDay <= '" + vEndDate.ToString("yyyy/MM/dd") + "' " + Environment.NewLine + vWStr_DepNo +
                                 "       ) z " + Environment.NewLine +
                                 " where (z.WorkDays < 18) or (z.DepGroup = '1' and z.TotalKMs < 2500) or (z.DepGroup <> '1' and z.TotalKMs < 2000)";
                }
                else
                {
                    vResultStr = vResultStr + Environment.NewLine +
                                 "union all " + Environment.NewLine +
                                 "select EmpNo, [Name], DepNo, DepName, DepGroup, CalMonth, WorkDays, TotalKMs, SplDays, UseDays, " + Environment.NewLine +
                                 "       (ESC01 / 8) ESCDay01, (ESC02 / 8) ESCDay02, (ESC03 / 8) ESCDay03, (ESC04 / 8) ESCDay04, (ESC05 / 8) ESCDay05, (OtherESC / 8) OtherESCDay " + Environment.NewLine +
                                 "  from ( " + Environment.NewLine +
                                 "        select EmpNo, [Name], DepNo, (select[Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                 "               cast(" + i.ToString() + " as int) CalMonth, " + Environment.NewLine +
                                 "               (select count(AssignNo) from RunSheetA where Driver = e.EmpNo and BuDate between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "') WorkDays, " + Environment.NewLine +
                                 "               isnull((select sum(ActualKM) from RunSheetA  where Driver = e.EmpNo and BuDate between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) TotalKMs, " + Environment.NewLine +
                                 "               (select DepGroup from Department where DepNo = e.DepNo) DepGroup, " + Environment.NewLine +
                                 "               isnull((select SplDays from YearHoliday where Years = " + vCalYear.ToString() + " and EmpNo = e.EmpNo), 0) SplDays, " + Environment.NewLine +
                                 "               isnull((select UseDays from YearHoliday where Years = " + vCalYear.ToString() + " and EmpNo = e.EmpNo), 0) UseDays, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '01' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC01, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '02' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC02, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '03' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC03, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '04' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC04, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType = '05' and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) ESC05, " + Environment.NewLine +
                                 "               isnull((select sum([Hours]) from ESCDuty where ApplyMan = e.EmpNo and ESCType not in ('01', '02', '03', '04', '05') and realDay between '" + vStartDate.ToString("yyyy/MM/dd") + "' and '" + vEndDate.ToString("yyyy/MM/dd") + "'), 0) OtherESC " + Environment.NewLine +
                                 "          from Employee e " + Environment.NewLine +
                                 "         where [Type] = '20' and (LeaveDay is null or LeaveDay > '" + vEndDate.ToString("yyyy/MM/dd") + "') " + Environment.NewLine +
                                 "           and AssumeDay <= '" + vEndDate.ToString("yyyy/MM/dd") + "' " + Environment.NewLine + vWStr_DepNo +
                                 "       ) z " + Environment.NewLine +
                                 " where (z.WorkDays < 18) or (z.DepGroup = '1' and z.TotalKMs < 2500) or (z.DepGroup <> '1' and z.TotalKMs < 2000)";
                }
            }
            vResultStr = vResultStr + Environment.NewLine + " order by z.DepNo, z.EmpNo, z.CalMonth";
            return vResultStr;
        }

        private void ListDataBind()
        {
            string vSelectStr = GetSelectStr();
            sdsDriverNoneSafeList.SelectCommand = "";
            sdsDriverNoneSafeList.SelectCommand = vSelectStr;
            gridDriverNoneSafeList.DataBind();
        }

        protected void eDepNo_Start_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_Start_Search.Text.Trim();
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Start_Search.Text = vDepNo_Temp.Trim();
            eDepName_Start_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDepNo_End_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_End_Search.Text.Trim();
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_End_Search.Text = vDepNo_Temp.Trim();
            eDepName_End_Search.Text = vDepName_Temp.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            ListDataBind();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            //統計報表
            int vCalYear = (eCalYear_Search.Text.Trim() != "") ? (Int32.Parse(eCalYear_Search.Text.Trim()) < 1911) ? Int32.Parse(eCalYear_Search.Text.Trim()) + 1911 : Int32.Parse(eCalYear_Search.Text.Trim()) : vToday.Year;
            int vMonth_S = (eCalMonth_Start_Search.Text.Trim() != "") ? Int32.Parse(eCalMonth_Start_Search.Text.Trim()) : vToday.Month;
            int vMonth_E = ((eCalMonth_Start_Search.Text.Trim() != "") && (eCalMonth_End_Search.Text.Trim() != "")) ? Int32.Parse(eCalMonth_End_Search.Text.Trim()) :
                           ((eCalMonth_Start_Search.Text.Trim() != "") && (eCalMonth_End_Search.Text.Trim() == "")) ? Int32.Parse(eCalMonth_Start_Search.Text.Trim()) :
                           vToday.Month;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                string vSelectStr = GetSelectStr();
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("DriverNoneSafe", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\DriverNoneSafeP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CalYear", (vCalYear - 1911).ToString()));
                rvPrint.LocalReport.SetParameters(new ReportParameter("CalMonth_S", vMonth_S.ToString()));
                rvPrint.LocalReport.SetParameters(new ReportParameter("CalMonth_E", vMonth_E.ToString()));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plShowData.Visible = false;
                string vCalYearStr = (eCalYear_Search.Text.Trim() != "") ? eCalYear_Search.Text.Trim() + " 年 " : DateTime.Now.Year.ToString() + " 年 ";
                string vCalMonthStr = ((eCalMonth_Start_Search.Text.Trim() != "") && (eCalMonth_End_Search.Text.Trim() != "")) ?
                                      eCalMonth_Start_Search.Text.Trim() + " 月到 " + eCalMonth_End_Search.Text.Trim() + " 月" :
                                      ((eCalMonth_Start_Search.Text.Trim() != "") && (eCalMonth_End_Search.Text.Trim() == "")) ?
                                      eCalMonth_Start_Search.Text.Trim() + " 月" :
                                      ((eCalMonth_Start_Search.Text.Trim() == "") && (eCalMonth_End_Search.Text.Trim() != "")) ?
                                      eCalMonth_End_Search.Text.Trim() + " 月" : "";                
                string vDepNoStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_Start_Search.Text.Trim() + " 站到 " + eDepNo_End_Search.Text.Trim() + " 站" :
                                   ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? eDepNo_Start_Search.Text.Trim() + " 站" :
                                   ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_End_Search.Text.Trim() + " 站" : "全部站別";
                string vRecordNote = "預覽報表_駕駛員未領精勤獎金一覽表" + Environment.NewLine +
                                     "DriverNoneSafeP.aspx" + Environment.NewLine +
                                     "計算年月：" + vCalYearStr + vCalMonthStr + Environment.NewLine +
                                     "站別：" + vDepNoStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            ListDataBind();
            plPrint.Visible = false;
            plShowData.Visible = true;
        }
    }
}