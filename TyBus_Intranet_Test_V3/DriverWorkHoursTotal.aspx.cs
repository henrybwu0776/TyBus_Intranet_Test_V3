using Amaterasu_Function;
using System;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverWorkHoursTotal : System.Web.UI.Page
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
                    //動態掛載日期輸入視窗
                    string vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eStopDate_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eStopDate_Search.Attributes["onClick"] = vDateScript;
                    if (!IsPostBack)
                    {
                        ClearAll();
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

        private void ClearAll()
        {
            eCalYear_Search.Text = (DateTime.Today.Year - 1911).ToString("D3");
            eCalMonth_Search.Text = DateTime.Today.Month.ToString("D2");
            eDriverNo_Search.Text = "";
            eDriverName_Search.Text = "";
            eStopDate_Search.Text = DateTime.Today.ToString("yyyy/MM/dd");
            plShowData.Visible = false;
            fvDriverWorkHoursTotalData.DataSourceID = "";
            eDriverNo_Search.Focus();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCalYM = (Int32.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString("D4") + "/" + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2");
            string vCalDate = vCalYM + "/01";
            DateTime vBreakDate_D = DateTime.Parse(eStopDate_Search.Text.Trim());
            if (DateTime.Parse(vCalDate).Month != vBreakDate_D.Month)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('截止日和計算月份必須相同！')");
                Response.Write("</" + "Script>");
                eCalMonth_Search.Focus();
            }
            else
            {
                string vDriverNo = eDriverNo_Search.Text.Trim();
                if (vDriverNo != "")
                {
                    string vBreakDate = vBreakDate_D.Year.ToString("D4") + "/" + vBreakDate_D.ToString("MM/dd");
                    string vCalYear = (Int32.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString("D4");
                    string vCalMonth = Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2");
                    string vTempStr = "";
                    string vSQLStr = "";
                    int vRCount = 0;
                    if (vLoginDepNo == "09")
                    {
                        vSQLStr = "select count(EmpNo) as RCount from Employee " + Environment.NewLine +
                                  " where EmpNo = '" + vDriverNo + "' and LeaveDay is null and Type = '20'";
                    }
                    else
                    {
                        vSQLStr = "select count(EmpNo) as RCount from Employee " + Environment.NewLine +
                                  " where EmpNo = '" + vDriverNo + "' and LeaveDay is null and Type = '20' and DepNo = '" + vLoginDepNo + "' ";
                    }
                    vTempStr = PF.GetValue(vConnStr, vSQLStr, "RCount");
                    vRCount = (vTempStr != "") ? Int32.Parse(vTempStr) : 0;
                    if (vRCount > 0)
                    {
                        //先計算駕駛員當月上班天數和當月可休天數
                        string vMonthStartDate = vCalDate; //當月第一天
                        string vMonthLastDate = PF.GetMonthLastDay(DateTime.Parse(vCalDate), "B"); //當月最後一天
                        int vMonthDays = DateTime.Parse(vMonthLastDate).Day; //當月天數

                        //應上時數上限
                        int vWorkHourLimit = 0;
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_WorkHourLimit' ";
                        vWorkHourLimit = int.Parse(PF.GetValue(vConnStr, vSQLStr, "Content"));
                        //休假加班時數上限
                        int vHoliOverLimit = 0;
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_HoliOverLimit' ";
                        vHoliOverLimit = int.Parse(PF.GetValue(vConnStr, vSQLStr, "Content"));
                        //1.33 加班時數上限
                        int vFrontOverLimit = 0;
                        vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_FrontOverLimit' ";
                        vFrontOverLimit = int.Parse(PF.GetValue(vConnStr, vSQLStr, "Content"));
                        //站休天數
                        int vBusHolidays = 0;
                        vSQLStr = "select Holiday from BusHoliday where Kind = '2' and BuDate = '" + vCalDate + "' ";
                        vBusHolidays = int.Parse(PF.GetValue(vConnStr, vSQLStr, "Holiday"));
                        //當月上班天數
                        int vWorkDays = 0;
                        vSQLStr = "select (DateDiff(Day, StartDay, EndDay) + 1) as WorkDays " + Environment.NewLine +
                                  "  from (" + Environment.NewLine +
                                  "        select case when AssumeDay between '" + vMonthStartDate + "' and '" + vBreakDate + "' then AssumeDay " + Environment.NewLine +
                                  "                    when BeginDay between '" + vMonthStartDate + "' and '" + vBreakDate + "' then BeginDay " + Environment.NewLine +
                                  "                    else '" + vMonthStartDate + "' end as StartDay, " + Environment.NewLine +
                                  "               case when StopDay between '" + vMonthStartDate + "' and '" + vBreakDate + "' then StopDay " + Environment.NewLine +
                                  "                    when LeaveDay between '" + vMonthStartDate + "' and '" + vBreakDate + "' then LeaveDay " + Environment.NewLine +
                                  "               else '" + vMonthLastDate + "' end as EndDay " + Environment.NewLine +
                                  "          from Employee " + Environment.NewLine +
                                  "         where EmpNo = '" + vDriverNo + "' " + Environment.NewLine +
                                  ") a";
                        vWorkDays = int.Parse(PF.GetValue(vConnStr, vSQLStr, "WorkDays"));
                        //駕駛員可休天數
                        int vDriverBreakDays = 0;
                        vSQLStr = "select convert(int, convert(float, Holiday) / convert(float, " + vMonthDays + ") * (datediff(day, StartDay, EndDay) + 1)) as QKDays " + Environment.NewLine +
                                  "  from (" + Environment.NewLine +
                                  "        select (select Holiday from BusHoliday where Kind = '2' and BuDate = '" + vCalDate + "') as Holiday, " + Environment.NewLine +
                                  "               case when AssumeDay between '" + vMonthStartDate + "' and '" + vBreakDate + "' then AssumeDay " + Environment.NewLine +
                                  "                    when BeginDay between '" + vMonthStartDate + "' and '" + vBreakDate + "' then BeginDay " + Environment.NewLine +
                                  "                    else '" + vMonthStartDate + "' end as StartDay, " + Environment.NewLine +
                                  "               case when StopDay between '" + vMonthStartDate + "' and '" + vBreakDate + "' then StopDay " + Environment.NewLine +
                                  "                    when LeaveDay between '" + vMonthStartDate + "' and '" + vBreakDate + "' then LeaveDay " + Environment.NewLine +
                                  "               else '" + vMonthLastDate + "' end as EndDay " + Environment.NewLine +
                                  "          from Employee " + Environment.NewLine +
                                  "         where EmpNo = '" + vDriverNo + "' " + Environment.NewLine +
                                  ") a";
                        vDriverBreakDays = int.Parse(PF.GetValue(vConnStr, vSQLStr, "QKDays"));
                        //計算連續未休天數
                        int vNOD = 0;
                        int vTempDays = 0;
                        for (int i = 0; i < vMonthDays; i++)
                        {
                            if (vNOD < 7)
                            {
                                vSQLStr = "select count(a.AssignNo) RCount " + Environment.NewLine +
                                          "  from RunSheetA a, RunSheetB b where b.AssignNo = a.AssignNo " + Environment.NewLine +
                                          "   and Driver = '" + vDriverNo + "' and isnull(b.Reducereason, '') = '' " +
                                          "   and BuDate = DateAdd(Day, " + i.ToString() + ", '" + vCalDate + "')";
                                vTempDays = int.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));
                                if (vTempDays == 0)
                                {
                                    vNOD = 0;
                                }
                                else
                                {
                                    vNOD++;
                                }
                            }
                        }
                        //標準工時 = (上班天數 - 駕駛員可休天數) * 8小時
                        int vStandardWH = (vWorkDays == vMonthDays) ? vWorkHourLimit :
                                          (((vWorkDays - vDriverBreakDays) * 8) > vWorkHourLimit) ? vWorkHourLimit : (vWorkDays - vDriverBreakDays) * 8;
                        //休假加班時數 (可休天數 * 8)
                        int vOverTime = (vWorkDays == vMonthDays) ? vHoliOverLimit :
                                        ((vDriverBreakDays * 8) > vHoliOverLimit) ? vHoliOverLimit : vDriverBreakDays * 8;
                        //月時數
                        int vOther = ((vStandardWH + vOverTime) > (vWorkHourLimit + vHoliOverLimit)) ? (vWorkHourLimit + vHoliOverLimit) : (vStandardWH + vOverTime);
                        //1.33加班時數
                        int vOverTime133 = vStandardWH + vOverTime + vFrontOverLimit;
                        vSQLStr = "select aa.Driver EmpNo, b.[Name] as EmpName, b.Birthday, b.Assumeday, " + Environment.NewLine +
                                  "       cast(" + vStandardWH + " as int) as HARM, cast(" + vOther + " as int) as [Other], '" + vNOD + "' as NotoffDay, " + Environment.NewLine +
                                  "       case when workhr - cast(" + vStandardWH + " as int) < 0 then " + Environment.NewLine +
                                  "            case when convert(int, workmin) % 60 = 0 then '-' + convert(varchar, cast(" + vStandardWH + " as int) - workhr) + ':00' " + Environment.NewLine +
                                  "                 else '-' + convert(varchar, cast(" + vStandardWH + " as int) - workhr - 1) + ':' + convert(varchar, 60 - convert(int, workmin)) end " + Environment.NewLine +
                                  "            when workhr - cast(" + vOther + " as int) >= 0 then '" + vHoliOverLimit + "' " + Environment.NewLine +
                                  "            else convert(varchar, workhr - cast(" + vStandardWH + " as int)) + ':' + " + Environment.NewLine +
                                  "                 case when convert(int, workmin) % 60 = 0 then '00' " + Environment.NewLine +
                                  "                      else convert(varchar, convert(int, workmin) % 60) end " + Environment.NewLine +
                                  "        end as [OverTime], " + Environment.NewLine +
                                  "       case when workhr - cast(" + vOther + " as int) < 0 then null " + Environment.NewLine +
                                  "            when workhr - cast(" + vOverTime133 + " as int) >= 0 then '" + vFrontOverLimit + "' " + Environment.NewLine +
                                  "            else convert(varchar, workhr - cast(" + vOther + " as int)) + ':' + " + Environment.NewLine +
                                  "                 case when convert(int, workmin) % 60 = 0 then '00' " + Environment.NewLine +
                                  "                      else convert(varchar, convert(int, workmin) % 60) end " + Environment.NewLine +
                                  "        end as [OverTime133], " + Environment.NewLine +
                                  "       case when workhr -cast(" + vOverTime133 + " as int) < 0 then null " + Environment.NewLine +
                                  "            else convert(varchar, workhr - cast(" + vOverTime133 + " as int)) + ':' + " + Environment.NewLine +
                                  "                 case when convert(int, workmin) % 60 = 0 then '00' " + Environment.NewLine +
                                  "                      else convert(varchar, convert(int, workmin) % 60) end " + Environment.NewLine +
                                  "        end as [OverTime166], " + Environment.NewLine +
                                  "       aa.workhr as WorkHR, aa.workmin as WorkMin, aa.rent as Rent, aa.QK, aa.leave as Leave, (aa.workhr + aa.leave) as WorkHR_T, " + Environment.NewLine +
                                  "       case when aa.driver in (select empno from RsEmpTotal where WORKHR >= 12 and budate = '" + vBreakDate + "') then '1' else '0' end as [OT] " + Environment.NewLine +
                                  "from ( " + Environment.NewLine +
                                  "      select driver, " + Environment.NewLine +
                                  "             sum(workhr) + sum(convert(int, workmin)) / 60 + (select isnull(sum(hours), 0) " + Environment.NewLine +
                                  "                                                                from Escduty " + Environment.NewLine +
                                  "                                                               where driver = applyman and datepart(month, realday) = '" + vCalMonth + "' " + Environment.NewLine +
                                  "                                                                 and datepart(year, realday) = '" + vCalYear + "') as workhr, " + Environment.NewLine +
                                  "             convert(int, sum(workmin)) % 60 as workmin, sum(convert(float, isnull(RentNumber, 0))) as rent, " + Environment.NewLine +
                                  "             datediff(day, '" + vCalDate + "', '" + vBreakDate + "') + 1 - count(driver) as QK, " + Environment.NewLine +
                                  "             (select isnull(sum(hours), 0) from Escduty " + Environment.NewLine +
                                  "               where driver = applyman and datepart(month, realday) = '" + vCalMonth + "' " + Environment.NewLine +
                                  "                 and datepart(year, realday) = '" + vCalYear + "') as leave " + Environment.NewLine +
                                  "        from runsheeta a " + Environment.NewLine +
                                  "       where datepart(month, budate) = '" + vCalMonth + "' and datepart(year, budate)= '" + vCalYear + "' " + Environment.NewLine +
                                  "         and convert(float, isnull(workhr, 0)) + convert(float, isnull(workmin, 0)) > 0 and budate <= '" + vBreakDate + "' " + Environment.NewLine +
                                  "         and Driver = '" + vDriverNo + "' " + Environment.NewLine +
                                  "       group by Driver " + Environment.NewLine +
                                  "     ) aa left join Employee b on b.EmpNo = aa.Driver";
                        sdsDriverWorkHouyrsTotalData.SelectCommand = "";
                        sdsDriverWorkHouyrsTotalData.SelectCommand = vSQLStr;
                        fvDriverWorkHoursTotalData.DataSourceID = "sdsDriverWorkHouyrsTotalData";
                        sdsDriverWorkHouyrsTotalData.DataBind();
                        plShowData.Visible = true;

                        string vRecordDriverStr = (eDriverNo_Search.Text.Trim() != "") ? eDriverNo_Search.Text.Trim() : "全部";
                        string vRecordStopDateStr = (eStopDate_Search.Text.Trim() != "") ? eStopDate_Search.Text.Trim() : "無指定";
                        string vRecordNote = "查詢資料_駕駛員工時累計查詢" + Environment.NewLine +
                                             "DriverWorkHoursTotal.aspx" + Environment.NewLine +
                                             "行車日期：" + eCalYear_Search.Text.Trim() + "年" + eCalMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                             "駕駛員：" + vRecordDriverStr + Environment.NewLine +
                                             "截止日期：" + vRecordStopDateStr;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('在單位 [ " + Session["DepName"].ToString() + " ] 找不到駕駛員 [ " + vDriverNo + " ] 的資料！')");
                        Response.Write("</" + "Script>");
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('駕駛員工號不可空白！')");
                    Response.Write("</" + "Script>");
                    eDriverNo_Search.Focus();
                }
            }
        }

        protected void bbClear_Click(object sender, EventArgs e)
        {
            ClearAll();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvDriverWorkHoursTotalData_DataBound(object sender, EventArgs e)
        {
            if (fvDriverWorkHoursTotalData.CurrentMode == FormViewMode.ReadOnly)
            {
                if (Int32.Parse(lbRecordCount.Text.Trim()) > 0)
                {
                    Label eTemp_DriveHR = (Label)fvDriverWorkHoursTotalData.FindControl("eDriveHR_List");
                    Label eTemp_Leave = (Label)fvDriverWorkHoursTotalData.FindControl("eLeave_List");
                    Label eTemp_WorkHR = (Label)fvDriverWorkHoursTotalData.FindControl("eWorkHR_List");
                    Label eTemp_WorkMin = (Label)fvDriverWorkHoursTotalData.FindControl("eWorkMin_List");
                    Label eTemp_DriverSum = (Label)fvDriverWorkHoursTotalData.FindControl("eDriverSum_List");
                    Label eTemp_TotalSum = (Label)fvDriverWorkHoursTotalData.FindControl("eStanderSum_List");
                    Label eTemp_WorkOverSevenDay = (Label)fvDriverWorkHoursTotalData.FindControl("eWorkOverSevenDay_List");
                    Label eTemp_NotoffDay = (Label)fvDriverWorkHoursTotalData.FindControl("eNotoffDay_List");
                    Label eTemp_OT = (Label)fvDriverWorkHoursTotalData.FindControl("eOT_List");
                    Label eTemp_HasOverTime = (Label)fvDriverWorkHoursTotalData.FindControl("eHasOverTime_List");
                    Label eTemp_Birthday = (Label)fvDriverWorkHoursTotalData.FindControl("eBirthday_List");
                    Label eTemp_Assumeday = (Label)fvDriverWorkHoursTotalData.FindControl("eAssumeday_List");
                    Label eTemp_Retire = (Label)fvDriverWorkHoursTotalData.FindControl("eRetire_List");
                    Label eTemp_Other = (Label)fvDriverWorkHoursTotalData.FindControl("eOther_List");

                    DateTime vCalDate = DateTime.Parse((Int32.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString("D4") + "/" + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2") + "/01");
                    DateTime vBreakDate = DateTime.Parse(eStopDate_Search.Text.Trim());
                    DateTime vBirthday = (eTemp_Birthday.Text.Trim() != "") ? DateTime.Parse(eTemp_Birthday.Text.Trim()) : DateTime.FromBinary(0);
                    DateTime vAssumeday = (eTemp_Assumeday.Text.Trim() != "") ? DateTime.Parse(eTemp_Assumeday.Text.Trim()) : DateTime.FromBinary(0);

                    //判定有沒有超過七天未休
                    eTemp_WorkOverSevenDay.Visible = (Int32.Parse(eTemp_NotoffDay.Text.Trim()) >= 7);
                    //判定有沒有超時
                    eTemp_HasOverTime.Visible = (eTemp_OT.Text.Trim() == "1");
                    //判定是不是屆退人員
                    int vWorkMonth = (eTemp_Assumeday.Text.Trim() != "") ? ((vBreakDate.Year - vAssumeday.Year) * 12) + (vBreakDate.Month - vAssumeday.Month) : 0;
                    int vBirthMonth = (eTemp_Birthday.Text.Trim() != "") ? ((vBreakDate.Year - vBirthday.Year) * 12) + (vBreakDate.Month - vBirthday.Month) : 0;
                    eTemp_Retire.Visible = ((vWorkMonth >= 300) || (vBirthMonth >= 780) || ((vWorkMonth >= 180) && (vBirthMonth >= 660)));

                    double vTotalWorkHR = (double)((int.Parse(eTemp_WorkHR.Text.Trim()) * 60) + int.Parse(eTemp_WorkMin.Text.Trim())) / 60.0;
                    string vMonthLastDay = PF.GetMonthLastDay(vCalDate, "B");
                    //當月天數
                    int vMonthDays = DateTime.Parse(vMonthLastDay).Day;
                    //天數
                    int vDayCount = 0;
                    vDayCount = (vCalDate.Month == vBreakDate.Month) ? vDayCount = vBreakDate.Day : vDayCount = vMonthDays;
                    //累計工時
                    double vWorkTotalMin = (double.Parse(eTemp_WorkHR.Text.Trim()) * 60 + double.Parse(eTemp_WorkMin.Text.Trim())) / 60;
                    //個人累計
                    double vDriverSum = Math.Round((vWorkTotalMin / double.Parse(eTemp_Other.Text.Trim())) * 100, 2, MidpointRounding.AwayFromZero);
                    eTemp_DriverSum.Text = vDriverSum.ToString() + " %";
                    //標準累計
                    double vTotalSum = Math.Round((double.Parse(eTemp_Other.Text.Trim()) / vMonthDays) * (vDayCount / double.Parse(eTemp_Other.Text.Trim())) * 100, 2, MidpointRounding.AwayFromZero);
                    eTemp_TotalSum.Text = vTotalSum.ToString() + " %";
                }
            }
        }

        protected void sdsDriverWorkHouyrsTotalData_Selected(object sender, SqlDataSourceStatusEventArgs e)
        {
            lbRecordCount.Text = e.AffectedRows.ToString();
        }
    }
}