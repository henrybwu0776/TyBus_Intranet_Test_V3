using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class OfficialDriverHours : System.Web.UI.Page
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

        private void ClearInsertPanel()
        {
            eDriverDate_Insert.Text = "";
            eDriver_Insert.Text = "";
            eDriverName_Insert.Text = "";
            eCalDate_Insert.Text = "";
            ddlWorkState_Insert.SelectedIndex = 0;
            eRemark_Insert.Text = "";
            eWorkHours_Insert.Text = "";
            eWorkMins_Insert.Text = "";
            eWorkTimes_Insert.Text = "";
            eExtra100_Insert.Text = "0";
            eExtra133_Insert.Text = "0";
            eExtra166_Insert.Text = "0";
            eExtra266_Insert.Text = "0";
            plInsert.Visible = false;
        }

        private void ClearEditPanel()
        {
            eDriverDate_Edit.Text = "";
            eDriver_Edit.Text = "";
            eDriverName_Edit.Text = "";
            eCalDate_Edit.Text = "";
            ddlWorkState_Edit.SelectedIndex = 0;
            eRemark_Edit.Text = "";
            eWorkHours_Edit.Text = "";
            eWorkMins_Edit.Text = "";
            eWorkTimes_Edit.Text = "";
            eExtra100_Edit.Text = "0";
            eExtra133_Edit.Text = "0";
            eExtra166_Edit.Text = "0";
            eExtra266_Edit.Text = "0";
            plEdit.Visible = false;
        }

        protected void bbCalculate_Insert_Click(object sender, EventArgs e)
        {
            double vWorkTime = double.Parse(eWorkHours_Insert.Text.Trim()) + (double.Parse(eWorkMins_Insert.Text.Trim()) / 60.0);
            eWorkTimes_Insert.Text = vWorkTime.ToString();
            //先全部歸零
            eExtra100_Insert.Text = "0";
            eExtra133_Insert.Text = "0";
            eExtra166_Insert.Text = "0";
            eExtra266_Insert.Text = "0";
            //再依不同的日期類別計算
            switch (ddlWorkState_Insert.SelectedValue.Trim())
            {
                case "":
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('請選擇日期類別！')");
                    Response.Write("</" + "Script>");
                    ddlWorkState_Insert.Focus();
                    break;
                case "0":
                    //平日
                    eExtra166_Insert.Text = (vWorkTime > 10.0) ? (vWorkTime - 10.0).ToString() : "0";
                    eExtra133_Insert.Text = ((vWorkTime > 8.0) && (vWorkTime <= 10.0)) ? (vWorkTime - 8.0).ToString() :
                                            (vWorkTime > 10.0) ? "2" : "0";
                    break;
                case "1":
                    //休息日
                    eExtra100_Insert.Text = (vWorkTime > 0.0) ? "8" : "0";
                    eExtra133_Insert.Text = (vWorkTime <= 2.0) ? vWorkTime.ToString() : "2";
                    eExtra166_Insert.Text = ((vWorkTime > 2.0) && (vWorkTime <= 8)) ? (vWorkTime - 2.0).ToString() :
                                            (vWorkTime > 8.0) ? "6" : "0";
                    eExtra266_Insert.Text = (vWorkTime > 8.0) ? (vWorkTime - 8.0).ToString() : "0";
                    break;
                case "2":
                    //例假日
                    eExtra100_Insert.Text = (vWorkTime > 0.0) ? "24" : "0";
                    eExtra133_Insert.Text = ((vWorkTime > 8.0) && (vWorkTime <= 10.0)) ? (vWorkTime - 8.0).ToString() :
                                            (vWorkTime > 10.0) ? "2" : "0";
                    eExtra166_Insert.Text = (vWorkTime > 10.0) ? (vWorkTime - 10.0).ToString() : "0";
                    break;
                case "3":
                    //國定假日
                    eExtra100_Insert.Text = (vWorkTime > 0.0) ? "16" : "0";
                    eExtra133_Insert.Text = (vWorkTime > 10.0) ? "2" :
                                            ((vWorkTime > 8.0) && (vWorkTime <= 10.0)) ? (vWorkTime - 8.0).ToString() : "0";
                    eExtra166_Insert.Text = (vWorkTime > 10.0) ? (vWorkTime - 10.0).ToString() : "0";
                    break;
            }
        }

        protected void bbInsert_Click(object sender, EventArgs e)
        {
            if (eDriverDate_Insert.Text.Trim() != "")
            {
                sdsOfficialDriverHoursDetail.InsertParameters.Clear();
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("DriverDate", DbType.String, eDriverDate_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("CalDate", DbType.Date, eCalDate_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("Driver", DbType.String, eDriver_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("WorkHours", DbType.Int32, eWorkHours_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("WorkMins", DbType.Int32, eWorkMins_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("TotalWorkHours", DbType.Double, eWorkTimes_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("Extra100", DbType.Double, eExtra100_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("Extra133", DbType.Double, eExtra133_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("Extra166", DbType.Double, eExtra166_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("Extra266", DbType.Double, eExtra266_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, eRemark_Insert.Text.Trim()));
                sdsOfficialDriverHoursDetail.InsertParameters.Add(new Parameter("WorkState", DbType.String, ddlWorkState_Insert.SelectedValue.Trim()));
                try
                {
                    sdsOfficialDriverHoursDetail.Insert();
                    ClearInsertPanel();
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

        protected void bbCancel_Insert_Click(object sender, EventArgs e)
        {
            ClearInsertPanel();
        }

        protected void bbCalculate_Edit_Click(object sender, EventArgs e)
        {
            double vWorkTime = double.Parse(eWorkHours_Edit.Text.Trim()) + (double.Parse(eWorkMins_Edit.Text.Trim()) / 60.0);
            eWorkTimes_Edit.Text = vWorkTime.ToString();
            //先全部歸零
            eExtra100_Edit.Text = "0";
            eExtra133_Edit.Text = "0";
            eExtra166_Edit.Text = "0";
            eExtra266_Edit.Text = "0";
            //再依不同的日期類別計算
            switch (ddlWorkState_Edit.SelectedValue.Trim())
            {
                case "":
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('請選擇日期類別！')");
                    Response.Write("</" + "Script>");
                    ddlWorkState_Edit.Focus();
                    break;
                case "0":
                    //平日
                    eExtra166_Edit.Text = (vWorkTime > 10.0) ? (vWorkTime - 10.0).ToString() : "0";
                    eExtra133_Edit.Text = ((vWorkTime > 8.0) && (vWorkTime <= 10.0)) ? (vWorkTime - 8.0).ToString() :
                                          (vWorkTime > 10.0) ? "2" : "0";
                    break;
                case "1":
                    //休息日
                    eExtra100_Edit.Text = (vWorkTime > 0.0) ? "8" : "0";
                    eExtra133_Edit.Text = (vWorkTime <= 2.0) ? vWorkTime.ToString() : "2";
                    eExtra166_Edit.Text = ((vWorkTime > 2.0) && (vWorkTime <= 8.0)) ? (vWorkTime - 2.0).ToString() :
                                          (vWorkTime > 8.0) ? "6" : "0";
                    eExtra266_Edit.Text = (vWorkTime > 8.0) ? (vWorkTime - 8.0).ToString() : "0";
                    break;
                case "2":
                    //例假日
                    eExtra100_Edit.Text = (vWorkTime > 0.0) ? "24" : "0";
                    eExtra133_Edit.Text = ((vWorkTime > 8.0) && (vWorkTime <= 10.0)) ? (vWorkTime - 8.0).ToString() :
                                          (vWorkTime > 10.0) ? "2" : "0";
                    eExtra166_Edit.Text = (vWorkTime > 10.0) ? (vWorkTime - 10.0).ToString() : "0";
                    break;
                case "3":
                    //國定假日
                    eExtra100_Edit.Text = (vWorkTime > 0.0) ? "16" : "0";
                    eExtra133_Edit.Text = (vWorkTime > 10.0) ? "2" :
                                          ((vWorkTime > 8.0) && (vWorkTime <= 10.0)) ? (vWorkTime - 8.0).ToString() : "0";
                    eExtra166_Edit.Text = (vWorkTime > 10.0) ? (vWorkTime - 10.0).ToString() : "0";
                    break;
            }
        }

        protected void bbModify_Click(object sender, EventArgs e)
        {
            if (eDriverDate_Edit.Text.Trim() != "")
            {
                sdsOfficialDriverHoursDetail.UpdateParameters.Clear();
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("DriverDate", DbType.String, eDriverDate_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("CalDate", DbType.Date, eCalDate_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("Driver", DbType.String, eDriver_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("WorkHours", DbType.Int32, eWorkHours_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("WorkMins", DbType.Int32, eWorkMins_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("TotalWorkHours", DbType.Double, eWorkTimes_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("Extra100", DbType.Double, eExtra100_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("Extra133", DbType.Double, eExtra133_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("Extra166", DbType.Double, eExtra166_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("Extra266", DbType.Double, eExtra266_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, eRemark_Edit.Text.Trim()));
                sdsOfficialDriverHoursDetail.UpdateParameters.Add(new Parameter("WorkState", DbType.String, ddlWorkState_Edit.SelectedValue.Trim()));
                try
                {
                    sdsOfficialDriverHoursDetail.Update();
                    ClearEditPanel();
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

        protected void bbCancel_Edit_Click(object sender, EventArgs e)
        {
            ClearEditPanel();
        }

        protected void ddlDriver_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDriver_Search.Text = ddlDriver_Search.SelectedValue.Trim();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "";
            string vTempDate = calOfficialDriverHours.VisibleDate.Year.ToString("D4");
            DateTime vPrintYM = (vTempDate != "0001") ? calOfficialDriverHours.VisibleDate : DateTime.Today;
            string vEmpNo = ddlDriver_Search.SelectedValue;
            if (vEmpNo != "")
            {
                //先取回當月份有沒有哪一天是沒有輸入資料的
                int vMonthDays = PF.GetMonthDays(vPrintYM);
                int vRCount = 0;
                string vSearchStr = "";
                string vTargetDays = "";
                for (int i = 1; i <= vMonthDays; i++)
                {
                    vTargetDays = vPrintYM.Year.ToString("D4") + "/" + vPrintYM.ToString("MM") + "/" + i.ToString("D2");
                    vSearchStr = "select count(Driver) RCount from OfficialDriverHours where Driver = '" + vEmpNo + "' and CalDate = '" + vTargetDays + "' ";
                    vRCount = int.Parse(PF.GetValue(vConnStr, vSearchStr, "RCount"));
                    if (vRCount == 0)
                    {
                        vSQLStr = "insert into OfficialDriverHours(DriverDate, CalDate, Driver, WorkHours, WorkMins, TotalWorkHours)" + Environment.NewLine +
                                  "values ('" + vEmpNo + DateTime.Parse(vTargetDays).ToString("yyyyMMdd") + "', '" + vTargetDays + "', '" + vEmpNo + "', 0,0,0)";
                        PF.ExecSQL(vConnStr, vSQLStr);
                    }
                }
                //開始產生報表
                vSQLStr = "select cast('" + (vPrintYM.Year - 1911).ToString("D3") + "年' as varchar) PrintYear, " + Environment.NewLine +
                          "       cast('" + vPrintYM.Month.ToString("D2") + "月份' as varchar) PrintMonth, " + Environment.NewLine +
                          "       Driver, (select [Name] from Employee where EmpNo = a.Driver) Driver_Name, " + Environment.NewLine +
                          "       CalDate, WorkHours, WorkMins, Extra100, Extra133, Extra166, Extra266, " + Environment.NewLine +
                          "       WorkState, (select ClassTxt from DBDICB where FKey = '行車記錄單A     runsheeta       WORKSTATE' and ClassNo = a.WorkState) WorkState_C " + Environment.NewLine +
                          "  from OfficialDriverHours a " + Environment.NewLine +
                          " where CalDate Between '" + PF.GetMonthFirstDay(vPrintYM, "B") + "' and '" + PF.GetMonthLastDay(vPrintYM, "B") + "' ";
                string vWStr_Driver = (vEmpNo != "") ? " and Driver = '" + vEmpNo + "' " : "";
                vSQLStr = vSQLStr + Environment.NewLine +
                          vWStr_Driver + Environment.NewLine +
                          " order by Driver, CalDate";
                sdsDriverHoursPrint.SelectCommand = "";
                sdsDriverHoursPrint.SelectCommand = vSQLStr;
                sdsDriverHoursPrint.DataBind();

                ReportDataSource rdsPrint = new ReportDataSource("OfficialDriverHoursP", sdsDriverHoursPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\OfficialDriverHoursP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請選擇駕駛員！')");
                Response.Write("</" + "Script>");
                ddlDriver_Search.Focus();
            }
        }

        /// <summary>
        /// 行事曆日期變更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void calOfficialDriverHours_SelectionChanged(object sender, EventArgs e)
        {
            plEdit.Visible = false;
            plInsert.Visible = false;
            string vSQLStr = "";
            string vWorkState = "";
            DateTime vSelectDate = calOfficialDriverHours.SelectedDate;

            if (eDriver_Search.Text.Trim() != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vTempCurrectDate = vSelectDate.Year.ToString("D4") + "/" + vSelectDate.ToString("MM/dd");
                vSQLStr = "select DriverDate, Driver, (select [Name] from Employee where EmpNo = a.Driver) DriverName, " + Environment.NewLine +
                          "       WorkHours, WorkMins, Extra100, Extra133, Extra166, Extra266, Remark, WorkState " + Environment.NewLine +
                          "  from OfficialDriverHours a " + Environment.NewLine +
                          " where a.Driver = '" + eDriver_Search.Text.Trim() + "' and a.CalDate = '" + vTempCurrectDate + "' ";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    if (drTemp.HasRows)
                    {
                        ddlWorkState_Edit.DataBind();
                        while (drTemp.Read())
                        {
                            eDriverDate_Edit.Text = drTemp["DriverDate"].ToString().Trim();
                            eDriver_Edit.Text = drTemp["Driver"].ToString().Trim();
                            eCalDate_Edit.Text = vTempCurrectDate;
                            eDriverName_Edit.Text = drTemp["DriverName"].ToString().Trim();
                            eWorkHours_Edit.Text = (drTemp["WorkHours"].ToString().Trim() != "") ? drTemp["WorkHours"].ToString().Trim() : "0";
                            eWorkMins_Edit.Text = (drTemp["WorkMins"].ToString().Trim() != "") ? drTemp["WorkMins"].ToString().Trim() : "0";
                            eRemark_Edit.Text = drTemp["Remark"].ToString().Trim();
                            eExtra100_Edit.Text = (drTemp["Extra100"].ToString().Trim() != "") ? drTemp["Extra100"].ToString().Trim() : "0";
                            eExtra133_Edit.Text = (drTemp["Extra133"].ToString().Trim() != "") ? drTemp["Extra133"].ToString().Trim() : "0";
                            eExtra166_Edit.Text = (drTemp["Extra166"].ToString().Trim() != "") ? drTemp["Extra166"].ToString().Trim() : "0";
                            eExtra266_Edit.Text = (drTemp["Extra266"].ToString().Trim() != "") ? drTemp["Extra266"].ToString().Trim() : "0";
                            vWorkState = (drTemp["WorkState"].ToString().Trim() != "") ? drTemp["WorkState"].ToString().Trim() : "0";
                            for (int i = 0; i < ddlWorkState_Edit.Items.Count; i++)
                            {
                                if (ddlWorkState_Edit.Items[i].Value.Trim() == vWorkState)
                                {
                                    ddlWorkState_Edit.SelectedIndex = i;
                                }
                            }
                        }
                        plEdit.Visible = true;
                    }
                    else
                    {
                        ddlWorkState_Insert.DataBind();
                        eDriver_Insert.Text = eDriver_Search.Text.Trim();
                        eCalDate_Insert.Text = vTempCurrectDate;
                        eDriverDate_Insert.Text = eDriver_Insert.Text.Trim() + DateTime.Parse(vTempCurrectDate).ToString("yyyyMMdd");
                        vSQLStr = "select [Name] from Employee where EmpNo = '" + eDriver_Search.Text.Trim() + "' and LeaveDay is null and IsOfficialDriver = 'V'";
                        eDriverName_Insert.Text = PF.GetValue(vConnStr, vSQLStr, "Name");
                        eWorkHours_Insert.Text = "0";
                        eWorkMins_Insert.Text = "0";
                        eRemark_Insert.Text = "";
                        eExtra100_Insert.Text = "0";
                        eExtra133_Insert.Text = "0";
                        eExtra166_Insert.Text = "0";
                        eExtra266_Insert.Text = "0";
                        ddlWorkState_Insert.SelectedIndex = 0;
                        plInsert.Visible = true;
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請選擇駕駛員！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void calOfficialDriverHours_DayRender(object sender, DayRenderEventArgs e)
        {
            if (e.Day.IsOtherMonth)
            {
                e.Cell.ForeColor = System.Drawing.Color.White;
            }
            if (eDriver_Search.Text.Trim() != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vTempCurrectDate = e.Day.Date.Year.ToString("D4") + "/" + e.Day.Date.ToString("MM/dd");
                string vSQLStr = "select (select [Name] from Employee where EmpNo = a.Driver) DriverName, " + Environment.NewLine +
                                 "       WorkHours, WorkMins, Extra100, Extra133, Extra166, Extra266 " + Environment.NewLine +
                                 "  from OfficialDriverHours a " + Environment.NewLine +
                                 " where Driver = '" + eDriver_Search.Text.Trim() + "' and CalDate = '" + vTempCurrectDate + "' ";

                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        ListBox lbTemp = new ListBox();
                        lbTemp.Items.Insert(0, drTemp["DriverName"].ToString());
                        lbTemp.Items.Insert(1, "上班時數：" + drTemp["WorkHours"].ToString() + ":" + drTemp["WorkMins"]);
                        lbTemp.Items.Insert(2, "加給日薪時數：" + drTemp["Extra100"].ToString());
                        lbTemp.Items.Insert(3, "1.33加班時數：" + drTemp["Extra133"].ToString());
                        lbTemp.Items.Insert(4, "1.66加班時數：" + drTemp["Extra166"].ToString());
                        lbTemp.Items.Insert(5, "2.66加班時數：" + drTemp["Extra266"].ToString());
                        lbTemp.Rows = 6;
                        lbTemp.Attributes.Remove("Style");
                        lbTemp.Attributes.Add("Style", "Width: 95%");

                        e.Cell.Controls.Add(new System.Web.UI.LiteralControl("<br />"));
                        e.Cell.Controls.Add(lbTemp);
                    }
                    cmdTemp.Cancel();
                    drTemp.Close();
                }
            }
            else
            {
                //Response.Write("<Script language='Javascript'>");
                //Response.Write("alert('請選擇駕駛員！')");
                //Response.Write("</" + "Script>");
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            rvPrint.LocalReport.DataSources.Clear();
            plPrint.Visible = false;
        }
    }
}