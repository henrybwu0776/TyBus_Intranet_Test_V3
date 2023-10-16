using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverWorkHoursTCT : System.Web.UI.Page
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
                        if (Int32.Parse(vLoginDepNo) > 10)
                        {
                            sdsDepNoList.SelectCommand = "SELECT e.DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DepNo)) AS DepName " + Environment.NewLine +
                                                         "  FROM EmployeeDepNo e " + Environment.NewLine +
                                                         " WHERE e.EmpNo = '" + vLoginID + "' " + Environment.NewLine +
                                                         "   AND isnull(e.IsUsed, 0) = 1 " + Environment.NewLine +
                                                         " ORDER BY e.IsDefault DESC, e.DepNo";
                            eDepNo.Text = PF.GetValue(vConnStr, "select DepNo from EmployeeDepNo where EmpNo = '" + vLoginID + "' and IsDefault = 1", "DepNo");
                        }
                        else
                        {
                            sdsDepNoList.SelectCommand = "select cast('' as varchar) DepNo, cast('' as varchar) DepName " + Environment.NewLine +
                                                         " union all " + Environment.NewLine +
                                                         "select DepNo, [Name] as DepName " + Environment.NewLine +
                                                         "  from Department " + Environment.NewLine +
                                                         " where InSHReport = 'V'" + Environment.NewLine +
                                                         " order by DepNo ";
                            eDepNo.Text = "";
                        }
                        sdsDepNoList.DataBind();
                        ddlDepNo.SelectedIndex = 0;
                        eCalYear.Text = (vToday.Year > 1911) ? (vToday.Year - 1911).ToString().Trim() : vToday.Year.ToString().Trim();
                        eCalMonth.Text = vToday.Month.ToString("D2");
                        plReport.Visible = false;
                        plSearch.Visible = true;
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

        private string GetSelectStr_Report()
        {
            string vSelStr = "";
            string vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "   and z.DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriverNo.Text.Trim() != "") ? "   and z.Driver = '" + eDriverNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vStartDate_Curr = (Int32.Parse(eCalYear.Text.Trim()) + 1911).ToString("D4") + "/" + eCalMonth.Text.Trim() + "/01";
            string vStartDate_Last = (Int32.Parse(eCalYear.Text.Trim()) + 1910).ToString("D4") + "/" + eCalMonth.Text.Trim() + "/01";
            string vEndDate_Curr = PF.GetMonthLastDay(DateTime.Parse(vStartDate_Curr), "B");
            string vEndDate_Last = PF.GetMonthLastDay(DateTime.Parse(vStartDate_Last), "B");
            string vRetireCalDate = DateTime.Parse(vEndDate_Curr).AddDays(1).ToString("yyyy/MM/dd");
            string vWStr_Retire = (rbDriverStatus.SelectedValue == "0") ?
                                  " where 1 = 1" + Environment.NewLine :
                                  " where z.Driver in (select EmpNo from Employee " + Environment.NewLine +
                                  "                     where [Type] = '20' " + Environment.NewLine +
                                  "                       and (LeaveDay is null or LeaveDay > '" + vEndDate_Curr + "') " + Environment.NewLine +
                                  "                       and ((datediff(month,birthday,'" + vRetireCalDate + "') >= 654 and datediff(month,assumeday,'" + vRetireCalDate + "') >= 174) " + Environment.NewLine +
                                  "                            or(datediff(month, birthday, '" + vRetireCalDate + "') >= 714 and datediff(month, assumeday, '" + vRetireCalDate + "') >= 114) " + Environment.NewLine +
                                  "                            or(datediff(month, birthday, '" + vRetireCalDate + "') >= 774) " + Environment.NewLine +
                                  "                            or(datediff(month, assumeday, '" + vRetireCalDate + "') >= 294)))" + Environment.NewLine;

            if (rbSelectType.SelectedValue == "0")
            {
                vSelStr = "select z.DepNo, (select[Name] from Department where DepNo = z.DepNo) DepName, " + Environment.NewLine +
                          "       z.Driver, (select[Name] from Employee where EmpNo = z.Driver) DriverName, " + Environment.NewLine +
                          "       (select WorkType from EMployee where EmpNo = z.Driver) WorkType, " + Environment.NewLine +
                          "       (select isnull(right('000' + cast(Year(LeaveDay) - 1911 as nvarchar), 3) + '/' + " + Environment.NewLine +
                          "                      right('00' + cast(Month(LeaveDay) as nvarchar), 2) + '/' + " + Environment.NewLine +
                          "                      right('00' + cast(day(LeaveDay) as nvarchar), 2), '') " + Environment.NewLine +
                          "          from Employee where EmpNo = z.Driver)  as LeaveDay_C, " + Environment.NewLine +
                          "       sum(z.WorkTime_Current) Total_Current, " + Environment.NewLine +
                          "       sum(z.WorkTime_LastYear) Total_LastYear, " + Environment.NewLine +
                          "       sum(z.WorkTime_Current) - sum(z.WorkTime_LastYear) TotalDiff " + Environment.NewLine +
                          "  from (" + Environment.NewLine +
                          "        select DepNo, Driver, sum(WorkHR * 60 + WorkMin) WorkTime_Current, cast(0 as int) WorkTime_LastYear " + Environment.NewLine +
                          "          from RunSheetA " + Environment.NewLine +
                          "         where BuDate between '" + vStartDate_Curr + "' and '" + vEndDate_Curr + "' " + Environment.NewLine +
                          "         group by DepNo, Driver " + Environment.NewLine +
                          "         union all " + Environment.NewLine +
                          "        select DepNo, Driver, cast(0 as int) WorkTime_Current, sum(WorkHR * 60 + WorkMin) WorkTime_LastYear " + Environment.NewLine +
                          "          from RunSheetA " + Environment.NewLine +
                          "         where BuDate between '" + vStartDate_Last + "' and '" + vEndDate_Last + "' " + Environment.NewLine +
                          "         group by DepNo, Driver " + Environment.NewLine +
                          "       ) z " + Environment.NewLine + vWStr_Retire + vWStr_DepNo + vWStr_Driver +
                          " group by z.DepNo, z.Driver " + Environment.NewLine +
                          " order by z.DepNo, z.Driver";
            }
            else
            {
                vSelStr = "select z.DepNo, (select[Name] from Department where DepNo = z.DepNo) DepName, " + Environment.NewLine +
                          "       z.Driver, (select[Name] from Employee where EmpNo = z.Driver) DriverName, " + Environment.NewLine +
                          "       (select WorkType from EMployee where EmpNo = z.Driver) WorkType, " + Environment.NewLine +
                          "       (select isnull(right('000' + cast(Year(LeaveDay) - 1911 as nvarchar), 3) + '/' + " + Environment.NewLine +
                          "                      right('00' + cast(Month(LeaveDay) as nvarchar), 2) + '/' + " + Environment.NewLine +
                          "                      right('00' + cast(day(LeaveDay) as nvarchar), 2), '') " + Environment.NewLine +
                          "          from Employee where EmpNo = z.Driver)  as LeaveDay_C, " + Environment.NewLine +
                          "       sum(z.ActualKm_Current) Total_Current, " + Environment.NewLine +
                          "       sum(z.ActualKm_LastYear) Total_LastYear, " + Environment.NewLine +
                          "       sum(z.ActualKm_Current) - sum(z.ActualKm_LastYear) TotalDiff " + Environment.NewLine +
                          "  from (" + Environment.NewLine +
                          "        select DepNo, Driver, sum(ActualKm) ActualKm_Current, cast(0 as int) ActualKm_LastYear " + Environment.NewLine +
                          "          from RunSheetA " + Environment.NewLine +
                          "         where BuDate between '" + vStartDate_Curr + "' and '" + vEndDate_Curr + "' " + Environment.NewLine +
                          "         group by DepNo, Driver " + Environment.NewLine +
                          "         union all " + Environment.NewLine +
                          "        select DepNo, Driver, cast(0 as int) ActualKm_Current, sum(ActualKm) ActualKm_LastYear " + Environment.NewLine +
                          "          from RunSheetA " + Environment.NewLine +
                          "         where BuDate between '" + vStartDate_Last + "' and '" + vEndDate_Last + "' " + Environment.NewLine +
                          "         group by DepNo, Driver " + Environment.NewLine +
                          "       ) z " + Environment.NewLine + vWStr_Retire + vWStr_DepNo + vWStr_Driver +
                          " group by z.DepNo, z.Driver " + Environment.NewLine +
                          " order by z.DepNo, z.Driver";
            }
            return vSelStr;
        }

        private string GetSelectStr_Excel()
        {
            string vSelStr = "";
            string vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "   and z.DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriverNo.Text.Trim() != "") ? "   and z.Driver = '" + eDriverNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vStartDate_Curr = (Int32.Parse(eCalYear.Text.Trim()) + 1911).ToString("D4") + "/" + eCalMonth.Text.Trim() + "/01";
            string vStartDate_Last = (Int32.Parse(eCalYear.Text.Trim()) + 1910).ToString("D4") + "/" + eCalMonth.Text.Trim() + "/01";
            string vEndDate_Curr = PF.GetMonthLastDay(DateTime.Parse(vStartDate_Curr), "B");
            string vEndDate_Last = PF.GetMonthLastDay(DateTime.Parse(vStartDate_Last), "B");
            string vRetireCalDate = DateTime.Parse(vEndDate_Curr).AddDays(1).ToString("yyyy/MM/dd");
            string vWStr_Retire = (rbDriverStatus.SelectedValue == "0") ?
                                  " where 1 = 1" + Environment.NewLine :
                                  " where z.Driver in (select EmpNo from Employee " + Environment.NewLine +
                                  "                     where [Type] = '20' " + Environment.NewLine +
                                  "                       and (LeaveDay is null or LeaveDay >= '" + vStartDate_Curr + "') " + Environment.NewLine +
                                  "                       and ((datediff(month,birthday,'" + vRetireCalDate + "') >= 654 and datediff(month,assumeday,'" + vRetireCalDate + "') >= 174) " + Environment.NewLine +
                                  "                            or(datediff(month, birthday, '" + vRetireCalDate + "') >= 714 and datediff(month, assumeday, '" + vRetireCalDate + "') >= 114) " + Environment.NewLine +
                                  "                            or(datediff(month, birthday, '" + vRetireCalDate + "') >= 774) " + Environment.NewLine +
                                  "                            or(datediff(month, assumeday, '" + vRetireCalDate + "') >= 294)))" + Environment.NewLine;

            if (rbSelectType.SelectedValue == "0")
            {
                vSelStr = "select z.DepNo, (select[Name] from Department where DepNo = z.DepNo) DepName, " + Environment.NewLine +
                          "       z.Driver, (select[Name] from Employee where EmpNo = z.Driver) DriverName, " + Environment.NewLine +
                          "       (select WorkType from EMployee where EmpNo = z.Driver) WorkType, " + Environment.NewLine +
                          "       (select isnull(right('000' + cast(Year(LeaveDay) - 1911 as nvarchar), 3) + '/' + " + Environment.NewLine +
                          "                      right('00' + cast(Month(LeaveDay) as nvarchar), 2) + '/' + " + Environment.NewLine +
                          "                      right('00' + cast(day(LeaveDay) as nvarchar), 2), '') " + Environment.NewLine +
                          "          from Employee where EmpNo = z.Driver)  as LeaveDay_C, " + Environment.NewLine +
                          "       right('000' + cast(cast(sum(z.WorkTime_Current) as int) / 60 as nvarchar), 3) + ':' + right('00' + cast(cast(sum(z.WorkTime_Current) as int) % 60 as nvarchar), 2) Total_Current, " + Environment.NewLine +
                          "       right('000' + cast(cast(sum(z.WorkTime_LastYear) as int) / 60 as nvarchar), 3) + ':' + right('00' + cast(cast(sum(z.WorkTime_LastYear) as int) % 60 as nvarchar), 2) Total_LastYear, " + Environment.NewLine +
                          "       case when sum(z.WorkTime_Current) >= sum(z.WorkTime_LastYear) then " + Environment.NewLine +
                          "                 cast(cast((sum(z.WorkTime_Current) - sum(z.WorkTime_LastYear)) as int) / 60 as nvarchar) + ':' + right('00' + cast(cast((sum(z.WorkTime_Current) - sum(z.WorkTime_LastYear)) as int) % 60 as nvarchar), 2) " + Environment.NewLine +
                          "            else '-' + cast(ABS(cast((sum(z.WorkTime_Current) - sum(z.WorkTime_LastYear)) as int)) / 60 as nvarchar) + ':' + right('00' + cast(cast((sum(z.WorkTime_Current) - sum(z.WorkTime_LastYear)) as int) % 60 as nvarchar), 2) " + Environment.NewLine +
                          "             end TotalDiff " + Environment.NewLine +
                          "  from (" + Environment.NewLine +
                          "        select DepNo, Driver, sum(WorkHR * 60 + WorkMin) WorkTime_Current, cast(0 as int) WorkTime_LastYear " + Environment.NewLine +
                          "          from RunSheetA " + Environment.NewLine +
                          "         where BuDate between '" + vStartDate_Curr + "' and '" + vEndDate_Curr + "' " + Environment.NewLine +
                          "         group by DepNo, Driver " + Environment.NewLine +
                          "         union all " + Environment.NewLine +
                          "        select DepNo, Driver, cast(0 as int) WorkTime_Current, sum(WorkHR * 60 + WorkMin) WorkTime_LastYear " + Environment.NewLine +
                          "          from RunSheetA " + Environment.NewLine +
                          "         where BuDate between '" + vStartDate_Last + "' and '" + vEndDate_Last + "' " + Environment.NewLine +
                          "         group by DepNo, Driver " + Environment.NewLine +
                          "       ) z " + Environment.NewLine + vWStr_Retire + vWStr_DepNo + vWStr_Driver +
                          " group by z.DepNo, z.Driver " + Environment.NewLine +
                          " order by z.DepNo, z.Driver";
            }
            else
            {
                vSelStr = "select z.DepNo, (select[Name] from Department where DepNo = z.DepNo) DepName, " + Environment.NewLine +
                          "       z.Driver, (select[Name] from Employee where EmpNo = z.Driver) DriverName, " + Environment.NewLine +
                          "       (select WorkType from EMployee where EmpNo = z.Driver) WorkType, " + Environment.NewLine +
                          "       (select isnull(right('000' + cast(Year(LeaveDay) - 1911 as nvarchar), 3) + '/' + " + Environment.NewLine +
                          "                      right('00' + cast(Month(LeaveDay) as nvarchar), 2) + '/' + " + Environment.NewLine +
                          "                      right('00' + cast(day(LeaveDay) as nvarchar), 2), '') " + Environment.NewLine +
                          "          from Employee where EmpNo = z.Driver)  as LeaveDay_C, " + Environment.NewLine +
                          "       sum(z.ActualKm_Current) Total_Current, " + Environment.NewLine +
                          "       sum(z.ActualKm_LastYear) Total_LastYear, " + Environment.NewLine +
                          "       sum(z.ActualKm_Current) - sum(z.ActualKm_LastYear) TotalDiff " + Environment.NewLine +
                          "  from (" + Environment.NewLine +
                          "        select DepNo, Driver, sum(ActualKm) ActualKm_Current, cast(0 as int) ActualKm_LastYear " + Environment.NewLine +
                          "          from RunSheetA " + Environment.NewLine +
                          "         where BuDate between '" + vStartDate_Curr + "' and '" + vEndDate_Curr + "' " + Environment.NewLine +
                          "         group by DepNo, Driver " + Environment.NewLine +
                          "         union all " + Environment.NewLine +
                          "        select DepNo, Driver, cast(0 as int) ActualKm_Current, sum(ActualKm) ActualKm_LastYear " + Environment.NewLine +
                          "          from RunSheetA " + Environment.NewLine +
                          "         where BuDate between '" + vStartDate_Last + "' and '" + vEndDate_Last + "' " + Environment.NewLine +
                          "         group by DepNo, Driver " + Environment.NewLine +
                          "       ) z " + Environment.NewLine + vWStr_Retire + vWStr_DepNo + vWStr_Driver +
                          " group by z.DepNo, z.Driver " + Environment.NewLine +
                          " order by z.DepNo, z.Driver";
            }
            return vSelStr;
        }

        protected void ddlDepNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDepNo.Text = ddlDepNo.SelectedValue.Trim();
        }

        protected void eDriverNo_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo_Temp = eDriverNo.Text.Trim();
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
            string vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo_Temp;
                vSQLStr = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName + "' order by EmpNo DESC";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eDriverNo.Text = vEmpNo_Temp;
            lbDriverName.Text = vEmpName;
            if (vEmpNo_Temp == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('查無指定駕駛員！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                string vSelectStr = GetSelectStr_Report();
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable("DriverWorkHoursTCTP");
                daPrint.Fill(dtPrint);
                if (dtPrint.Rows.Count > 0)
                {
                    string vReportName = (rbDriverStatus.SelectedValue == "0") ?
                                         eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月" + rbSelectType.SelectedItem.Text.Trim() + "兩期比較表" :
                                         eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月" + "屆退人員" + rbSelectType.SelectedItem.Text.Trim() + "兩期比較表";
                    ReportDataSource rdsPrint = new ReportDataSource("DriverWorkHoursTCTP", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    if (rbSelectType.SelectedValue == "0")
                    {
                        rvPrint.LocalReport.ReportPath = @"Report\DriverWorkHoursTCTP_HR.rdlc";
                    }
                    else
                    {
                        rvPrint.LocalReport.ReportPath = @"Report\DriverWorkHoursTCTP_KM.rdlc";
                    }
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", PF.GetValue(vConnStr, "select [Name] from [Custom] where Code = 'A000'", "Name")));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                    rvPrint.LocalReport.Refresh();
                    plReport.Visible = true;
                    plSearch.Visible = false;

                    string vRecordDepNoStr = (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : "全部";
                    string vRecordDriverStr = (eDriverNo.Text.Trim() != "") ? eDriverNo.Text.Trim() : rbDriverStatus.Text.Trim();
                    string vRecordNote = "列印資料_" + vReportName + Environment.NewLine +
                                         "DriverWorkHoursTCT.aspx" + Environment.NewLine +
                                         "行車日期：" + eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月" + Environment.NewLine +
                                         "站別：" + vRecordDepNoStr + Environment.NewLine +
                                         "駕駛員：" + vRecordDriverStr + Environment.NewLine +
                                         "統計區分：" + rbSelectType.Text.Trim();
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            double vTotal_Current = 0;
            double vTotal_LastYear = 0;
            double vTotalDiff = 0;
            string vTotalTimeStr = "";
            string[] vTempStrs;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                string vSelectStr = GetSelectStr_Excel();
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //
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

                    HSSFCellStyle csData_Right = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData_Right.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csData_Right.Alignment = HorizontalAlignment.Right;


                    HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
                    fontData.FontHeightInPoints = 12;
                    csData.SetFont(fontData);
                    csData_Right.SetFont(fontData);

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
                    string vFileName = (rbDriverStatus.SelectedValue == "0") ?
                                       eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月" + rbSelectType.SelectedItem.Text.Trim() + "兩期比較表" :
                                       eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月" + "屆退人員" + rbSelectType.SelectedItem.Text.Trim() + "兩期比較表";
                    int vLinesNo = 0;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員工號" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "部門代號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "部門" :
                                      (drExcel.GetName(i).ToUpper() == "TOTAL_LASTYEAR") ? "前期" + rbSelectType.SelectedItem.Text.Trim() :
                                      (drExcel.GetName(i).ToUpper() == "TOTAL_CURRENT") ? "本月" + rbSelectType.SelectedItem.Text.Trim() :
                                      (drExcel.GetName(i).ToUpper() == "TOTALDIFF") ? "兩期差異" :
                                      (drExcel.GetName(i).ToUpper() == "WORKTYPE") ? "在職狀況" :
                                      (drExcel.GetName(i).ToUpper() == "LEAVEDAY_C") ? "離職日期" : "";
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    while (drExcel.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if (rbSelectType.SelectedValue == "0")
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                vHeaderText = drExcel.GetName(i).Trim().ToUpper();
                                vTempStrs = drExcel[i].ToString().Split(':');
                                if (vHeaderText == "TOTAL_CURRENT")
                                {
                                    vTotal_Current += (Int32.Parse(vTempStrs[0]) * 60 + Int32.Parse(vTempStrs[1]));
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Right;
                                }
                                else if (vHeaderText == "TOTAL_LASTYEAR")
                                {
                                    vTotal_LastYear += (Int32.Parse(vTempStrs[0]) * 60 + Int32.Parse(vTempStrs[1]));
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Right;
                                }
                                else if (vHeaderText == "TOTALDIFF")
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Right;
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                }
                            }
                            else
                            {
                                vHeaderText = drExcel.GetName(i).Trim().ToUpper();
                                if (vHeaderText == "TOTAL_CURRENT")
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                    vTotal_Current += double.Parse(drExcel[i].ToString());
                                }
                                else if (vHeaderText == "TOTAL_LASTYEAR")
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                    vTotal_LastYear += double.Parse(drExcel[i].ToString());
                                }
                                else if (vHeaderText == "TOTALDIFF")
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                    vTotalDiff += double.Parse(drExcel[i].ToString());
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                }
                            }
                        }
                    }
                    //新增一筆小結
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        if (i == 0)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("小計");
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                        }
                        else
                        {
                            vHeaderText = drExcel.GetName(i).Trim().ToUpper();
                            if (vHeaderText == "TOTAL_CURRENT")
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(i);
                                if (rbSelectType.SelectedValue == "0")
                                {
                                    vTotalTimeStr = ((Int32)vTotal_Current / 60).ToString() + ":" + (vTotal_Current % 60).ToString();
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vTotalTimeStr);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Right;
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vTotal_Current);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                }
                            }
                            else if (vHeaderText == "TOTAL_LASTYEAR")
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(i);
                                if (rbSelectType.SelectedValue == "0")
                                {
                                    vTotalTimeStr = ((Int32)vTotal_LastYear / 60).ToString() + ":" + (vTotal_LastYear % 60).ToString();
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vTotalTimeStr);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Right;
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vTotal_LastYear);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                }
                            }
                            else if (vHeaderText == "TOTALDIFF")
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(i);
                                if (rbSelectType.SelectedValue == "0")
                                {
                                    vTotalDiff = vTotal_Current - vTotal_LastYear;
                                    vTotalTimeStr = ((Int32)vTotalDiff >= 0) ?
                                                    ((Int32)vTotalDiff / 60).ToString() + ":" + ((Int32)vTotalDiff % 60).ToString("D2") :
                                                    "-" + ((Int32)(vTotalDiff * -1) / 60).ToString() + ":" + ((Int32)(vTotalDiff * -1) % 60).ToString("D2");
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vTotalTimeStr);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Right;
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vTotalDiff);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                }
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
                            string vRecordDepNoStr = (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : "全部";
                            string vRecordDriverStr = (eDriverNo.Text.Trim() != "") ? eDriverNo.Text.Trim() : rbDriverStatus.Text.Trim();
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "DriverWorkHoursTCT.aspx" + Environment.NewLine +
                                                 "行車日期：" + eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr + Environment.NewLine +
                                                 "統計區分：" + rbSelectType.Text.Trim();
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
            plSearch.Visible = true;
        }
    }
}