using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class AnecdoteCaseP : System.Web.UI.Page
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
                        eBuildDate_Start_Search.Text = (eBuildDate_Start_Search.Text.Trim() != "") ? eBuildDate_Start_Search.Text.Trim() : PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C");
                        eBuildDate_End_Search.Text = (eBuildDate_End_Search.Text.Trim() != "") ? eBuildDate_End_Search.Text.Trim() : PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C");

                        plShowData.Visible = true;
                        plReport.Visible = false;
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

        private string CaseDataBind_99()
        {
            string vSelectStr = "";
            if (vLoginID != "")
            {
                sdsAnecdoteCaseP_99.SelectCommand = "";
                string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? "   and DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_BuildDate = ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                         "   and BuildDate between '" + PF.TransDateString(eBuildDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eBuildDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() == "")) ?
                                         "   and BuildDate = '" + PF.TransDateString(eBuildDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eBuildDate_Start_Search.Text.Trim() == "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                         "   and BuildDate = '" + PF.TransDateString(eBuildDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";
                string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? " and Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
                vSelectStr = "select (select count(CaseNo) from AnecdoteCase where DepNo = a.DepNo " + vWStr_BuildDate + vWStr_DepNo + ") DepCount, " + Environment.NewLine +
                             "       (select count(CaseNo) from AnecdoteCase where DepNo = a.DepNo and HasInsurance = a.HasInsurance" + vWStr_BuildDate + vWStr_DepNo + ") InsuCount, " + Environment.NewLine +
                             "       a.HasInsurance, a.DepNo, a.DepName, a.BuildDate, a.Driver, a.DriverName, " + Environment.NewLine +
                             "       sum(isnull(b.ThirdInsurance, 0)) ThirdInsurance, sum(isnull(b.CompInsurance, 0)) CompInsurance, " + Environment.NewLine +
                             "       sum(isnull(b.PassengerInsu, 0)) PassengerInsu, a.CaseOccurrence " + Environment.NewLine +
                             "  from AnecdoteCase a left join AnecdoteCaseB b on b.CaseNo = a.CaseNo " + Environment.NewLine +
                             " where 1 = 1 " + Environment.NewLine +
                             vWStr_BuildDate +
                             vWStr_DepNo +
                             vWStr_Driver +
                             " group by a.CaseNo, a.DepNo, a.DepName, a.BuildDate, a.Driver, a.DriverName, a.HasInsurance, a.CaseOccurrence " + Environment.NewLine +
                             " order by a.DepNo, a.HasInsurance ";
                sdsAnecdoteCaseP_99.SelectCommand = vSelectStr;
                gridAnecdoteCaseP_99.DataBind();
            }
            return vSelectStr;
        }

        private string CaseDataBind_1(Boolean vHasInsu)
        {
            string vSelectStr = "";
            if (vLoginID != "")
            {
                sdsAnecdoteCaseP_1.SelectCommand = "";
                string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? "   and a.DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and a.DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_BuildDate = ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                         "   and a.BuildDate between '" + PF.TransDateString(eBuildDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eBuildDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() == "")) ?
                                         "   and a.BuildDate = '" + PF.TransDateString(eBuildDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eBuildDate_Start_Search.Text.Trim() == "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                         "   and a.BuildDate = '" + PF.TransDateString(eBuildDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";
                string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? " and a.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_Main = (vHasInsu) ? " where a.HasInsurance = 1 " + Environment.NewLine : " where a.HasInsurance = 0 " + Environment.NewLine;
                vSelectStr = "select case when a.HasInsurance = 1 then '已出險' else '未出險' end InsuranceState, " + Environment.NewLine +
                             "       a.DepNo, a.DepName, a.BuildDate, a.Driver, a.DriverName, b.RelCar_ID, b.Relationship, a.CaseOccurrence " + Environment.NewLine +
                             "  from AnecdoteCase a left join AnecdoteCaseB b on b.CaseNo = a.CaseNo " + Environment.NewLine +
                             vWStr_Main + vWStr_DepNo + vWStr_BuildDate + vWStr_Driver +
                             " order by a.DepNo, a.BuildDate, a.Driver ";
                sdsAnecdoteCaseP_1.SelectCommand = vSelectStr;
                gridAnecdoteCaseP_1.DataBind();
            }
            return vSelectStr;
        }

        private string CaseDataBind_2(Boolean vHasInsu)
        {
            string vSelectStr = "";
            if (vLoginID != "")
            {
                sdsAnecdoteCaseP_2.SelectCommand = "";
                string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? "   and a.DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? "   and a.DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_BuildDate = ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                         "   and a.BuildDate between '" + PF.TransDateString(eBuildDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eBuildDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() == "")) ?
                                         "   and a.BuildDate = '" + PF.TransDateString(eBuildDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                         ((eBuildDate_Start_Search.Text.Trim() == "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                         "   and a.BuildDate = '" + PF.TransDateString(eBuildDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";
                string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? " and a.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_Main = (vHasInsu) ? " where a.HasInsurance = 1 " + Environment.NewLine : " where a.HasInsurance = 0 " + Environment.NewLine;
                vSelectStr = "select case when a.HasInsurance = 1 then '已出險' else '未出險' end InsuranceState, " + Environment.NewLine +
                             "       a.DriverName, a.Driver, a.BuildDate, b.RelCar_ID, b.Relationship, a.CaseOccurrence " + Environment.NewLine +
                             "  from AnecdoteCase a left join AnecdoteCaseB b on b.CaseNo = a.CaseNo " + Environment.NewLine +
                             vWStr_Main + vWStr_DepNo + vWStr_BuildDate + vWStr_Driver +
                             " order by a.Driver, a.BuildDate ";
                sdsAnecdoteCaseP_2.SelectCommand = vSelectStr;
                gridAnecdoteCaseP_2.DataBind();
            }
            return vSelectStr;
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
            lbDepNo_Start_Search.Text = vDepName_Temp.Trim();
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
            lbDepNo_End_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            string vDriver = eDriver_Search.Text.Trim();
            string vDriverName = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vDriver + "' and LeaveDay is null";
            vDriverName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriver;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null";
                vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eDriver_Search.Text = vDriver;
            eDriverName_Search.Text = vDriverName;
        }

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            string vSelectStr = "";
            plShowData.Visible = true;
            plReport.Visible = false;
            gridAnecdoteCaseP_1.Visible = false;
            gridAnecdoteCaseP_2.Visible = false;
            gridAnecdoteCaseP_99.Visible = false;
            switch (rbSelectPrintMode.SelectedValue)
            {
                case "11":
                    vSelectStr = CaseDataBind_1(false);
                    gridAnecdoteCaseP_1.Visible = (gridAnecdoteCaseP_1.Rows.Count > 0);
                    bbPrint.Visible = (gridAnecdoteCaseP_1.Rows.Count > 0);
                    bbExcel.Visible = (gridAnecdoteCaseP_1.Rows.Count > 0);
                    break;
                case "12":
                    vSelectStr = CaseDataBind_1(true);
                    gridAnecdoteCaseP_1.Visible = (gridAnecdoteCaseP_1.Rows.Count > 0);
                    bbPrint.Visible = (gridAnecdoteCaseP_1.Rows.Count > 0);
                    bbExcel.Visible = (gridAnecdoteCaseP_1.Rows.Count > 0);
                    break;
                case "21":
                    vSelectStr = CaseDataBind_2(false);
                    gridAnecdoteCaseP_2.Visible = (gridAnecdoteCaseP_2.Rows.Count > 0);
                    bbPrint.Visible = (gridAnecdoteCaseP_2.Rows.Count > 0);
                    bbExcel.Visible = (gridAnecdoteCaseP_2.Rows.Count > 0);
                    break;
                case "22":
                    vSelectStr = CaseDataBind_2(true);
                    gridAnecdoteCaseP_2.Visible = (gridAnecdoteCaseP_2.Rows.Count > 0);
                    bbPrint.Visible = (gridAnecdoteCaseP_2.Rows.Count > 0);
                    bbExcel.Visible = (gridAnecdoteCaseP_2.Rows.Count > 0);
                    break;
                case "99":
                    vSelectStr = CaseDataBind_99();
                    gridAnecdoteCaseP_99.Visible = (gridAnecdoteCaseP_99.Rows.Count > 0);
                    bbPrint.Visible = (gridAnecdoteCaseP_99.Rows.Count > 0);
                    bbExcel.Visible = (gridAnecdoteCaseP_99.Rows.Count > 0);
                    break;
            }
        }

        private void PrintReport_1()
        {
            ReportDataSource rdsPrint = new ReportDataSource("AnecdoteCaseP", sdsAnecdoteCaseP_1);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\AnecdoteCaseP_DepNo.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
        }

        private void PrintReport_2()
        {
            ReportDataSource rdsPrint = new ReportDataSource("AnecdoteCaseP", sdsAnecdoteCaseP_2);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\AnecdoteCaseP_Driver.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
        }

        private void PrintReport99()
        {
            //統計報表
            ReportDataSource rdsPrint = new ReportDataSource("AnecdoteCaseP", sdsAnecdoteCaseP_99);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\AnecdoteCaseP.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.Refresh();
            plReport.Visible = true;
            plShowData.Visible = false;
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vSelectStr = "";
            switch (rbSelectPrintMode.SelectedValue)
            {
                case "11":
                    vSelectStr = CaseDataBind_1(false);
                    PrintReport_1();
                    break;
                case "12":
                    vSelectStr = CaseDataBind_1(true);
                    PrintReport_1();
                    break;
                case "21":
                    vSelectStr = CaseDataBind_2(false);
                    PrintReport_2();
                    break;
                case "22":
                    vSelectStr = CaseDataBind_2(true);
                    PrintReport_2();
                    break;
                case "99":
                    //統計報表
                    CaseDataBind_99();
                    PrintReport99();
                    break;
            }

            string vRecordModeStr = "肇事案件統計表_" + rbSelectPrintMode.SelectedItem.Text.Trim();
            string vRecordDateStr = ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() != "")) ? eBuildDate_Start_Search.Text.Trim() + "~" + eBuildDate_End_Search.Text.Trim() :
                                    ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() == "")) ? eBuildDate_Start_Search.Text.Trim() :
                                    ((eBuildDate_Start_Search.Text.Trim() == "") && (eBuildDate_End_Search.Text.Trim() != "")) ? eBuildDate_End_Search.Text.Trim() : "不指定日期";
            string vRecordDepNoStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_Start_Search.Text.Trim() + "~" + eDepNo_End_Search.Text.Trim() :
                                     ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? eDepNo_Start_Search.Text.Trim() :
                                     ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_End_Search.Text.Trim() : "全部";
            string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
            string vRecordNote = "列印資料_" + vRecordModeStr + Environment.NewLine +
                                 "AnecdoteCaseP.aspx" + Environment.NewLine +
                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                 "駕駛員：" + vRecordDriverStr + Environment.NewLine +
                                 "發生日期：" + vRecordDateStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
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
                    vHeaderText = (drExcel.GetName(i) == "InsuranceState") ? "出險" :
                                  (drExcel.GetName(i) == "DepNo") ? "部門編號" :
                                  (drExcel.GetName(i) == "DepName") ? "部門" :
                                  (drExcel.GetName(i) == "Driver") ? "員工編號" :
                                  (drExcel.GetName(i) == "DriverName") ? "駕駛員姓名" :
                                  (drExcel.GetName(i) == "BuildDate") ? "發生日期" :
                                  (drExcel.GetName(i) == "RelCar_ID") ? "對方車號" :
                                  (drExcel.GetName(i) == "Relationship") ? "對方姓名" :
                                  (drExcel.GetName(i) == "DepCount") ? "站別小計" :
                                  (drExcel.GetName(i) == "ThirdInsurance") ? "第三責任險" :
                                  (drExcel.GetName(i) == "CompInsurance") ? "強制險" :
                                  (drExcel.GetName(i) == "PassengerInsu") ? "乘客險" :
                                  (drExcel.GetName(i) == "CaseOccurrence") ? "肇事經過" : drExcel.GetName(i);
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
                        if ((drExcel.GetName(i) == "DepCount") ||
                            (drExcel.GetName(i) == "InsuCount") ||
                            (drExcel.GetName(i) == "ThirdInsurance") ||
                            (drExcel.GetName(i) == "CompInsurance") ||
                            (drExcel.GetName(i) == "PassengerInsu"))
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                        }
                        else if ((drExcel.GetName(i) == "BuildDate") && (drExcel[i].ToString() != ""))
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

                        string vRecordModeStr = "肇事案件統計表_" + rbSelectPrintMode.SelectedItem.Text.Trim();
                        string vRecordDateStr = ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() != "")) ? eBuildDate_Start_Search.Text.Trim() + "~" + eBuildDate_End_Search.Text.Trim() :
                                                ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() == "")) ? eBuildDate_Start_Search.Text.Trim() :
                                                ((eBuildDate_Start_Search.Text.Trim() == "") && (eBuildDate_End_Search.Text.Trim() != "")) ? eBuildDate_End_Search.Text.Trim() : "不指定日期";
                        string vRecordDepNoStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_Start_Search.Text.Trim() + "~" + eDepNo_End_Search.Text.Trim() :
                                                 ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? eDepNo_Start_Search.Text.Trim() :
                                                 ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_End_Search.Text.Trim() : "全部";
                        string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                        string vRecordNote = "匯出檔案_" + vRecordModeStr + Environment.NewLine +
                                             "AnecdoteCaseP.aspx" + Environment.NewLine +
                                             "站別：" + vRecordDepNoStr + Environment.NewLine +
                                             "駕駛員：" + vRecordDriverStr + Environment.NewLine +
                                             "發生日期：" + vRecordDateStr;
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
                    Response.Write("alert('" + eMessage.Message + eMessage.ToString() + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelectStr = "";
            switch (rbSelectPrintMode.SelectedValue)
            {
                case "11":
                    vSelectStr = CaseDataBind_1(false);
                    break;
                case "12":
                    vSelectStr = CaseDataBind_1(true);
                    break;
                case "21":
                    vSelectStr = CaseDataBind_2(false);
                    break;
                case "22":
                    vSelectStr = CaseDataBind_2(true);
                    break;
                case "99":
                    //統計報表
                    vSelectStr = CaseDataBind_99();
                    break;
            }
            if (vSelectStr != "")
            {
                SaveExcel(rbSelectPrintMode.SelectedItem.Text.Trim(), vSelectStr);
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