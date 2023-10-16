using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class LinesStatistics : System.Web.UI.Page
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
                    string vCountDateURL = "InputDate.aspx?TextboxID=" + eCountDate_Start_Search.ClientID;
                    string vCountDateScript = "window.open('" + vCountDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCountDate_Start_Search.Attributes["onClick"] = vCountDateScript;

                    vCountDateURL = "InputDate.aspx?TextboxID=" + eCountDate_End_Search.ClientID;
                    vCountDateScript = "window.open('" + vCountDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCountDate_End_Search.Attributes["onClick"] = vCountDateScript;

                    vCountDateURL = "InputDate.aspx?TextboxID=" + eSourceDate.ClientID;
                    vCountDateScript = "window.open('" + vCountDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eSourceDate.Attributes["onClick"] = vCountDateScript;

                    if (!IsPostBack)
                    {
                        eCountDate_Start_Search.Text = PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "B");
                        eCountDate_End_Search.Text = PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "B");
                        eSourceDate.Text = DateTime.Today.AddDays(-1).ToString("yyyy/MM/dd");
                        plReport.Visible = false;
                    }
                    else
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

        private string GetSelStr_01()
        {
            DateTime vStartDate = DateTime.Parse(eCountDate_Start_Search.Text.Trim());
            DateTime vEndDate = DateTime.Parse(eCountDate_End_Search.Text.Trim());
            string vShiftMode = rbListType.SelectedValue;
            string vMinRate = (Double.Parse(eMinQualifiedRate.Text.Trim()) / 100).ToString();
            string vShiftDate_S = (eCountDate_Start_Search.Text.Trim() != "") ? vStartDate.Year.ToString() + "/" + vStartDate.ToString("MM/dd") : "";
            string vShiftDate_E = (eCountDate_Start_Search.Text.Trim() != "") ? vEndDate.Year.ToString() + "/" + vEndDate.ToString("MM/dd") : "";
            string vWStr_ShiftDate = ((vShiftDate_S != "") && (vShiftDate_E != "")) ?
                                     "           and ShiftDate between '" + vShiftDate_S + "' and '" + vShiftDate_E + "' " + Environment.NewLine :
                                     ((vShiftDate_S != "") && (vShiftDate_E == "")) ?
                                     "           and ShiftDate = '" + vShiftDate_S + "' " + Environment.NewLine :
                                     ((vShiftDate_S == "") && (vShiftDate_E != "")) ?
                                     "           and ShiftDate = '" + vShiftDate_E + "' " + Environment.NewLine : "";
            string vReturnStr = "select z.LinesNo, z.LinesNote, " + Environment.NewLine +
                                "       sum([01]) [D01], sum([02]) [D02], sum([03]) [D03], sum([04]) [D04], sum([05]) [D05], " + Environment.NewLine +
                                "       sum([06]) [D06], sum([07]) [D07], sum([08]) [D08], sum([09]) [D09], sum([10]) [D10], " + Environment.NewLine +
                                "       sum([11]) [D11], sum([12]) [D12], sum([13]) [D13], sum([14]) [D14], sum([15]) [D15], " + Environment.NewLine +
                                "       sum([16]) [D16], sum([17]) [D17], sum([18]) [D18], sum([19]) [D19], sum([20]) [D20], " + Environment.NewLine +
                                "       sum([21]) [D21], sum([22]) [D22], sum([23]) [D23], sum([24]) [D24], sum([25]) [D25], " + Environment.NewLine +
                                "       sum([26]) [D26], sum([27]) [D27], sum([28]) [D28], sum([29]) [D29], sum([30]) [D30], sum([31]) [D31], " + Environment.NewLine +
                                "       sum(z.Unqualified) TotalDays " + Environment.NewLine +
                                "  from (" + Environment.NewLine +
                                "        select LinesNo, LinesNote, " + Environment.NewLine +
                                "               case when day(ShiftDate) = 1 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [01], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 2 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [02], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 3 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [03], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 4 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [04], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 5 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [05], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 6 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [06], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 7 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [07], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 8 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [08], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 9 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [09], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 10 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [10], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 11 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [11], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 12 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [12], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 13 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [13], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 14 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [14], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 15 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [15], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 16 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [16], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 17 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [17], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 18 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [18], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 19 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [19], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 20 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [20], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 21 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [21], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 22 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [22], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 23 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [23], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 24 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [24], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 25 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [25], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 26 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [26], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 27 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [27], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 28 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [28], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 29 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [29], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 30 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [30], " + Environment.NewLine +
                                "               case when day(ShiftDate) = 31 and (case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end as [31], " + Environment.NewLine +
                                "               case when(case when ApprovedShift = 0 then 0 else cast(QualifiedShift as float) / cast(ApprovedShift as float) end) < " + vMinRate + " then 1 else 0 end Unqualified " + Environment.NewLine +
                                "          from LinesStatistics " + Environment.NewLine +
                                "         where ShiftMode = '" + vShiftMode + "' " + Environment.NewLine +
                                vWStr_ShiftDate +
                                "       ) z " + Environment.NewLine +
                                " group by z.LinesNo, z.LinesNote " + Environment.NewLine +
                                " order by z.LinesNo";
            return vReturnStr;
        }

        private string GetSelStr_02()
        {
            DateTime vStartDate = DateTime.Parse(eCountDate_Start_Search.Text.Trim());
            DateTime vEndDate = DateTime.Parse(eCountDate_End_Search.Text.Trim());
            string vShiftMode = rbListType.SelectedValue;
            string vShiftDate_S = (eCountDate_Start_Search.Text.Trim() != "") ? vStartDate.Year.ToString() + "/" + vStartDate.ToString("MM/dd") : "";
            string vShiftDate_E = (eCountDate_Start_Search.Text.Trim() != "") ? vEndDate.Year.ToString() + "/" + vEndDate.ToString("MM/dd") : "";
            string vWStr_ShiftDate = ((vShiftDate_S != "") && (vShiftDate_E != "")) ?
                                     "   and ShiftDate between '" + vShiftDate_S + "' and '" + vShiftDate_E + "' " + Environment.NewLine :
                                     ((vShiftDate_S != "") && (vShiftDate_E == "")) ?
                                     "   and ShiftDate = '" + vShiftDate_S + "' " + Environment.NewLine :
                                     ((vShiftDate_S == "") && (vShiftDate_E != "")) ?
                                     "   and ShiftDate = '" + vShiftDate_E + "' " + Environment.NewLine : "";
            string vReturnStr = "select LinesNo, LinesNote, " + Environment.NewLine +
                                "       sum(ApprovedShift) as ApprovedShift, " + Environment.NewLine +
                                "       sum(QualifiedShift) QualifiedShift, " + Environment.NewLine +
                                "       case when sum(ApprovedShift) = 0 then 0 " + Environment.NewLine +
                                "            else sum(cast(QualifiedShift as float)) / sum(cast(ApprovedShift as float)) end QRate " + Environment.NewLine +
                                "  from LinesStatistics " + Environment.NewLine +
                                " where ShiftMode = '" + vShiftMode + "' " + Environment.NewLine +
                                vWStr_ShiftDate +
                                " group by LinesNo, LinesNote " + Environment.NewLine +
                                " order by LinesNo ";
            return vReturnStr;
        }

        private string GetSelStr_03()
        {
            DateTime vStartDate = DateTime.Parse(eCountDate_Start_Search.Text.Trim());
            DateTime vEndDate = DateTime.Parse(eCountDate_End_Search.Text.Trim());
            string vShiftMode = rbListType.SelectedValue;
            string vShiftDate_S = (eCountDate_Start_Search.Text.Trim() != "") ? vStartDate.Year.ToString() + "/" + vStartDate.ToString("MM/dd") : "";
            string vShiftDate_E = (eCountDate_Start_Search.Text.Trim() != "") ? vEndDate.Year.ToString() + "/" + vEndDate.ToString("MM/dd") : "";
            string vWStr_ShiftDate = ((vShiftDate_S != "") && (vShiftDate_E != "")) ?
                                     "   and ShiftDate between '" + vShiftDate_S + "' and '" + vShiftDate_E + "' " + Environment.NewLine :
                                     ((vShiftDate_S != "") && (vShiftDate_E == "")) ?
                                     "   and ShiftDate = '" + vShiftDate_S + "' " + Environment.NewLine :
                                     ((vShiftDate_S == "") && (vShiftDate_E != "")) ?
                                     "   and ShiftDate = '" + vShiftDate_E + "' " + Environment.NewLine : "";
            string vReturnStr = "select row_number() over (order by LinesNo) as RowsNo, isnull(Station, '') Station, LinesNo, LinesNote, " + Environment.NewLine +
                                "       case when sum(ApprovedShift) = 0 then 0 " + Environment.NewLine +
                                "            else cast(sum(QualifiedShift) as float) / cast(sum(ApprovedShift) as float) end QualifiedRate, " + Environment.NewLine +
                                "       case when sum(IdentificationShift) = 0 then 0 " + Environment.NewLine +
                                "            else cast(sum(InsertIDCardShift + DriverInputShift) as float) / cast(sum(IdentificationShift) as float) end IdentificationRate, " + Environment.NewLine +
                                "       case when sum(OnTimeTotalShift) = 0 then 0 " + Environment.NewLine +
                                "            else cast(sum(OntimeShift) as float) / cast(sum(OnTimeTotalShift) as float) end OnTimeRate " + Environment.NewLine +
                                "  from LinesStatistics " + Environment.NewLine +
                                " where ShiftMode = '03' " + Environment.NewLine +
                                vWStr_ShiftDate +
                                " group by Station, LinesNo, LinesNote " + Environment.NewLine +
                                " order by LinesNo";
            return vReturnStr;
        }

        private void Preview_01()
        {
            string vSelectStr = GetSelStr_01();
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

                ReportDataSource rdsPrint = new ReportDataSource("CityAreaBus", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\CityAreaBusP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "市區公車路線"));
                rvPrint.LocalReport.Refresh();
                plReport.Visible = true;
                plSearch.Visible = false;

                string vReportTypeStr = rbListType.SelectedItem.Text.Trim();
                string vCountDateStr = ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() != "")) ? "自 " + eCountDate_Start_Search.Text.Trim() + " 起至 " + eCountDate_End_Search.Text.Trim() + " 止" :
                                       ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() == "")) ? eCountDate_Start_Search.Text.Trim() :
                                       ((eCountDate_Start_Search.Text.Trim() == "") && (eCountDate_End_Search.Text.Trim() != "")) ? eCountDate_End_Search.Text.Trim() : "";
                string vRecordNote = "預覽報表_各路線合格率統計" + Environment.NewLine +
                                     "LinesStatistics.aspx" + Environment.NewLine +
                                     "過濾類別：" + vReportTypeStr + Environment.NewLine +
                                     "統計日期：" + vCountDateStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        private void Preview_02()
        {
            string vSelectStr = GetSelStr_02();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                DateTime vStartDate = DateTime.Parse(eCountDate_Start_Search.Text.Trim());
                DateTime vEndDate = DateTime.Parse(eCountDate_End_Search.Text.Trim());
                string vReportName = "免費公車路線" + vStartDate.Year.ToString() + vStartDate.ToString("MMdd") + "-" + vEndDate.ToString("MMdd");
                string vDateRange = vStartDate.Year.ToString() + "-" + vStartDate.ToString("MM-dd") + " ~ " + vEndDate.Year.ToString() + "-" + vEndDate.ToString("MM-dd");
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("FreeBus", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\FreeBusP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("DateRange", vDateRange));
                rvPrint.LocalReport.Refresh();
                plReport.Visible = true;
                plSearch.Visible = false;

                string vReportTypeStr = rbListType.SelectedItem.Text.Trim();
                string vCountDateStr = ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() != "")) ? "自 " + eCountDate_Start_Search.Text.Trim() + " 起至 " + eCountDate_End_Search.Text.Trim() + " 止" :
                                       ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() == "")) ? eCountDate_Start_Search.Text.Trim() :
                                       ((eCountDate_Start_Search.Text.Trim() == "") && (eCountDate_End_Search.Text.Trim() != "")) ? eCountDate_End_Search.Text.Trim() : "";
                string vRecordNote = "預覽報表_各路線合格率統計" + Environment.NewLine +
                                     "LinesStatistics.aspx" + Environment.NewLine +
                                     "過濾類別：" + vReportTypeStr + Environment.NewLine +
                                     "統計日期：" + vCountDateStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        private void Preview_03()
        {
            string vSelectStr = GetSelStr_03();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vStartDate = DateTime.Parse(eCountDate_Start_Search.Text.Trim());
            DateTime vEndDate = DateTime.Parse(eCountDate_End_Search.Text.Trim());
            string vShiftDate_S = (eCountDate_Start_Search.Text.Trim() != "") ? vStartDate.Year.ToString() + "/" + vStartDate.ToString("MM/dd") : "";
            string vShiftDate_E = (eCountDate_Start_Search.Text.Trim() != "") ? vEndDate.Year.ToString() + "/" + vEndDate.ToString("MM/dd") : "";
            string vWStr_ShiftDate = ((vShiftDate_S != "") && (vShiftDate_E != "")) ?
                                     "   and ShiftDate between '" + vShiftDate_S + "' and '" + vShiftDate_E + "' " + Environment.NewLine :
                                     ((vShiftDate_S != "") && (vShiftDate_E == "")) ?
                                     "   and ShiftDate = '" + vShiftDate_S + "' " + Environment.NewLine :
                                     ((vShiftDate_S == "") && (vShiftDate_E != "")) ?
                                     "   and ShiftDate = '" + vShiftDate_E + "' " + Environment.NewLine : "";
            string vReportName = "桃園客運動態資訊系統國道公路合格率一覽表";
            string vGetValueStr = "select case when sum(ApprovedShift) = 0 then 0 " + Environment.NewLine +
                                  "            else cast(sum(QualifiedShift) as float) / cast(sum(ApprovedShift) as float) end RateValue " + Environment.NewLine +
                                  "  from LinesStatistics " + Environment.NewLine +
                                  " where ShiftMode = '03' " + Environment.NewLine +
                                  vWStr_ShiftDate;
            double vTotalQualifiedRate = double.Parse(PF.GetValue(vConnStr, vGetValueStr, "RateValue"));
            vGetValueStr = "select case when sum(IdentificationShift) = 0 then 0 " + Environment.NewLine +
                           "            else cast(sum(InsertIDCardShift + DriverInputShift) as float) / cast(sum(IdentificationShift) as float) end RateValue " + Environment.NewLine +
                           "  from LinesStatistics " + Environment.NewLine +
                           " where ShiftMode = '03' " + Environment.NewLine +
                           vWStr_ShiftDate;
            double vTotalIdenRate = double.Parse(PF.GetValue(vConnStr, vGetValueStr, "RateValue"));
            vGetValueStr = "select case when sum(OnTimeTotalShift) = 0 then 0 " + Environment.NewLine +
                           "            else cast(sum(OntimeShift) as float) / cast(sum(OnTimeTotalShift) as float) end RateValue " + Environment.NewLine +
                           "  from LinesStatistics " + Environment.NewLine +
                           " where ShiftMode = '03' " + Environment.NewLine +
                           vWStr_ShiftDate;
            double vTotalOntimeRate = double.Parse(PF.GetValue(vConnStr, vGetValueStr, "RateValue"));
            string vTitleStr = (vStartDate.Year - 1911).ToString() + "/" + vStartDate.ToString("MM/dd") + "---" +
                               (vEndDate.Year - 1911).ToString() + "/" + vEndDate.ToString("MM/dd") +
                               "路線合格率" + vTotalQualifiedRate.ToString("P", CultureInfo.InvariantCulture) +
                               "身分識別合格率" + vTotalIdenRate.ToString("P", CultureInfo.InvariantCulture) +
                               "預估到站準確率" + vTotalOntimeRate.ToString("P", CultureInfo.InvariantCulture);
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("HighwayBus", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\HighwayBusP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("TitleName", vTitleStr));
                rvPrint.LocalReport.Refresh();
                plReport.Visible = true;
                plSearch.Visible = false;

                string vReportTypeStr = rbListType.SelectedItem.Text.Trim();
                string vCountDateStr = ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() != "")) ? "自 " + eCountDate_Start_Search.Text.Trim() + " 起至 " + eCountDate_End_Search.Text.Trim() + " 止" :
                                       ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() == "")) ? eCountDate_Start_Search.Text.Trim() :
                                       ((eCountDate_Start_Search.Text.Trim() == "") && (eCountDate_End_Search.Text.Trim() != "")) ? eCountDate_End_Search.Text.Trim() : "";
                string vRecordNote = "預覽報表_各路線合格率統計" + Environment.NewLine +
                                     "LinesStatistics.aspx" + Environment.NewLine +
                                     "過濾類別：" + vReportTypeStr + Environment.NewLine +
                                     "統計日期：" + vCountDateStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        private void ExportExcel_01() //市區公車路線
        {
            DateTime vStartDate = DateTime.Parse(eCountDate_Start_Search.Text.Trim());
            DateTime vEndDate = DateTime.Parse(eCountDate_End_Search.Text.Trim());
            string vSelStr = GetSelStr_01();
            string vTitleStr = "市區公車路線 ";
            string vFileName = "市區公車行車趟次統計表（依路線）[" + vStartDate.Year.ToString() + "-" + vStartDate.ToString("MM-dd") +
                               "~" + vEndDate.Year.ToString() + "-" + vEndDate.ToString("MM-dd") + "]";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
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

                    HSSFCellStyle csData_Percentage = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData_Percentage.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csData_Percentage.DataFormat = format.GetFormat("0.00%");

                    string vHeaderText = "";
                    int vLinesNo = 0;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vTitleStr);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(vTitleStr);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 1));
                    for (int i = 2; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper().Trim() == "TOTALDAYS") ? "合計天數" : drExcel.GetName(i).ToUpper().Trim();
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                    }

                    while (drExcel.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            vHeaderText = drExcel.GetName(i).ToUpper().Trim();
                            if ((vHeaderText == "LINESNO") || (vHeaderText == "LINESNOTE"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (vHeaderText == "TOTALDAYS")
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                if (drExcel[i].ToString() == "0")
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("");
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                }
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
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
                            string vReportTypeStr = rbListType.SelectedItem.Text.Trim();
                            string vCountDateStr = ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() != "")) ? "自 " + eCountDate_Start_Search.Text.Trim() + " 起至 " + eCountDate_End_Search.Text.Trim() + " 止" :
                                                   ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() == "")) ? eCountDate_Start_Search.Text.Trim() :
                                                   ((eCountDate_Start_Search.Text.Trim() == "") && (eCountDate_End_Search.Text.Trim() != "")) ? eCountDate_End_Search.Text.Trim() : "";
                            string vRecordNote = "匯出報表_各路線合格率統計" + Environment.NewLine +
                                                 "LinesStatistics.aspx" + Environment.NewLine +
                                                 "過濾類別：" + vReportTypeStr + Environment.NewLine +
                                                 "統計日期：" + vCountDateStr;
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

        private void ExportExcel_02() //免費巴士路線
        {
            DateTime vStartDate = DateTime.Parse(eCountDate_Start_Search.Text.Trim());
            DateTime vEndDate = DateTime.Parse(eCountDate_End_Search.Text.Trim());
            string vSelStr = GetSelStr_02();
            string vTitleStr = "免費巴士路線 " + (vStartDate.Year - 1911).ToString() + vStartDate.ToString("MMdd") + "--" + vEndDate.ToString("MMdd");
            string vFileName = "免費巴士行車趟次統計表（依路線）[" + vStartDate.Year.ToString() + "-" + vStartDate.ToString("MM-dd") +
                               "~" + vEndDate.Year.ToString() + "-" + vEndDate.ToString("MM-dd") + "]";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
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

                    HSSFCellStyle csData_Percentage = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData_Percentage.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csData_Percentage.DataFormat = format.GetFormat("0.00%");

                    string vHeaderText = "";
                    int vLinesNo = 0;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vTitleStr);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(vTitleStr);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 1));
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("核定班次數");
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("合格班次數");
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("合格比例");

                    while (drExcel.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            vHeaderText = drExcel.GetName(i).ToUpper();
                            if ((vHeaderText == "LINESNO") || (vHeaderText == "LINESNOTE"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if ((vHeaderText == "APPROVEDSHIFT") || (vHeaderText == "QUALIFIEDSHIFT"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else if (vHeaderText == "QRATE")
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Percentage;
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
                            string vReportTypeStr = rbListType.SelectedItem.Text.Trim();
                            string vCountDateStr = ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() != "")) ? "自 " + eCountDate_Start_Search.Text.Trim() + " 起至 " + eCountDate_End_Search.Text.Trim() + " 止" :
                                                   ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() == "")) ? eCountDate_Start_Search.Text.Trim() :
                                                   ((eCountDate_Start_Search.Text.Trim() == "") && (eCountDate_End_Search.Text.Trim() != "")) ? eCountDate_End_Search.Text.Trim() : "";
                            string vRecordNote = "匯出報表_各路線合格率統計" + Environment.NewLine +
                                                 "LinesStatistics.aspx" + Environment.NewLine +
                                                 "過濾類別：" + vReportTypeStr + Environment.NewLine +
                                                 "統計日期：" + vCountDateStr;
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

        private void ExportExcel_03() //國道公路動態系統
        {
            DateTime vStartDate = DateTime.Parse(eCountDate_Start_Search.Text.Trim());
            DateTime vEndDate = DateTime.Parse(eCountDate_End_Search.Text.Trim());
            string vShiftDate_S = (eCountDate_Start_Search.Text.Trim() != "") ? vStartDate.Year.ToString() + "/" + vStartDate.ToString("MM/dd") : "";
            string vShiftDate_E = (eCountDate_Start_Search.Text.Trim() != "") ? vEndDate.Year.ToString() + "/" + vEndDate.ToString("MM/dd") : "";
            string vWStr_ShiftDate = ((vShiftDate_S != "") && (vShiftDate_E != "")) ?
                                     "   and ShiftDate between '" + vShiftDate_S + "' and '" + vShiftDate_E + "' " + Environment.NewLine :
                                     ((vShiftDate_S != "") && (vShiftDate_E == "")) ?
                                     "   and ShiftDate = '" + vShiftDate_S + "' " + Environment.NewLine :
                                     ((vShiftDate_S == "") && (vShiftDate_E != "")) ?
                                     "   and ShiftDate = '" + vShiftDate_E + "' " + Environment.NewLine : "";
            string vSelStr = GetSelStr_03();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vReportName = "桃園客運動態資訊系統國道公路合格率一覽表";
            string vGetValueStr = "select case when sum(ApprovedShift) = 0 then 0 " + Environment.NewLine +
                                  "            else cast(sum(QualifiedShift) as float) / cast(sum(ApprovedShift) as float) end RateValue " + Environment.NewLine +
                                  "  from LinesStatistics " + Environment.NewLine +
                                  " where ShiftMode = '03' " + Environment.NewLine +
                                  vWStr_ShiftDate;
            double vTotalQualifiedRate = double.Parse(PF.GetValue(vConnStr, vGetValueStr, "RateValue"));
            vGetValueStr = "select case when sum(IdentificationShift) = 0 then 0 " + Environment.NewLine +
                           "            else cast(sum(InsertIDCardShift + DriverInputShift) as float) / cast(sum(IdentificationShift) as float) end RateValue " + Environment.NewLine +
                           "  from LinesStatistics " + Environment.NewLine +
                           " where ShiftMode = '03' " + Environment.NewLine +
                           vWStr_ShiftDate;
            double vTotalIdenRate = double.Parse(PF.GetValue(vConnStr, vGetValueStr, "RateValue"));
            vGetValueStr = "select case when sum(OnTimeTotalShift) = 0 then 0 " + Environment.NewLine +
                           "            else cast(sum(OntimeShift) as float) / cast(sum(OnTimeTotalShift) as float) end RateValue " + Environment.NewLine +
                           "  from LinesStatistics " + Environment.NewLine +
                           " where ShiftMode = '03' " + Environment.NewLine +
                           vWStr_ShiftDate;
            double vTotalOntimeRate = double.Parse(PF.GetValue(vConnStr, vGetValueStr, "RateValue"));
            string vTitleStr = (vStartDate.Year - 1911).ToString() + "/" + vStartDate.ToString("MM/dd") + "---" +
                               (vEndDate.Year - 1911).ToString() + "/" + vEndDate.ToString("MM/dd") +
                               "路線合格率" + vTotalQualifiedRate.ToString("P", CultureInfo.InvariantCulture) +
                               "身分識別合格率" + vTotalIdenRate.ToString("P", CultureInfo.InvariantCulture) +
                               "預估到站準確率" + vTotalOntimeRate.ToString("P", CultureInfo.InvariantCulture);
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
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

                    HSSFCellStyle csData_Percentage = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData_Percentage.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csData_Percentage.DataFormat = format.GetFormat("0.00%");

                    string vHeaderText = "";
                    int vLinesNo = 0;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vReportName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(vReportName);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 6));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    vLinesNo++;
                    //寫入標題列2
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(vTitleStr);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 6));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    vLinesNo++;
                    //寫入標題列3
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "RowsNo") ? "序號" :
                                      (drExcel.GetName(i) == "LinesNo") ? "路線編碼" :
                                      (drExcel.GetName(i) == "Station") ? "派車單位" :
                                      (drExcel.GetName(i) == "LinesNote") ? "路線別  起--(經由)---迄" :
                                      (drExcel.GetName(i) == "QualifiedRate") ? "路線合格率" :
                                      (drExcel.GetName(i) == "IdentificationRate") ? "身分識別率" :
                                      (drExcel.GetName(i) == "OnTimeRate") ? "到站準點率" : "";
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
                            vHeaderText = drExcel.GetName(i).ToUpper();
                            if ((vHeaderText == "LINESNO") || (vHeaderText == "LINESNOTE") || (vHeaderText == "STATION"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (vHeaderText == "ROWSNO")
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else if ((vHeaderText == "QUALIFIEDRATE") || (vHeaderText == "IDENTIFICATIONRATE") || (vHeaderText == "ONTIMERATE"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Percentage;
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
                            string vReportTypeStr = rbListType.SelectedItem.Text.Trim();
                            string vCountDateStr = ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() != "")) ? "自 " + eCountDate_Start_Search.Text.Trim() + " 起至 " + eCountDate_End_Search.Text.Trim() + " 止" :
                                                   ((eCountDate_Start_Search.Text.Trim() != "") && (eCountDate_End_Search.Text.Trim() == "")) ? eCountDate_Start_Search.Text.Trim() :
                                                   ((eCountDate_Start_Search.Text.Trim() == "") && (eCountDate_End_Search.Text.Trim() != "")) ? eCountDate_End_Search.Text.Trim() : "";
                            string vRecordNote = "匯出報表_各路線合格率統計" + Environment.NewLine +
                                                 "LinesStatistics.aspx" + Environment.NewLine +
                                                 "過濾類別：" + vReportTypeStr + Environment.NewLine +
                                                 "統計日期：" + vCountDateStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vReportName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vReportName + ".xls"));
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

        private void ImportExcelData_01(string fUploadFileName)
        {
            int vPosIndex = 0;
            string vIndexNo = "";
            string vCompany = "";
            string vLinesNo = "";
            string vLinesNote = "";
            int vApprovedShift = 0; //核定班次數
            int vQualifiedShift = 0; //合格班次數
            int vUnqualifiedShift = 0; //不合格班次數
            double vQualifiedRate = 0.0; //合格比率
            string vBuMan = "";
            string vShiftDate_Str = "";
            string vBuDate_Str = "";
            string vShiftMode = "";

            DateTime vShiftDate;
            DateTime vBuDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vExtName = Path.GetExtension(fuExcel.FileName);
            fuExcel.SaveAs(fUploadFileName);

            if (vExtName == ".xls")
            {
                HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                {
                    HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                    if ((vRowTemp_H != null) && (vRowTemp_H.GetCell(0).StringCellValue == "桃園客運"))
                    {
                        vCompany = vRowTemp_H.GetCell(0).StringCellValue;
                        vLinesNote = vRowTemp_H.GetCell(1).StringCellValue;
                        vPosIndex = vLinesNote.IndexOf("]");
                        vLinesNo = vLinesNote.Substring(1, vPosIndex - 1);
                        vLinesNote = vLinesNote.Substring(vPosIndex + 1).Replace(" ", "");
                        vShiftDate = DateTime.Parse(vRowTemp_H.GetCell(2).StringCellValue);
                        vApprovedShift = (Int32)vRowTemp_H.GetCell(3).NumericCellValue;
                        vQualifiedShift = (Int32)vRowTemp_H.GetCell(4).NumericCellValue;
                        vQualifiedRate = vRowTemp_H.GetCell(5).NumericCellValue;
                        vUnqualifiedShift = (Int32)vRowTemp_H.GetCell(6).NumericCellValue;
                        vBuMan = vLoginID;
                        vBuDate = DateTime.Today;
                        vShiftMode = (vLinesNo.Substring(0, 1).ToUpper() == "L") ? "02" : "01";
                        vIndexNo = "0000000000" + vLinesNo;
                        vIndexNo = (vShiftDate.Year > 3822) ?
                                      vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                      (vShiftDate.Year - 1911).ToString("D4") +
                                      vShiftDate.Month.ToString("D2") +
                                      vShiftDate.Day.ToString("D2") :
                                   (vShiftDate.Year < 1911) ?
                                      vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                      (vShiftDate.Year + 1911).ToString("D4") +
                                      vShiftDate.Month.ToString("D2") +
                                      vShiftDate.Day.ToString("D2") :
                                   vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                      vShiftDate.Year.ToString("D4") +
                                      vShiftDate.Month.ToString("D2") +
                                      vShiftDate.Day.ToString("D2");
                        vShiftDate_Str = (vShiftDate.Year > 3822) ? (vShiftDate.Year - 1911).ToString("D4") + "/" + vShiftDate.ToString("MM/dd") :
                                         (vShiftDate.Year < 1911) ? (vShiftDate.Year + 1911).ToString("D4") + "/" + vShiftDate.ToString("MM/dd") :
                                         vShiftDate.Year.ToString("D4") + "/" + vShiftDate.ToString("MM/dd");
                        vBuDate_Str = (vBuDate.Year > 3822) ? (vBuDate.Year - 1911).ToString("D4") + "/" + vBuDate.ToString("MM/dd") :
                                      (vBuDate.Year < 1911) ? (vBuDate.Year + 1911).ToString("D4") + "/" + vBuDate.ToString("MM/dd") :
                                      vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd");

                        sdsLinesStatistics.InsertParameters.Clear();
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("LinesNote", DbType.String, vLinesNote));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("ShiftDate", DbType.Date, vShiftDate_Str));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("ApprovedShift", DbType.Int32, vApprovedShift.ToString()));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("QualifiedShift", DbType.Int32, vQualifiedShift.ToString()));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("UnqualifiedShift", DbType.Int32, vUnqualifiedShift.ToString()));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("QualifiedRate", DbType.Double, vQualifiedRate.ToString()));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("BuMan", DbType.String, vBuMan));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDate_Str));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("Station", DbType.String, ""));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("IdentificationRate", DbType.Double, ""));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("OntimeRate", DbType.Double, ""));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("Remark", DbType.String, ""));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("ShiftMode", DbType.String, vShiftMode));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("IdentificationShift", DbType.Int32, "0"));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("InsertIDCardShift", DbType.Int32, "0"));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("DriverInputShift", DbType.Int32, "0"));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("OnTimeTotalShift", DbType.Int32, "0"));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("OntimeShift", DbType.Int32, "0"));
                        sdsLinesStatistics.Insert();

                    }
                }
            }
            else if (vExtName == ".xlsx")
            {
                XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                {
                    XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                    if ((vRowTemp_X != null) && (vRowTemp_X.GetCell(0).StringCellValue == "桃園客運"))
                    {
                        vCompany = vRowTemp_X.GetCell(0).StringCellValue;
                        vLinesNote = vRowTemp_X.GetCell(1).StringCellValue;
                        vPosIndex = vLinesNote.IndexOf("]");
                        vLinesNo = vLinesNote.Substring(1, vPosIndex - 1);
                        vLinesNote = vLinesNote.Substring(vPosIndex + 1).Replace(" ", "");
                        vShiftDate = DateTime.Parse(vRowTemp_X.GetCell(2).StringCellValue);
                        vApprovedShift = (Int32)vRowTemp_X.GetCell(3).NumericCellValue;
                        vQualifiedShift = (Int32)vRowTemp_X.GetCell(4).NumericCellValue;
                        vQualifiedRate = vRowTemp_X.GetCell(5).NumericCellValue;
                        vUnqualifiedShift = (Int32)vRowTemp_X.GetCell(6).NumericCellValue;
                        vBuMan = vLoginID;
                        vBuDate = DateTime.Today;
                        vIndexNo = "0000000000" + vLinesNo;
                        vShiftMode = (vLinesNo.Substring(0, 1).ToUpper() == "L") ? "02" : "01";
                        vIndexNo = (vShiftDate.Year > 3822) ?
                                      vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                      (vShiftDate.Year - 1911).ToString("D4") +
                                      vShiftDate.Month.ToString("D2") +
                                      vShiftDate.Day.ToString("D2") :
                                   (vShiftDate.Year < 1911) ?
                                      vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                      (vShiftDate.Year + 1911).ToString("D4") +
                                      vShiftDate.Month.ToString("D2") +
                                      vShiftDate.Day.ToString("D2") :
                                   vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                      vShiftDate.Year.ToString("D4") +
                                      vShiftDate.Month.ToString("D2") +
                                      vShiftDate.Day.ToString("D2");
                        vShiftDate_Str = (vShiftDate.Year > 3822) ? (vShiftDate.Year - 1911).ToString("D4") + "/" + vShiftDate.ToString("MM/dd") :
                                         (vShiftDate.Year < 1911) ? (vShiftDate.Year + 1911).ToString("D4") + "/" + vShiftDate.ToString("MM/dd") :
                                         vShiftDate.Year.ToString("D4") + "/" + vShiftDate.ToString("MM/dd");
                        vBuDate_Str = (vBuDate.Year > 3822) ? (vBuDate.Year - 1911).ToString("D4") + "/" + vBuDate.ToString("MM/dd") :
                                      (vBuDate.Year < 1911) ? (vBuDate.Year + 1911).ToString("D4") + "/" + vBuDate.ToString("MM/dd") :
                                      vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd");

                        sdsLinesStatistics.InsertParameters.Clear();
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("LinesNote", DbType.String, vLinesNote));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("ShiftDate", DbType.Date, vShiftDate_Str));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("ApprovedShift", DbType.Int32, vApprovedShift.ToString()));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("QualifiedShift", DbType.Int32, vQualifiedShift.ToString()));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("UnqualifiedShift", DbType.Int32, vUnqualifiedShift.ToString()));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("QualifiedRate", DbType.Double, vQualifiedRate.ToString()));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("BuMan", DbType.String, vBuMan));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDate_Str));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("Station", DbType.String, ""));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("IdentificationRate", DbType.Double, ""));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("OntimeRate", DbType.Double, ""));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("Remark", DbType.String, ""));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("ShiftMode", DbType.String, vShiftMode));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("IdentificationShift", DbType.Int32, "0"));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("InsertIDCardShift", DbType.Int32, "0"));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("DriverInputShift", DbType.Int32, "0"));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("OnTimeTotalShift", DbType.Int32, "0"));
                        sdsLinesStatistics.InsertParameters.Add(new Parameter("OntimeShift", DbType.Int32, "0"));
                        sdsLinesStatistics.Insert();
                    }
                }
            }
        }

        private void ImportExcelData_02(string fUploadFileName)
        {
            string vInputType = ""; //因為輸入的EXCEL有三種格式，用這個變數來區分
            string vDataCountStr = "";
            string vUpdateStr = "";
            int vDataCount = 0;
            int vPosIndex = 0;
            string vIndexNo = "";
            string vCompany = "";
            string vStation = "";
            string vLinesNo = "";
            string vLinesNote = "";
            int vApprovedShift = 0; //核定班次數
            int vQualifiedShift = 0; //合格班次數
            int vUnqualifiedShift = 0; //不合格班次數
            double vQualifiedRate = 0.0; //合格比率
            int vIdentificationShift = 0; //身份識別總筆數
            int vInsertIDCardShift = 0; //插卡筆數
            int vDriverInputShift = 0; //駕駛員輸入筆數
            double vIdentificationRate = 0.0; //身份識別率
            int vOnTimeTotalShift = 0; //準確率總筆數
            int vOntimeShift = 0; //準確筆數
            double vOntimeRate = 0.0; //準確率
            string vBuMan = "";
            string vShiftDate_Str = "";
            string vBuDate_Str = "";
            string vShiftMode = "";

            DateTime vShiftDate;
            DateTime vBuDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vExtName = Path.GetExtension(fuExcel.FileName);
            fuExcel.SaveAs(fUploadFileName);

            if (vExtName == ".xls")
            {
                HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                {
                    HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                    if ((vRowTemp_H != null) && (i == 0))
                    {
                        vInputType = (vRowTemp_H.GetCell(0).StringCellValue.IndexOf("插入身份識別統計") >= 0) ? "身份識別" :
                                     (vRowTemp_H.GetCell(0).StringCellValue.IndexOf("行車班次統計") >= 0) ? "班次統計" :
                                     (vRowTemp_H.GetCell(0).StringCellValue.IndexOf("預估到站準確率") >= 0) ? "到站準點率" : "";
                    }
                    else if ((vRowTemp_H != null) && (vInputType != "") && (vRowTemp_H.GetCell(0).StringCellValue.Trim() != ""))
                    {
                        if (vRowTemp_H.GetCell(4).CellType == CellType.Numeric)
                        {
                            if (vInputType == "到站準點率")
                            {
                                vCompany = vRowTemp_H.GetCell(1).StringCellValue;
                                vLinesNote = vRowTemp_H.GetCell(2).StringCellValue;
                                vStation = "";
                                vApprovedShift = 0; //核定班次數
                                vQualifiedShift = 0; //合格班次數
                                vUnqualifiedShift = 0; //不合格班次數
                                vQualifiedRate = 0.0; //合格比率
                                vIdentificationShift = 0; //身份識別總筆數
                                vInsertIDCardShift = 0; //插卡筆數
                                vDriverInputShift = 0; //駕駛員輸入筆數
                                vIdentificationRate = 0.0; //身份識別率
                                vOnTimeTotalShift = (Int32)vRowTemp_H.GetCell(3).NumericCellValue; //準確率總筆數
                                vOntimeShift = (Int32)vRowTemp_H.GetCell(4).NumericCellValue; //準確筆數
                                vOntimeRate = vRowTemp_H.GetCell(5).NumericCellValue; //準確率
                            }
                            else if (vInputType == "身份識別")
                            {
                                vCompany = vRowTemp_H.GetCell(0).StringCellValue;
                                vStation = vRowTemp_H.GetCell(1).StringCellValue;
                                vLinesNote = vRowTemp_H.GetCell(3).StringCellValue;
                                vApprovedShift = 0; //核定班次數
                                vQualifiedShift = 0; //合格班次數
                                vUnqualifiedShift = 0; //不合格班次數
                                vQualifiedRate = 0.0; //合格比率
                                vIdentificationShift = (Int32)vRowTemp_H.GetCell(4).NumericCellValue; //身份識別總筆數
                                vInsertIDCardShift = (Int32)vRowTemp_H.GetCell(5).NumericCellValue; //插卡筆數
                                vDriverInputShift = (Int32)vRowTemp_H.GetCell(6).NumericCellValue; //駕駛員輸入筆數
                                vIdentificationRate = (vRowTemp_H.GetCell(7).NumericCellValue / 100.0); //身份識別率
                                vOnTimeTotalShift = 0; //準確率總筆數
                                vOntimeShift = 0; //準確筆數
                                vOntimeRate = 0.0; //準確率
                            }
                            else if (vInputType == "班次統計")
                            {
                                vCompany = vRowTemp_H.GetCell(0).StringCellValue;
                                vStation = vRowTemp_H.GetCell(1).StringCellValue;
                                vLinesNote = vRowTemp_H.GetCell(3).StringCellValue;
                                vApprovedShift = (Int32)vRowTemp_H.GetCell(4).NumericCellValue; //核定班次數
                                vQualifiedShift = (Int32)vRowTemp_H.GetCell(5).NumericCellValue; //合格班次數
                                vUnqualifiedShift = (Int32)vRowTemp_H.GetCell(7).NumericCellValue; //不合格班次數
                                vQualifiedRate = Double.Parse(vRowTemp_H.GetCell(6).StringCellValue.Replace("%", "").Trim()) / 100.0; //合格比率
                                vIdentificationShift = 0; //身份識別總筆數
                                vInsertIDCardShift = 0; //插卡筆數
                                vDriverInputShift = 0; //駕駛員輸入筆數
                                vIdentificationRate = 0.0; //身份識別率
                                vOnTimeTotalShift = 0; //準確率總筆數
                                vOntimeShift = 0; //準確筆數
                                vOntimeRate = 0.0; //準確率
                            }
                            vPosIndex = vLinesNote.IndexOf("】");
                            vLinesNo = vLinesNote.Substring(1, vPosIndex - 1);
                            vLinesNote = vLinesNote.Substring(vPosIndex + 1).Replace(" ", "");
                            vShiftDate = DateTime.Parse(eSourceDate.Text.Trim());
                            vBuMan = vLoginID;
                            vBuDate = DateTime.Today;
                            vShiftMode = "03";
                            vIndexNo = "0000000000" + vLinesNo;
                            vIndexNo = (vShiftDate.Year > 3822) ?
                                          vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                          (vShiftDate.Year - 1911).ToString("D4") +
                                          vShiftDate.Month.ToString("D2") +
                                          vShiftDate.Day.ToString("D2") :
                                       (vShiftDate.Year < 1911) ?
                                          vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                          (vShiftDate.Year + 1911).ToString("D4") +
                                          vShiftDate.Month.ToString("D2") +
                                          vShiftDate.Day.ToString("D2") :
                                       vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                          vShiftDate.Year.ToString("D4") +
                                          vShiftDate.Month.ToString("D2") +
                                          vShiftDate.Day.ToString("D2");
                            vShiftDate_Str = (vShiftDate.Year > 3822) ? (vShiftDate.Year - 1911).ToString("D4") + "/" + vShiftDate.ToString("MM/dd") :
                                             (vShiftDate.Year < 1911) ? (vShiftDate.Year + 1911).ToString("D4") + "/" + vShiftDate.ToString("MM/dd") :
                                             vShiftDate.Year.ToString("D4") + "/" + vShiftDate.ToString("MM/dd");
                            vBuDate_Str = (vBuDate.Year > 3822) ? (vBuDate.Year - 1911).ToString("D4") + "/" + vBuDate.ToString("MM/dd") :
                                          (vBuDate.Year < 1911) ? (vBuDate.Year + 1911).ToString("D4") + "/" + vBuDate.ToString("MM/dd") :
                                          vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd");

                            vDataCountStr = "select Count(IndexNo) RCount " + Environment.NewLine +
                                            "  from LinesStatistics " + Environment.NewLine +
                                            " where LinesNo = '" + vLinesNo + "' " + Environment.NewLine +
                                            "   and ShiftDate = '" + vShiftDate_Str + "' " + Environment.NewLine +
                                            "   and ShiftMode = '" + vShiftMode + "' ";
                            vDataCount = Int32.Parse(PF.GetValue(vConnStr, vDataCountStr, "RCount"));
                            if (vDataCount == 0)
                            {
                                sdsLinesStatistics.InsertParameters.Clear();
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("LinesNote", DbType.String, vLinesNote));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("ShiftDate", DbType.Date, vShiftDate_Str));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("ApprovedShift", DbType.Int32, vApprovedShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("QualifiedShift", DbType.Int32, vQualifiedShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("UnqualifiedShift", DbType.Int32, vUnqualifiedShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("QualifiedRate", DbType.Double, vQualifiedRate.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("BuMan", DbType.String, vBuMan));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDate_Str));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("Station", DbType.String, vStation));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("IdentificationShift", DbType.Int32, vIdentificationShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("InsertIDCardShift", DbType.Int32, vInsertIDCardShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("DriverInputShift", DbType.Int32, vDriverInputShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("IdentificationRate", DbType.Double, vIdentificationRate.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("OnTimeTotalShift", DbType.Int32, vOnTimeTotalShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("OntimeShift", DbType.Int32, vOntimeShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("OntimeRate", DbType.Double, vOntimeRate.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("Remark", DbType.String, ""));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("ShiftMode", DbType.String, vShiftMode));
                                sdsLinesStatistics.Insert();
                            }
                            else if (vDataCount > 0)
                            {
                                vUpdateStr = "update LinesStatistics " + Environment.NewLine +
                                             "   set Station = case when Station is null then '" + vStation + "' else Station end, " + Environment.NewLine +
                                             "       ApprovedShift = ApprovedShift + " + vApprovedShift + ", " + Environment.NewLine +
                                             "       QualifiedShift = QualifiedShift + " + vQualifiedShift + ", " + Environment.NewLine +
                                             "       UnqualifiedShift = UnqualifiedShift + " + vUnqualifiedShift + ", " + Environment.NewLine +
                                             "       QualifiedRate = case when cast(ApprovedShift + " + vApprovedShift + " as float) = 0 then 0 else cast(QualifiedShift + " + vQualifiedShift + " as float) / cast(ApprovedShift + " + vApprovedShift + " as float) end, " + Environment.NewLine +
                                             "       IdentificationShift = IdentificationShift + " + vIdentificationShift + ", " + Environment.NewLine +
                                             "       InsertIDCardShift = InsertIDCardShift + " + vInsertIDCardShift + ", " + Environment.NewLine +
                                             "       DriverInputShift = DriverInputShift + " + vDriverInputShift + ", " + Environment.NewLine +
                                             "       IdentificationRate = case when cast(IdentificationShift + " + vIdentificationShift + " as float) = 0 then 0 else cast(InsertIDCardShift + " + vInsertIDCardShift + " + DriverInputShift + " + vDriverInputShift + " as float) / cast(IdentificationShift + " + vIdentificationShift + " as float) end, " + Environment.NewLine +
                                             "       OnTimeTotalShift = OnTimeTotalShift + " + vOnTimeTotalShift + ", " + Environment.NewLine +
                                             "       OntimeShift = OntimeShift + " + vOntimeShift + ", " + Environment.NewLine +
                                             "       OntimeRate = case when cast(OnTimeTotalShift + " + vOnTimeTotalShift + " as float) = 0 then 0 else cast(OntimeShift + " + vOntimeShift + " as float) / cast(OnTimeTotalShift + " + vOnTimeTotalShift + " as float) end " + Environment.NewLine +
                                             " where LinesNo = '" + vLinesNo + "' " + Environment.NewLine +
                                             "   and ShiftDate = '" + vShiftDate_Str + "' " + Environment.NewLine +
                                             "   and ShiftMode = '" + vShiftMode + "' ";
                                PF.ExecSQL(vConnStr, vUpdateStr);
                            }
                        }
                    }
                }
            }
            else if (vExtName == ".xlsx")
            {
                XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                {
                    XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                    if ((vRowTemp_X != null) && (i == 0))
                    {
                        vInputType = (vRowTemp_X.GetCell(0).StringCellValue.IndexOf("插入身份識別統計") >= 0) ? "身份識別" :
                                     (vRowTemp_X.GetCell(0).StringCellValue.IndexOf("行車班次統計") >= 0) ? "班次統計" :
                                     (vRowTemp_X.GetCell(0).StringCellValue.IndexOf("預估到站準確率") >= 0) ? "到站準點率" : "";
                    }
                    else if ((vRowTemp_X != null) && (vInputType != "") && (vRowTemp_X.GetCell(0).StringCellValue.Trim() != ""))
                    {
                        if (vRowTemp_X.GetCell(4).CellType == CellType.Numeric)
                        {
                            if (vInputType == "到站準點率")
                            {
                                vCompany = vRowTemp_X.GetCell(1).StringCellValue;
                                vLinesNote = vRowTemp_X.GetCell(2).StringCellValue;
                                vStation = "";
                                vApprovedShift = 0; //核定班次數
                                vQualifiedShift = 0; //合格班次數
                                vUnqualifiedShift = 0; //不合格班次數
                                vQualifiedRate = 0.0; //合格比率
                                vIdentificationShift = 0; //身份識別總筆數
                                vInsertIDCardShift = 0; //插卡筆數
                                vDriverInputShift = 0; //駕駛員輸入筆數
                                vIdentificationRate = 0.0; //身份識別率
                                vOnTimeTotalShift = (Int32)vRowTemp_X.GetCell(3).NumericCellValue; //準確率總筆數
                                vOntimeShift = (Int32)vRowTemp_X.GetCell(4).NumericCellValue; //準確筆數
                                vOntimeRate = vRowTemp_X.GetCell(5).NumericCellValue; //準確率
                            }
                            else if (vInputType == "身份識別")
                            {
                                vCompany = vRowTemp_X.GetCell(0).StringCellValue;
                                vStation = vRowTemp_X.GetCell(1).StringCellValue;
                                vLinesNote = vRowTemp_X.GetCell(3).StringCellValue;
                                vApprovedShift = 0; //核定班次數
                                vQualifiedShift = 0; //合格班次數
                                vUnqualifiedShift = 0; //不合格班次數
                                vQualifiedRate = 0.0; //合格比率
                                vIdentificationShift = (Int32)vRowTemp_X.GetCell(4).NumericCellValue; //身份識別總筆數
                                vInsertIDCardShift = (Int32)vRowTemp_X.GetCell(5).NumericCellValue; //插卡筆數
                                vDriverInputShift = (Int32)vRowTemp_X.GetCell(6).NumericCellValue; //駕駛員輸入筆數
                                vIdentificationRate = (vRowTemp_X.GetCell(7).NumericCellValue / 100.0); //身份識別率
                                vOnTimeTotalShift = 0; //準確率總筆數
                                vOntimeShift = 0; //準確筆數
                                vOntimeRate = 0.0; //準確率
                            }
                            else if (vInputType == "班次統計")
                            {
                                vCompany = vRowTemp_X.GetCell(0).StringCellValue;
                                vStation = vRowTemp_X.GetCell(1).StringCellValue;
                                vLinesNote = vRowTemp_X.GetCell(3).StringCellValue;
                                vApprovedShift = (Int32)vRowTemp_X.GetCell(4).NumericCellValue; //核定班次數
                                vQualifiedShift = (Int32)vRowTemp_X.GetCell(5).NumericCellValue; //合格班次數
                                vUnqualifiedShift = (Int32)vRowTemp_X.GetCell(7).NumericCellValue; //不合格班次數
                                vQualifiedRate = vRowTemp_X.GetCell(6).NumericCellValue; //合格比率
                                vIdentificationShift = 0; //身份識別總筆數
                                vInsertIDCardShift = 0; //插卡筆數
                                vDriverInputShift = 0; //駕駛員輸入筆數
                                vIdentificationRate = 0.0; //身份識別率
                                vOnTimeTotalShift = 0; //準確率總筆數
                                vOntimeShift = 0; //準確筆數
                                vOntimeRate = 0.0; //準確率
                            }
                            vPosIndex = vLinesNote.IndexOf("】");
                            vLinesNo = vLinesNote.Substring(1, vPosIndex - 1);
                            vLinesNote = vLinesNote.Substring(vPosIndex + 1).Replace(" ", "");
                            vShiftDate = DateTime.Parse(eSourceDate.Text.Trim());
                            vBuMan = vLoginID;
                            vBuDate = DateTime.Today;
                            vShiftMode = "03";
                            vIndexNo = "0000000000" + vLinesNo;
                            vIndexNo = (vShiftDate.Year > 3822) ?
                                          vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                          (vShiftDate.Year - 1911).ToString("D4") +
                                          vShiftDate.Month.ToString("D2") +
                                          vShiftDate.Day.ToString("D2") :
                                       (vShiftDate.Year < 1911) ?
                                          vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                          (vShiftDate.Year + 1911).ToString("D4") +
                                          vShiftDate.Month.ToString("D2") +
                                          vShiftDate.Day.ToString("D2") :
                                       vShiftMode + vIndexNo.Substring(vLinesNo.Length, 10) +
                                          vShiftDate.Year.ToString("D4") +
                                          vShiftDate.Month.ToString("D2") +
                                          vShiftDate.Day.ToString("D2");
                            vShiftDate_Str = (vShiftDate.Year > 3822) ? (vShiftDate.Year - 1911).ToString("D4") + "/" + vShiftDate.ToString("MM/dd") :
                                             (vShiftDate.Year < 1911) ? (vShiftDate.Year + 1911).ToString("D4") + "/" + vShiftDate.ToString("MM/dd") :
                                             vShiftDate.Year.ToString("D4") + "/" + vShiftDate.ToString("MM/dd");
                            vBuDate_Str = (vBuDate.Year > 3822) ? (vBuDate.Year - 1911).ToString("D4") + "/" + vBuDate.ToString("MM/dd") :
                                          (vBuDate.Year < 1911) ? (vBuDate.Year + 1911).ToString("D4") + "/" + vBuDate.ToString("MM/dd") :
                                          vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd");

                            vDataCountStr = "select Count(IndexNo) RCount " + Environment.NewLine +
                                            "  from LinesStatistics " + Environment.NewLine +
                                            " where LinesNo = '" + vLinesNo + "' " + Environment.NewLine +
                                            "   and ShiftDate = '" + vShiftDate_Str + "' " + Environment.NewLine +
                                            "   and ShiftMode = '" + vShiftMode + "' ";
                            vDataCount = Int32.Parse(PF.GetValue(vConnStr, vDataCountStr, "RCount"));
                            if (vDataCount == 0)
                            {
                                sdsLinesStatistics.InsertParameters.Clear();
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("LinesNote", DbType.String, vLinesNote));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("ShiftDate", DbType.Date, vShiftDate_Str));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("ApprovedShift", DbType.Int32, vApprovedShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("QualifiedShift", DbType.Int32, vQualifiedShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("UnqualifiedShift", DbType.Int32, vUnqualifiedShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("QualifiedRate", DbType.Double, vQualifiedRate.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("BuMan", DbType.String, vBuMan));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDate_Str));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("Station", DbType.String, vStation));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("IdentificationShift", DbType.Int32, vIdentificationShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("InsertIDCardShift", DbType.Int32, vInsertIDCardShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("DriverInputShift", DbType.Int32, vDriverInputShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("IdentificationRate", DbType.Double, vIdentificationRate.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("OnTimeTotalShift", DbType.Int32, vOnTimeTotalShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("OntimeShift", DbType.Int32, vOntimeShift.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("OntimeRate", DbType.Double, vOntimeRate.ToString()));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("Remark", DbType.String, ""));
                                sdsLinesStatistics.InsertParameters.Add(new Parameter("ShiftMode", DbType.String, vShiftMode));
                                sdsLinesStatistics.Insert();
                            }
                            else if (vDataCount > 0)
                            {
                                vUpdateStr = "update LinesStatistics " + Environment.NewLine +
                                             "   set Station = case when Station is null then '" + vStation + "' else Station end, " + Environment.NewLine +
                                             "       ApprovedShift = ApprovedShift + " + vApprovedShift + ", " + Environment.NewLine +
                                             "       QualifiedShift = QualifiedShift + " + vQualifiedShift + ", " + Environment.NewLine +
                                             "       UnqualifiedShift = UnqualifiedShift + " + vUnqualifiedShift + ", " + Environment.NewLine +
                                             "       QualifiedRate = case when cast(ApprovedShift + " + vApprovedShift + " as float) = 0 then 0 else cast(QualifiedShift + " + vQualifiedShift + " as float) / cast(ApprovedShift + " + vApprovedShift + " as float) end, " + Environment.NewLine +
                                             "       IdentificationShift = IdentificationShift + " + vIdentificationShift + ", " + Environment.NewLine +
                                             "       InsertIDCardShift = InsertIDCardShift + " + vInsertIDCardShift + ", " + Environment.NewLine +
                                             "       DriverInputShift = DriverInputShift + " + vDriverInputShift + ", " + Environment.NewLine +
                                             "       IdentificationRate = case when cast(IdentificationShift + " + vIdentificationShift + " as float) = 0 then 0 else cast(InsertIDCardShift + " + vInsertIDCardShift + " + DriverInputShift + " + vDriverInputShift + " as float) / cast(IdentificationShift + " + vIdentificationShift + " as float) end, " + Environment.NewLine +
                                             "       OnTimeTotalShift = OnTimeTotalShift + " + vOnTimeTotalShift + ", " + Environment.NewLine +
                                             "       OntimeShift = OntimeShift + " + vOntimeShift + ", " + Environment.NewLine +
                                             "       OntimeRate = case when cast(OnTimeTotalShift + " + vOnTimeTotalShift + " as float) = 0 then 0 else cast(OntimeShift + " + vOntimeShift + " as float) / cast(OnTimeTotalShift + " + vOnTimeTotalShift + " as float) end " + Environment.NewLine +
                                             " where LinesNo = '" + vLinesNo + "' " + Environment.NewLine +
                                             "   and ShiftDate = '" + vShiftDate_Str + "' " + Environment.NewLine +
                                             "   and ShiftMode = '" + vShiftMode + "' ";
                                PF.ExecSQL(vConnStr, vUpdateStr);
                            }
                        }
                    }
                }
            }
        }

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            switch (rbListType.SelectedValue)
            {
                case "01": //市區公車路線
                    Preview_01();
                    break;
                case "02": //免費巴士路線
                    Preview_02();
                    break;
                case "03": //國道公路動態系統
                    Preview_03();
                    break;
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            switch (rbListType.SelectedValue)
            {
                case "01": //市區公車路線
                    ExportExcel_01();
                    break;
                case "02": //免費巴士路線
                    ExportExcel_02();
                    break;
                case "03": //國道公路動態系統
                    ExportExcel_03();
                    break;
            }
        }

        protected void bbImport_Click(object sender, EventArgs e)
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            string vShiftMode = rbListType.SelectedValue.Trim();

            if (fuExcel.FileName.Trim() != "")
            {
                switch (vShiftMode)
                {
                    case "01": //市區公車路線
                        ImportExcelData_01(vUploadPath + fuExcel.FileName.Trim());
                        break;

                    case "02": //免費巴士路線
                        ImportExcelData_01(vUploadPath + fuExcel.FileName.Trim());
                        break;

                    case "03": //國道公路動態系統
                        if (eSourceDate.Text.Trim() != "")
                        {
                            ImportExcelData_02(vUploadPath + fuExcel.FileName.Trim());
                        }
                        else
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('請先指定匯入資料的日期！')");
                            Response.Write("</" + "Script>");
                        }
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇要匯入的檔案！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plSearch.Visible = true;
            plReport.Visible = false;
        }

        protected void sdsLinesStatistics_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            Response.Write("<Script language='Javascript'>");
            Response.Write("alert('資料轉入完畢！')");
            Response.Write("</" + "Script>");
        }
    }
}