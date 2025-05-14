using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Reporting.WebForms;
using NPOI.SS.Util;
using System.IO;

namespace TyBus_Intranet_Test_V3
{
    public partial class ExportRentCarData : System.Web.UI.Page
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
                    if (!IsPostBack)
                    {
                        eCalYear_Search.Text = DateTime.Today.Year.ToString();
                        eCalMonth_Search.Text = DateTime.Today.Month.ToString();
                        plSearch.Visible = true;
                        plPrint.Visible = false;
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

        private string GetSumaryByDep()
        {
            int vYear_Int = 0;
            int vMonth_Int = 0;
            string vYear_Real = Int32.TryParse(eCalYear_Search.Text.Trim(), out vYear_Int) ? vYear_Int.ToString() : DateTime.Today.Year.ToString();
            string vYear_Last = Int32.TryParse(eCalYear_Search.Text.Trim(), out vYear_Int) ? (vYear_Int - 1).ToString() : (DateTime.Today.Year - 1).ToString();
            string vMonth_Real = Int32.TryParse(eCalMonth_Search.Text.Trim(), out vMonth_Int) ? vMonth_Int.ToString() : DateTime.Today.Month.ToString();
            string vResultStr = "select t.DepNo, d.[Name], sum(t.AMT_1) TotalAMT_1, sum(t.AMT_2) TotalAMT_2, sum(t.AMT_3) TotalAMT_3, sum(t.AMT_4) TotalAMT_4 " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "       select rd.DepNo, isnull(rd.AMT, 0) AMT_1, cast(0 as float) AMT_2, cast(0 as float) AMT_3, cast(0 as float) AMT_4 " + Environment.NewLine +
                                "         from RentD rd left join Car_InfoA ca on ca.Car_ID = rd.Car_ID " + Environment.NewLine +
                                "        where year(rd.RentDate) = " + vYear_Last + " and month(rd.RentDate) = " + vMonth_Real + Environment.NewLine +
                                "          and ca.Point in ('A1', 'A2', 'B2') " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select rd.DepNo, cast(0 as float) AMT_1, isnull(rd.AMT, 0) AMT_2, cast(0 as float) AMT_3, cast(0 as float) AMT_4 " + Environment.NewLine +
                                "         from RentD rd left join Car_InfoA ca on ca.Car_ID = rd.Car_ID " + Environment.NewLine +
                                "        where year(rd.RentDate) = " + vYear_Real + " and month(rd.RentDate) = " + vMonth_Real + Environment.NewLine +
                                "          and ca.Point in ('A1', 'A2', 'B2') " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select rd.DepNo, cast(0 as float) AMT_1, cast(0 as float) AMT_2, isnull(rd.AMT, 0) AMT_3, cast(0 as float) AMT_4 " + Environment.NewLine +
                                "         from RentD rd left join Car_InfoA ca on ca.Car_ID = rd.Car_ID " + Environment.NewLine +
                                "        where year(rd.RentDate) = " + vYear_Last + " and month(rd.RentDate) = " + vMonth_Real + Environment.NewLine +
                                "          and ca.Point = 'C1' " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select rd.DepNo, cast(0 as float) AMT_1, cast(0 as float) AMT_2, cast(0 as float) AMT_3, isnull(rd.AMT, 0) AMT_4 " + Environment.NewLine +
                                "         from RentD rd left join Car_InfoA ca on ca.Car_ID = rd.Car_ID " + Environment.NewLine +
                                "        where year(rd.RentDate) = " + vYear_Real + " and month(rd.RentDate) = " + vMonth_Real + Environment.NewLine +
                                "          and ca.Point = 'C1' " + Environment.NewLine +
                                "       ) t left join Department d on d.DepNo = t.DepNo " + Environment.NewLine +
                                " group by t.DepNo, d.[Name]" + Environment.NewLine +
                                " order by t.DepNo";

            return vResultStr;
        }

        private string GetSumaryByCar()
        {
            int vYear_Int = 0;
            int vMonth_Int = 0;
            string vYear_Real = Int32.TryParse(eCalYear_Search.Text.Trim(), out vYear_Int) ? vYear_Int.ToString() : DateTime.Today.Year.ToString();
            string vYear_Last = Int32.TryParse(eCalYear_Search.Text.Trim(), out vYear_Int) ? (vYear_Int - 1).ToString() : (DateTime.Today.Year - 1).ToString();
            string vMonth_Real = Int32.TryParse(eCalMonth_Search.Text.Trim(), out vMonth_Int) ? vMonth_Int.ToString() : DateTime.Today.Month.ToString();
            string vResultStr = "select t.Car_ID, sum(t.AMT_1) TotalAMT_1, sum(t.AMT_2) TotalAMT_2, sum(t.AMT_3) TotalAMT_3, sum(t.AMT_4) TotalAMT_4 " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "       select rd.Car_ID, isnull(rd.AMT, 0) AMT_1, cast(0 as float) AMT_2, cast(0 as float) AMT_3, cast(0 as float) AMT_4 " + Environment.NewLine +
                                "         from RentD rd left join Car_InfoA ca on ca.Car_ID = rd.Car_ID " + Environment.NewLine +
                                "        where year(rd.RentDate) = " + vYear_Last + " and month(rd.RentDate) = " + vMonth_Real + Environment.NewLine +
                                "          and ca.Point = 'C1' " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select rd.Car_ID, cast(0 as float) AMT_1, isnull(rd.AMT, 0) AMT_2, cast(0 as float) AMT_3, cast(0 as float) AMT_4 " + Environment.NewLine +
                                "         from RentD rd left join Car_InfoA ca on ca.Car_ID = rd.Car_ID " + Environment.NewLine +
                                "        where year(rd.RentDate) = " + vYear_Real + " and month(rd.RentDate) = " + vMonth_Real + Environment.NewLine +
                                "          and ca.Point = 'C1' " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select rd.Car_ID, cast(0 as float) AMT_1, cast(0 as float) AMT_2, isnull(rd.AMT, 0) AMT_3, cast(0 as float) AMT_4 " + Environment.NewLine +
                                "         from RentD rd left join Car_InfoA ca on ca.Car_ID = rd.Car_ID " + Environment.NewLine +
                                "        where year(rd.RentDate) = " + vYear_Last + " and month(rd.RentDate) between 1 and " + vMonth_Real + Environment.NewLine +
                                "          and ca.Point = 'C1' " + Environment.NewLine +
                                "        union all " + Environment.NewLine +
                                "       select rd.Car_ID, cast(0 as float) AMT_1, cast(0 as float) AMT_2, cast(0 as float) AMT_3, isnull(rd.AMT, 0) AMT_4 " + Environment.NewLine +
                                "         from RentD rd left join Car_InfoA ca on ca.Car_ID = rd.Car_ID " + Environment.NewLine +
                                "        where year(rd.RentDate) = " + vYear_Real + " and month(rd.RentDate) between 1 and " + vMonth_Real + Environment.NewLine +
                                "          and ca.Point = 'C1' " + Environment.NewLine +
                                "       ) t " + Environment.NewLine +
                                " group by t.Car_ID" + Environment.NewLine +
                                " order by t.Car_ID";
            return vResultStr;
        }

        private string GetSumaryByCar_2()
        {
            int vYear_Int = 0;
            int vMonth_Int = 0;
            string vYear_Real = Int32.TryParse(eCalYear_Search.Text.Trim(), out vYear_Int) ? vYear_Int.ToString() : DateTime.Today.Year.ToString();
            string vYear_Last = Int32.TryParse(eCalYear_Search.Text.Trim(), out vYear_Int) ? (vYear_Int - 1).ToString() : (DateTime.Today.Year - 1).ToString();
            string vMonth_Real = Int32.TryParse(eCalMonth_Search.Text.Trim(), out vMonth_Int) ? vMonth_Int.ToString() : DateTime.Today.Month.ToString();
            string vResultStr = "select ca.Car_ID, " + Environment.NewLine +
                                "       sum(isnull(t.AMT_1, 0)) TotalAMT_1, sum(isnull(t.AMT_2, 0)) TotalAMT_2, " + Environment.NewLine +
                                "       sum(isnull(t.AMT_3, 0)) TotalAMT_3, sum(isnull(t.AMT_4, 0)) TotalAMT_4  " + Environment.NewLine +
                                "  from Car_InfoA ca left join " + Environment.NewLine +
                                "       (select Car_ID, isnull(AMT, 0) AMT_1, cast(0 as float) AMT_2, cast(0 as float) AMT_3, cast(0 as float) AMT_4 " + Environment.NewLine +
                                "          from RentD " + Environment.NewLine +
                                "         where year(RentDate) = " + vYear_Last + " and month(RentDate) = " + vMonth_Real + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select Car_ID, cast(0 as float) AMT_1, isnull(AMT, 0) AMT_2, cast(0 as float) AMT_3, cast(0 as float) AMT_4 " + Environment.NewLine +
                                "          from RentD " + Environment.NewLine +
                                "         where year(RentDate) = " + vYear_Real + " and month(RentDate) = " + vMonth_Real + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select Car_ID, cast(0 as float) AMT_1, cast(0 as float) AMT_2, isnull(AMT, 0) AMT_3, cast(0 as float) AMT_4 " + Environment.NewLine +
                                "          from RentD " + Environment.NewLine +
                                "         where year(RentDate) = " + vYear_Last + " and month(RentDate) between 1 and " + vMonth_Real + Environment.NewLine +
                                "         union all " + Environment.NewLine +
                                "        select Car_ID, cast(0 as float) AMT_1, cast(0 as float) AMT_2, cast(0 as float) AMT_3, isnull(AMT, 0) AMT_4 " + Environment.NewLine +
                                "          from RentD " + Environment.NewLine +
                                "         where year(RentDate) = " + vYear_Real + " and month(RentDate) between 1 and " + vMonth_Real + Environment.NewLine +
                                "       ) t on t.Car_ID = ca.Car_ID " + Environment.NewLine +
                                " where ca.Point = 'C1' and ca.Tran_Type < 3 " + Environment.NewLine +
                                " group by ca.Car_ID " + Environment.NewLine +
                                " order by ca.Car_ID ";
            return vResultStr;
        }

        protected void bbPrintByDep_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSumaryByDep();
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelectStr, connTemp);
                connTemp.Open();
                DataTable dtPrint = new DataTable();
                daTemp.Fill(dtPrint);
                if (dtPrint.Rows.Count > 0)
                {
                    int vYear = 0;
                    int vMonth = 0;
                    string vCompanyName = PF.GetValue(vConnStr, "select [Name] from Custom where Types = 'O' and Code = 'A000'", "Name");
                    string vReportName = "站別租車收入報表";
                    string vDataYear_Real = Int32.TryParse(eCalYear_Search.Text.Trim(), out vYear) ? vYear.ToString() : DateTime.Today.Year.ToString();
                    string vDataYear_Last = (Int32.Parse(vDataYear_Real) - 1).ToString();
                    string vDataMonth = Int32.TryParse(eCalMonth_Search.Text.Trim(), out vMonth) ? vMonth.ToString("D2") : DateTime.Today.Month.ToString("D2");
                    string vDataYM_Real = vDataYear_Real + "." + vDataMonth;
                    string vDataYM_Last = vDataYear_Last + "." + vDataMonth;

                    ReportDataSource rdsPrint = new ReportDataSource("RentDataByDep", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\RentDataByDepP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("DataYM_Real", vDataYM_Real));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("DataYM_Last", vDataYM_Last));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plSearch.Visible = false;

                    string vRecordNote = "預覽報表_站別租車收入報表" + Environment.NewLine +
                                         "ExportRentCarData.aspx" + Environment.NewLine +
                                         "年月別：" + vDataYM_Real;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('查無資料')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void bbPrintByCar_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSumaryByCar_2();
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelectStr, connTemp);
                connTemp.Open();
                DataTable dtPrint = new DataTable();
                daTemp.Fill(dtPrint);
                if (dtPrint.Rows.Count > 0)
                {
                    int vYear = 0;
                    int vMonth = 0;
                    string vCompanyName = PF.GetValue(vConnStr, "select [Name] from Custom where Types = 'O' and Code = 'A000'", "Name");
                    string vReportName = "車輛別遊覽租車收入報表";
                    string vDataYear_Real = Int32.TryParse(eCalYear_Search.Text.Trim(), out vYear) ? vYear.ToString() : DateTime.Today.Year.ToString();
                    string vDataYear_Last = (Int32.Parse(vDataYear_Real) - 1).ToString();
                    string vDataMonth = Int32.TryParse(eCalMonth_Search.Text.Trim(), out vMonth) ? vMonth.ToString("D2") : DateTime.Today.Month.ToString("D2");
                    string vDataYM_Real = vDataYear_Real + "." + vDataMonth;
                    string vDataYM_Last = vDataYear_Last + "." + vDataMonth;
                    string vTotalYM_Real = vDataYear_Real + ".01-" + vDataMonth;
                    string vTotalYM_Last = vDataYear_Last + ".01-" + vDataMonth;

                    ReportDataSource rdsPrint = new ReportDataSource("RentDataByCar", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\RentDataByCarP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("DataYM_Real", vDataYM_Real));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("DataYM_Last", vDataYM_Last));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("TotalYM_Real", vTotalYM_Real));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("TotalYM_Last", vTotalYM_Last));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plSearch.Visible = false;

                    string vRecordNote = "預覽報表_車輛別遊覽租車收入報表" + Environment.NewLine +
                                         "ExportRentCarData.aspx" + Environment.NewLine +
                                         "年月別：" + vDataYM_Real;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('查無資料')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

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
            csData_Float.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_Float.Alignment = HorizontalAlignment.Right;

            HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
            csData_Float.DataFormat = format.GetFormat("##0.00");

            string vHeaderText = "";
            int vLinesNo = 0;
            //DateTime vBuDate;

            int vYear = 0;
            int vMonth = 0;
            string vCompanyName = PF.GetValue(vConnStr, "select [Name] from Custom where Types = 'O' and Code = 'A000'", "Name");
            string vSheetName = "";
            string vReportName = "租車收入報表";
            string vDataYear_Real = Int32.TryParse(eCalYear_Search.Text.Trim(), out vYear) ? vYear.ToString() : DateTime.Today.Year.ToString();
            string vDataYear_Last = (Int32.Parse(vDataYear_Real) - 1).ToString();
            string vDataMonth = Int32.TryParse(eCalMonth_Search.Text.Trim(), out vMonth) ? vMonth.ToString("D2") : DateTime.Today.Month.ToString("D2");
            string vDataYM_Real = vDataYear_Real + "." + vDataMonth;
            string vDataYM_Last = vDataYear_Last + "." + vDataMonth;
            string vTotalYM_Real = vDataYear_Real + ".01-" + vDataMonth;
            string vTotalYM_Last = vDataYear_Last + ".01-" + vDataMonth;
            string vSelectStr = "";

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                vSelectStr = GetSumaryByDep();
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //有資料才做下面的動作
                    vSheetName = "站別租車收入表";
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入抬頭
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(vCompanyName);
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    wsExcel.AddMergedRegion(new CellRangeAddress(0, 0, 0, 5));
                    //寫入標題列
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "DEPNO") ? "站別編號" :
                                      (drExcel.GetName(i).ToUpper() == "NAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMT_1") ? vDataYM_Last + "區間租車收入" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMT_2") ? vDataYM_Real + "區間租車收入" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMT_3") ? vDataYM_Last + "遊覽租車收入" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMT_4") ? vDataYM_Real + "遊覽租車收入" : "";
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    //開始寫入資料
                    while (drExcel.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if ((drExcel.GetName(i).ToUpper() == "DEPNO") ||
                                (drExcel.GetName(i).ToUpper() == "NAME"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Float;
                            }
                        }
                    }
                    //寫入小計
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("合計");
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData;
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 1));
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellFormula("SUM(C3:C" + (vLinesNo - 1).ToString() + ")");
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_Float;
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellFormula("SUM(D3:D" + (vLinesNo - 1).ToString() + ")");
                    wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csData_Float;
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellFormula("SUM(E3:E" + (vLinesNo - 1).ToString() + ")");
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_Float;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellFormula("SUM(F3:F" + (vLinesNo - 1).ToString() + ")");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_Float;
                }
            }
            vLinesNo = 0;
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                vSelectStr = GetSumaryByCar_2();
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //有資料才做下面的動作
                    vSheetName = "車輛別租車收入表";
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入抬頭
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(vCompanyName);
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    wsExcel.AddMergedRegion(new CellRangeAddress(0, 0, 0, 5));
                    //寫入標題列
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "牌照號碼" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMT_1") ? vDataYM_Last + "遊覽租車收入" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMT_2") ? vDataYM_Real + "遊覽租車收入" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMT_3") ? vTotalYM_Last + "收入" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALAMT_4") ? vTotalYM_Real + "收入" : "";
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    //開始寫入資料
                    while (drExcel.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if (drExcel.GetName(i).ToUpper() == "CAR_ID")
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Float;
                            }
                        }
                    }
                    //寫入小計
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("合計");
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData;
                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellFormula("SUM(B3:B" + (vLinesNo - 1).ToString() + ")");
                    wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csData_Float;
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellFormula("SUM(C3:C" + (vLinesNo - 1).ToString() + ")");
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData_Float;
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellFormula("SUM(D3:D" + (vLinesNo - 1).ToString() + ")");
                    wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csData_Float;
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellFormula("SUM(E3:E" + (vLinesNo - 1).ToString() + ")");
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData_Float;
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
                    string vRecordNote = "預覽報表_車輛別遊覽租車收入報表" + Environment.NewLine +
                                         "ExportRentCarData.aspx" + Environment.NewLine +
                                         "年月別：" + vDataYM_Real;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                    string TourVerision = brObject.Type;
                    if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                    {
                        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("[" + vDataYM_Real + "]" + vReportName + ".xls", System.Text.Encoding.UTF8));
                    }
                    else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                    {
                        // 設定強制下載標頭
                        Response.AddHeader("Content-Disposition", string.Format("attachment; filename=[" + vDataYM_Real + "]" + vReportName + ".xls"));
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
                Response.Write("alert('" + eMessage.Message + "')");
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
            plPrint.Visible = false;
        }
    }
}