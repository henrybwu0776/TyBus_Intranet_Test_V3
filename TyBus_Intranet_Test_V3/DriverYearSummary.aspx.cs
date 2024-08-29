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

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverYearSummary : System.Web.UI.Page
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

                    //事件日期
                    string vCalDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalYM_S.ClientID;
                    string vCalDateScript = "window.open('" + vCalDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalYM_S.Attributes["onClick"] = vCalDateScript;

                    vCalDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalYM_E.ClientID;
                    vCalDateScript = "window.open('" + vCalDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalYM_E.Attributes["onClick"] = vCalDateScript;

                    if (!IsPostBack)
                    {
                        plReport.Visible = false;
                        eCalYM_S.Text = PF.GetMonthFirstDay(DateTime.Today.AddYears(-1), "C");
                        eCalYM_E.Text = PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C");
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

        private string GetSelectStr(string fDriverList)
        {
            DateTime vCalYM_S = DateTime.Parse(eCalYM_S.Text.Trim());
            string vCalYMStr_S = vCalYM_S.Year.ToString() + "/" + vCalYM_S.ToString("MM/dd");
            DateTime vCalYM_E = DateTime.Parse(eCalYM_E.Text.Trim());
            string vCalYMStr_E = vCalYM_E.Year.ToString() + "/" + vCalYM_E.ToString("MM/dd");
            string vResultStr = "select empno as driver, (select [Name] from Employee where EmpNo = a.EmpNo) DriverName, " + Environment.NewLine +
                                "       cast(cast(sum(isnull(WorkHR, 0)*60 + isnull(WorkMin, 0)) as int) / 60 as nvarchar) + ':' + right('0' + cast(cast(sum(isnull(WorkHR, 0)*60 + isnull(WorkMin, 0)) as int) % 60 as nvarchar), 2) as TotalWorkHR, " + Environment.NewLine +
                                "       sum(isnull(ACashAmt, 0) + isnull(ACashAmt1, 0)) as ACashAMT, " + Environment.NewLine +
                                "       sum(isnull(Aoamt, 0) + isnull(Aoamt1, 0)) as Aoamt, " + Environment.NewLine +
                                "       sum(isnull(Aspiece, 0) + isnull(Aspiece1, 0)) as Aspiece, " + Environment.NewLine +
                                "       sum(isnull(AunTraincome, 0)) as AunTraincome, " + Environment.NewLine +
                                "       sum(isnull(ARentIncomeA, 0)) as ARentIncomeA, " + Environment.NewLine +
                                "       sum(isnull(ARentIncomeB, 0)) as ARentIncomeB, " + Environment.NewLine +
                                "       sum(isnull(ALines, 0) + isnull(ALines1, 0)) as ALines, " + Environment.NewLine +
                                "       sum(isnull(ACashAmt, 0) + isnull(ACashAmt1, 0) + isnull(Aoamt, 0) + isnull(Aoamt1, 0) + isnull(Aspiece, 0) + isnull(Aspiece1, 0) + isnull(AunTraincome, 0) + isnull(ARentIncomeA, 0) + isnull(ARentIncomeB, 0) + isnull(ALines, 0) + isnull(ALines1, 0)) as TotalIncome, " + Environment.NewLine +
                                "       sum(isnull(CBus1Km, 0)) as CBus1Km, " + Environment.NewLine +
                                "       sum(isnull(CBus2Km, 0)) as CBus2Km, " + Environment.NewLine +
                                "       sum(isnull(CRentTraKm, 0)) as CRentTraKm, " + Environment.NewLine +
                                "       sum(isnull(CRentAKm, 0)) as CRentAKm, " + Environment.NewLine +
                                "       sum(isnull(CRentBKm, 0)) as CRentBKm, " + Environment.NewLine +
                                "       sum(isnull(CBus3Km, 0)) as CBus3Km, " + Environment.NewLine +
                                "       sum(isnull(CBus4Km, 0)) as CBus4Km, " + Environment.NewLine +
                                "       sum(isnull(CBus1Km, 0) + isnull(CBus2Km, 0) + isnull(CRentTraKm, 0) + isnull(CRentAKm, 0) + isnull(CRentBKm, 0) + isnull(CBus3Km, 0) + isnull(CBus4Km, 0)) as BunissKMs, " + Environment.NewLine +
                                "       sum(isnull(CBus5Km, 0)) as CBus5Km, " + Environment.NewLine +
                                "       sum(isnull(CBus1Km, 0) + isnull(CBus2Km, 0) + isnull(CRentTraKm, 0) + isnull(CRentAKm, 0) + isnull(CRentBKm, 0) + isnull(CBus3Km, 0) + isnull(CBus4Km, 0) + isnull(CBus5Km, 0)) as TotalKMs " + Environment.NewLine +
                                "  from RsEmpmonth a " + Environment.NewLine +
                                " where a.EmpNo in (" + fDriverList + ") " + Environment.NewLine +
                                "   and a.BuDate between '" + vCalYMStr_S + "' and '" + vCalYMStr_E + "' " + Environment.NewLine +
                                " group by a.empno " + Environment.NewLine +
                                " order by a.empno";
            return vResultStr;
        }

        /// <summary>
        /// 從 EXCEL 檔匯入要計算的駕駛員工號
        /// </summary>
        /// <returns></returns>
        private string GetDriverNoList()
        {
            string vResultStr = "";
            string vDriverNo = "";
            string vExtName = Path.GetExtension(fuDriverList.FileName);
            switch (vExtName)
            {
                case ".xls":
                    //Excel 97-2003
                    HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuDriverList.FileContent);
                    HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                    for (int i = 0; i <= sheetExcel_H.LastRowNum; i++)
                    {
                        HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                        if ((vRowTemp_H != null) && (vRowTemp_H.GetCell(0).StringCellValue.Trim() != ""))
                        {
                            vDriverNo = vRowTemp_H.GetCell(0).StringCellValue.Trim();
                            if (Int32.TryParse(vDriverNo, out int j))
                            {
                                vResultStr = (vResultStr == "") ? "'" + vDriverNo + "'" : vResultStr + ",'" + vDriverNo + "'";
                            }
                        }
                    }
                    break;

                case ".xlsx":
                    //Excel 2010--
                    XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuDriverList.FileContent);
                    XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                    for (int i = 0; i <= sheetExcel_X.LastRowNum; i++)
                    {
                        XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                        if ((vRowTemp_X != null) && (vRowTemp_X.GetCell(0).StringCellValue.Trim() != ""))
                        {
                            vDriverNo = vRowTemp_X.GetCell(0).StringCellValue.Trim();
                            if (Int32.TryParse(vDriverNo, out int j))
                            {
                                vResultStr = (vResultStr == "") ? "'" + vDriverNo + "'" : vResultStr + ",'" + vDriverNo + "'";
                            }
                        }
                    }
                    break;

                default:
                    vResultStr = "";
                    break;
            }
            return vResultStr;
        }

        /// <summary>
        /// 預覽報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vDriverList = GetDriverNoList();
            string vSelStr = GetSelectStr(vDriverList);
            DateTime vCalYM_S = DateTime.Parse(eCalYM_S.Text.Trim());
            string vCalYMStr_S = (vCalYM_S.Year - 1911).ToString() + "年" + vCalYM_S.Month.ToString("D2") + "月" + vCalYM_S.Day.ToString("D2") + "日";
            DateTime vCalYM_E = DateTime.Parse(eCalYM_E.Text.Trim());
            string vCalYMStr_E = (vCalYM_E.Year - 1911).ToString() + "年" + vCalYM_E.Month.ToString("D2") + "月" + vCalYM_E.Day.ToString("D2") + "日";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                DataTable dtPrint = new DataTable("DriverYearSummaryP");
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelStr, connPrint);
                connPrint.Open();
                daPrint.Fill(dtPrint);
                if (dtPrint.Rows.Count > 0)
                {
                    //有取得資料才開始預覽
                    //string vCompanyName = "桃園汽車客運股份有限公司";
                    string vTempStr = "select [Name] from Custom where Types = 'O' and Code = 'A000' ";
                    string vCompanyName = PF.GetValue(vConnStr, vTempStr, "Name");
                    string vReportName = "遊覽車評選收入、公里、總時數";
                    string vCalYM = "日期：" + vCalYMStr_S + "～" + vCalYMStr_E;
                    ReportDataSource rdsPrint = new ReportDataSource("DriverYearSummaryP", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\DriverYearSummaryP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CalYM", vCalYM));
                    rvPrint.LocalReport.Refresh();
                    plReport.Visible = true;
                }
            }
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vDriverList = GetDriverNoList();
            string vSelectStr = GetSelectStr(vDriverList);
            DateTime vCalYM_S = DateTime.Parse(eCalYM_S.Text.Trim());
            string vCalYMStr_S = (vCalYM_S.Year - 1911).ToString() + "年" + vCalYM_S.Month.ToString("D2") + "月" + vCalYM_S.Day.ToString("D2") + "日";
            DateTime vCalYM_E = DateTime.Parse(eCalYM_E.Text.Trim());
            string vCalYMStr_E = (vCalYM_E.Year - 1911).ToString() + "年" + vCalYM_E.Month.ToString("D2") + "月" + vCalYM_E.Day.ToString("D2") + "日";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
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
                    fontData.FontHeightInPoints = 12; //字體大小
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
                    //string vCompanyName = "桃園汽車客運股份有限公司";
                    string vTempStr = "select [Name] from Custom where Types = 'O' and Code = 'A000' ";
                    string vCompanyName = PF.GetValue(vConnStr, vTempStr, "Name");
                    string vFileName = "遊覽車評選收入、公里、總時數";
                    int vLinesNo = 0;
                    int vColumnCount = drExcel.FieldCount;
                    string vDateRange = "日期：" + vCalYMStr_S + "～" + vCalYMStr_E;

                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列-1 公司名稱
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(vCompanyName);
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, vColumnCount - 1));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    vLinesNo++;

                    //寫入標題列-2 報表名稱
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue(vFileName);
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, vColumnCount - 1));
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    vLinesNo++;

                    //寫入標題列-3 計算起迄日期
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(13).SetCellValue(vDateRange);
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 13, vColumnCount - 1));
                    wsExcel.GetRow(vLinesNo).GetCell(13).CellStyle = csData;
                    vLinesNo++;

                    //寫入標題列-4 統計群組
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("營運收入");
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 3, 10));
                    wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csTitle;

                    wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("行駛公里");
                    //wsExcel.AutoSizeColumn(0);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 11, vColumnCount - 1));
                    wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csTitle;
                    vLinesNo++;

                    //寫入標題列-5 真正的標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).Trim().ToUpper() == "DRIVER") ? "員工編號" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "DRIVERNAME") ? "姓名" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "TOTALWORKHR") ? "時數" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "ACASHAMT") ? "現投收入" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "AOAMT") ? "普+聯營" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "ASPIECE") ? "學生卡" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "AUNTRAINCOME") ? "交通車" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "ARENTINCOMEA") ? "區間租車" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "ARENTINCOMEB") ? "遊覽車" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "ALINES") ? "聯營路線" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "TOTALINCOME") ? "合計" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "CBUS1KM") ? "班車路線" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "CBUS2KM") ? "公車路線" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "CRENTTRAKM") ? "交通車" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "CRENTAKM") ? "區間租車" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "CRENTBKM") ? "遊覽車" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "CBUS3KM") ? "專車路線" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "CBUS4KM") ? "聯營路線" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "BUNISSKMS") ? "營運合計" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "CBUS5KM") ? "非營運" :
                                      (drExcel.GetName(i).Trim().ToUpper() == "TOTALKMS") ? "總計" : drExcel.GetName(i).Trim();

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
                            if (((drExcel.GetName(i).Trim().ToUpper() == "DRIVER") ||
                                 (drExcel.GetName(i).Trim().ToUpper() == "DRIVERNAME") ||
                                 (drExcel.GetName(i).Trim().ToUpper() == "TOTALWORKHR")) &&
                                (drExcel[i].ToString() != ""))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString().Trim()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
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

        /// <summary>
        /// 結束
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 關閉預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
        }
    }
}
