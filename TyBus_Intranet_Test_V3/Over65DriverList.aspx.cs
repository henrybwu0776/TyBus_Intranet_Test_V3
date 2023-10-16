using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class Over65DriverList : System.Web.UI.Page
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
                    //基準日
                    string vCountDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalBaseDate.ClientID;
                    string vCountDateScript = "window.open('" + vCountDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalBaseDate.Attributes["onClick"] = vCountDateScript;

                    if (!IsPostBack)
                    {
                        plSearch.Visible = true;
                        plShowData.Visible = true;
                        plPrint.Visible = false;
                        eCalBaseDate.Text = vToday.ToShortDateString();
                        eDepNo_Search.Text = "";
                        eDepName_Search.Text = "";
                    }
                    else
                    {
                        OpenData();
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
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and e.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            DateTime vCalDate_Temp_S = DateTime.Parse(eCalBaseDate.Text.Trim());
            string vCalDateStr_Temp_S = vCalDate_Temp_S.Year.ToString() + "/" + vCalDate_Temp_S.ToString("MM/dd");
            DateTime vCalDate_Temp_E = DateTime.Parse(eCalBaseDate.Text.Trim()).AddDays(60);
            string vCalDateStr_Temp_E = vCalDate_Temp_E.Year.ToString() + "/" + vCalDate_Temp_E.ToString("MM/dd");
            string vResultStr = "select DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                "       EmpNo, [Name], Birthday, IDCardNo, LicenceCheck, BBCall " + Environment.NewLine +
                                "  from Employee e " + Environment.NewLine +
                                " where Title in ('300', '301', '302', '303') " + Environment.NewLine +
                                "   and isnull(LeaveDay, '') = '' " + Environment.NewLine +
                                "   and '" + vCalDateStr_Temp_S + "' <= DATEADD(Year, 65, BirthDay) " + Environment.NewLine +
                                "   and '" + vCalDateStr_Temp_E + "' >= DATEADD(Year, 65, BirthDay) " + Environment.NewLine +
                                vWStr_DepNo +
                                " order by DepNo, EmpNo";
            return vResultStr;
        }

        private string GetSelectStr_Print()
        {
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and e.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            DateTime vCalDate_Temp_S = DateTime.Parse(eCalBaseDate.Text.Trim());
            string vCalDateStr_Temp_S = vCalDate_Temp_S.Year.ToString() + "/" + vCalDate_Temp_S.ToString("MM/dd");
            DateTime vCalDate_Temp_E = DateTime.Parse(eCalBaseDate.Text.Trim()).AddDays(60);
            string vCalDateStr_Temp_E = vCalDate_Temp_E.Year.ToString() + "/" + vCalDate_Temp_E.ToString("MM/dd");
            string vResultStr = "select DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                "       EmpNo, [Name], " + Environment.NewLine +
                                "       STUFF(CONVERT(VARCHAR(10), Birthday, 111), 1, 4,  YEAR(Birthday)-1911) as Birthday, " + Environment.NewLine +
                                "       IDCardNo, " + Environment.NewLine +
                                "       STUFF(CONVERT(VARCHAR(10), LicenceCheck, 111), 1, 4,  YEAR(LicenceCheck)-1911) as LicenceCheck, " + Environment.NewLine +
                                "       STUFF(CONVERT(VARCHAR(10), BBCall, 111), 1, 4,  YEAR(BBCall)-1911) as BBCall " + Environment.NewLine +
                                "  from Employee e " + Environment.NewLine +
                                " where Title in ('300', '301', '302', '303') " + Environment.NewLine +
                                "   and isnull(LeaveDay, '') = '' " + Environment.NewLine +
                                "   and '" + vCalDateStr_Temp_S + "' <= DATEADD(Year, 65, BirthDay) " + Environment.NewLine +
                                "   and '" + vCalDateStr_Temp_E + "' >= DATEADD(Year, 65, BirthDay) " + Environment.NewLine +
                                vWStr_DepNo +
                                " order by DepNo, EmpNo";
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSelectStr();
            sdsShowData.SelectCommand = vSelectStr;
            gridShowData.DataBind();
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDepNo_Search.Text.Trim();
            if (vDepNo_Temp != "")
            {
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
                string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName_Temp == "")
                {
                    vDepName_Temp = vDepNo_Temp.Trim();
                    vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                    vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_Search.Text = vDepNo_Temp.Trim();
                eDepName_Search.Text = vDepName_Temp.Trim();
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelectStr();
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

            string vHeaderText = "";
            int vLinesNo = 0;
            string vFileName = "年滿65歲駕駛員清冊";
            DateTime vBuDate;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "DEPNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "EMPNO") ? "工號" :
                                      (drExcel.GetName(i).ToUpper() == "NAME") ? "姓名" :
                                      (drExcel.GetName(i).ToUpper() == "BIRTHDAY") ? "出生日期" :
                                      (drExcel.GetName(i).ToUpper() == "IDCARDNO") ? "身分證號" :
                                      (drExcel.GetName(i).ToUpper() == "LICENCECHECK") ? "駕照審驗期限" :
                                      (drExcel.GetName(i).ToUpper() == "BBCALL") ? "定期訓練審驗期限" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i).ToUpper() == "BIRTHDAY") ||
                                 (drExcel.GetName(i).ToUpper() == "LICENCECHECK") ||
                                 (drExcel.GetName(i).ToUpper() == "BBCALL")) && (drExcel[i].ToString().Trim() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString().Trim());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue((vBuDate.Year - 1911).ToString() + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString().Trim());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
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
                            string vCalBaseDateStr = (eCalBaseDate.Text.Trim() != "") ? eCalBaseDate.Text.Trim() : "";
                            string vDepNoStr = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_年滿65歲駕駛員名冊" + Environment.NewLine +
                                                 "Over65DriverList.aspx" + Environment.NewLine +
                                                 "計算基準日：" + vCalBaseDateStr + Environment.NewLine +
                                                 "站別：" + vDepNoStr;
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

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelectStr_Print();
            string vReportName = "年滿65歲駕駛員清冊";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                DataTable dtPrint = new DataTable("Over65DriverListP");
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                daPrint.Fill(dtPrint);
                if (dtPrint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("Over65DriverListP", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\Over65DriverListP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;

                    string vCalBaseDateStr = (eCalBaseDate.Text.Trim() != "") ? eCalBaseDate.Text.Trim() : "";
                    string vDepNoStr = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "全部";
                    string vRecordNote = "預覽報表_年滿65歲駕駛員名冊" + Environment.NewLine +
                                         "Over65DriverList.aspx" + Environment.NewLine +
                                         "計算基準日：" + vCalBaseDateStr + Environment.NewLine +
                                         "站別：" + vDepNoStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plShowData.Visible = true;
            plSearch.Visible = true;
        }
    }
}