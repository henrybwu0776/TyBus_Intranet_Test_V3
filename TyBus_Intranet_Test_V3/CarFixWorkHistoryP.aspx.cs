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
using System.Web.UI;

namespace TyBus_Intranet_Test_V3
{
    public partial class CarFixWorkHistoryP : Page
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

                UnobtrusiveValidationMode = UnobtrusiveValidationMode.None;

                if (vLoginID != "")
                {
                    if (!IsPostBack)
                    {
                        eBudate_Start_Search.Text = PF.GetMonthFirstDay(vToday, "C");
                        eBudate_End_Search.Text = PF.GetMonthLastDay(vToday, "C");
                        eCarID_Search.Text = "";
                        plReport.Visible = false;
                        rvPrint.LocalReport.DataSources.Clear();
                        bbEXCEL.Visible = (vLoginDepNo == "09");
                    }

                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBudate_Start_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBudate_Start_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBudate_End_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBudate_End_Search.Attributes["onClick"] = vDateScript;
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
            string vSelectStr = "SELECT (SELECT CompanyNo FROM Car_InfoA WHERE (Car_ID = a.CAR_ID)) AS CompanyNo, " + Environment.NewLine +
                                "       (SELECT NAME FROM DepartMent WHERE (DEPNO = (SELECT CompanyNo FROM Car_InfoA AS Car_InfoA_1 WHERE (Car_ID = a.CAR_ID)))) AS CompanyNo_C, " + Environment.NewLine +
                                "       CAR_ID, BUDATE, BUMAN, (SELECT NAME FROM Employee WHERE (EMPNO = a.BUMAN)) AS BuMan_C, " + Environment.NewLine +
                                "       DEPNO, (SELECT NAME FROM Department AS Department_1 WHERE (DEPNO = a.DEPNO)) AS DepNo_C, " + Environment.NewLine +
                                "       service, (SELECT DESCRIPTION FROM DBDICB WHERE (CLASSNO = a.service) AND (FKEY = '工作單A         FixworkA        SERVICE')) AS Service_C, " + Environment.NewLine +
                                "       DRIVER, (SELECT NAME FROM EMployee AS EMployee_1 WHERE (EMPNO = a.DRIVER)) AS Driver_C, REMARK " + Environment.NewLine +
                                "  FROM FixWorkA AS a " + Environment.NewLine +
                                " WHERE (isnull(a.Car_ID, '') <> '') " + Environment.NewLine;
            string vWhereStr_CarID = (eCarID_Search.Text.Trim() != "") ? " and a.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWhereStr_BuDate = ((eBudate_Start_Search.Text.Trim() != "") && (eBudate_End_Search.Text.Trim() != "")) ? " and a.Budate between '" + DateTime.Parse(eBudate_Start_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBudate_Start_Search.Text.Trim()).ToString("MM/dd") + "' and '" + DateTime.Parse(eBudate_End_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBudate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                      ((eBudate_Start_Search.Text.Trim() != "") && (eBudate_End_Search.Text.Trim() == "")) ? " and a.Budate = '" + DateTime.Parse(eBudate_Start_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBudate_Start_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                      ((eBudate_Start_Search.Text.Trim() == "") && (eBudate_End_Search.Text.Trim() != "")) ? " and a.Budate = '" + DateTime.Parse(eBudate_End_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBudate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";
            vSelectStr = vSelectStr +
                         vWhereStr_CarID +
                         vWhereStr_BuDate +
                         " order by a.Car_ID, a.Budate ";
            return vSelectStr;
        }

        private DataTable GetSelectData()
        {
            DataTable dtTemp = new DataTable();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "IF object_id('tempdb..#Temp_CarHistory') IS NOT NULL " + Environment.NewLine +
                                  "BEGIN " + Environment.NewLine +
                                  "   DROP TABLE #Temp_CarHistory " + Environment.NewLine +
                                  "END";
            PF.ExecSQL(vConnStr, vSQLStr_Temp);
            string vWhereStr_CarID = (eCarID_Search.Text.Trim() != "") ? " and a.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWhereStr_BuDate = ((eBudate_Start_Search.Text.Trim() != "") && (eBudate_End_Search.Text.Trim() != "")) ? " and a.Budate between '" + DateTime.Parse(eBudate_Start_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBudate_Start_Search.Text.Trim()).ToString("MM/dd") + "' and '" + DateTime.Parse(eBudate_End_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBudate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                      ((eBudate_Start_Search.Text.Trim() != "") && (eBudate_End_Search.Text.Trim() == "")) ? " and a.Budate = '" + DateTime.Parse(eBudate_Start_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBudate_Start_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                      ((eBudate_Start_Search.Text.Trim() == "") && (eBudate_End_Search.Text.Trim() != "")) ? " and a.Budate = '" + DateTime.Parse(eBudate_End_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eBudate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";
            vSQLStr_Temp = "SELECT a.WorkNo, (SELECT CompanyNo FROM Car_InfoA WHERE (Car_ID = a.CAR_ID)) AS CompanyNo, " + Environment.NewLine +
                           "       (SELECT NAME FROM DepartMent WHERE (DEPNO = (SELECT CompanyNo FROM Car_InfoA AS Car_InfoA_1 WHERE (Car_ID = a.CAR_ID)))) AS CompanyNo_C, " + Environment.NewLine +
                           "       CAR_ID, BUDATE, BUMAN, (SELECT NAME FROM Employee WHERE (EMPNO = a.BUMAN)) AS BuMan_C, " + Environment.NewLine +
                           "       DEPNO, (SELECT NAME FROM Department AS Department_1 WHERE (DEPNO = a.DEPNO)) AS DepNo_C, " + Environment.NewLine +
                           "       service, (SELECT DESCRIPTION FROM DBDICB WHERE (CLASSNO = a.service) AND (FKEY = '工作單A         FixworkA        SERVICE')) AS Service_C, " + Environment.NewLine +
                           "       DRIVER, (SELECT NAME FROM EMployee AS EMployee_1 WHERE (EMPNO = a.DRIVER)) AS Driver_C, a.REMARK, " + Environment.NewLine +
                           "       (select [Name] from Employee where EmpNo = c.FixMan) FixManName " + Environment.NewLine +
                           "  into #Temp_CarHistory " + Environment.NewLine +
                           "  FROM FixWorkA AS a left join FixWorkC as c on c.WorkNo = a.WorkNo " + Environment.NewLine +
                           " WHERE (isnull(a.Car_ID, '') <> '') " + Environment.NewLine + vWhereStr_CarID + vWhereStr_BuDate +
                           " order by a.Car_ID, a.Budate " + Environment.NewLine +
                           "select distinct WorkNo, CompanyNo, CompanyNo_C, Car_ID, BuDate, BuMan, BuMan_C, DepNo, DepNo_C, [Service], Service_C, Driver, Driver_C, Remark, " + Environment.NewLine +
                           "       (select FixManName + ',' from #Temp_CarHistory t2 where t2.WorkNo = t1.WorkNo FOR XML PATH('')) FixManList " + Environment.NewLine +
                           "  from #Temp_CarHistory t1 ";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(vSQLStr_Temp, connTemp);
                connTemp.Open();
                daTemp.Fill(dtTemp);
            }
            return dtTemp;
        }

        private void OpenDataSource()
        {
            /* 2020.01.15 改用 TABLE 傳資料
            string vSelectStr = GetSelectStr();
            sdsCarFixWorkHistory_P.SelectCommand = "";
            sdsCarFixWorkHistory_P.SelectCommand = vSelectStr;
            gridCarFixWorkHistory_P.DataBind();
            */
            DataTable dtFixList = GetSelectData();
            gridCarFixWorkHistory_P.DataSourceID = "";
            gridCarFixWorkHistory_P.DataSource = dtFixList;

            //ReportDataSource rdsPrint = new ReportDataSource("DetailData", sdsCarFixWorkHistory_P);
            ReportDataSource rdsPrint = new ReportDataSource("DetailData", dtFixList);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\CarFixWorkHistory.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.Refresh();
            rvPrint.Visible = true;
        }

        protected void eCarID_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCarID = eCarID_Search.Text.Trim();
            string vSQLStr = "select count(Car_ID) RCount from Car_InfoA where Car_ID = '" + vCarID + "' and Tran_Type = '1' ";
            int vCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));
            if (vCount == 0)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('輸入車號不存在，請重新輸入！')");
                Response.Write("</" + "Script>");
                eCarID_Search.Focus();
            }
        }

        protected void bbEXCEL_Click(object sender, EventArgs e)
        {
            DataTable dtExcel = GetSelectData();
            if (dtExcel.Rows.Count > 0)
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
                string vFileName = "車輛資歷卡";
                DateTime vBuDate;
                //新增一個工作表
                wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                //寫入標題列
                vLinesNo = 0;
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < dtExcel.Columns.Count; i++)
                {
                    vHeaderText = (dtExcel.Columns[i].ColumnName.ToUpper() == "DEPNO") ? "維修時配置單位代號" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "DEPNO_C") ? "維修時配置單位" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "CAR_ID") ? "車號" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "COMPANYNO") ? "站別代號" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "COMPANYNO_C") ? "站別" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "DRIVER") ? "駕駛員代號" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "DRIVER_C") ? "駕駛員姓名" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "REMARK") ? "維修項目 / 主要零件" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "SERVICE") ? "維修代號編碼" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "SERVICE_C") ? "維修代號" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "BUMAN") ? "承修人工號" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "BUMAN_C") ? "承修人" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "WORKNO") ? "維修單號" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "FIXMANLIST") ? "維修員" :
                                  (dtExcel.Columns[i].ColumnName.ToUpper() == "BUDATE") ? "開單日" : dtExcel.Columns[i].ColumnName;
                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                }
                //開始寫入資料
                for (int j = 0; j < dtExcel.Rows.Count; j++)
                {
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < dtExcel.Columns.Count; i++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i);
                        if ((dtExcel.Columns[i].ColumnName.ToUpper() == "BUDATE") && (dtExcel.Rows[j][i].ToString() != ""))
                        {
                            string vTempStr = dtExcel.Rows[j][i].ToString();
                            vBuDate = DateTime.Parse(dtExcel.Rows[j][i].ToString());
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 1911).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(dtExcel.Rows[j][i].ToString());
                        }
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
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
                        string vRecordCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "全部";
                        string vRecordBuDateStr = ((eBudate_Start_Search.Text.Trim() != "") && (eBudate_End_Search.Text.Trim() != "")) ? eBudate_Start_Search.Text.Trim() + "~" + eBudate_End_Search.Text.Trim() :
                                                  ((eBudate_Start_Search.Text.Trim() != "") && (eBudate_End_Search.Text.Trim() == "")) ? eBudate_Start_Search.Text.Trim() :
                                                  ((eBudate_Start_Search.Text.Trim() == "") && (eBudate_End_Search.Text.Trim() != "")) ? eBudate_End_Search.Text.Trim() : "不分日期"; ;
                        string vRecordNote = "匯出檔案_車輛資歷卡" + Environment.NewLine +
                                             "CarFixWorkHistoryP.aspx" + Environment.NewLine +
                                             "車號：" + vRecordCarIDStr + Environment.NewLine +
                                             "開單日期：" + vRecordBuDateStr;
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
                    Response.Write("alert('" + eMessage.Message.Trim() + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            if (eCarID_Search.Text.Trim() == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請輸入車號！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                plShowData.Visible = false;
                OpenDataSource();
                plReport.Visible = true;
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            rvPrint.LocalReport.DataSources.Clear();
            plReport.Visible = false;
        }
    }
}