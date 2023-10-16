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
    public partial class EmpOverList : System.Web.UI.Page
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
                    string vBuDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eRealDay_S.ClientID;
                    string vBuDateScript = "window.open('" + vBuDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eRealDay_S.Attributes["onClick"] = vBuDateScript;

                    vBuDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eRealDay_E.ClientID;
                    vBuDateScript = "window.open('" + vBuDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eRealDay_E.Attributes["onClick"] = vBuDateScript;

                    if (!IsPostBack)
                    {
                        eRealDay_S.Text = PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C");
                        eRealDay_E.Text = PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C");
                        plShowData.Visible = false;
                        plReport.Visible = false;
                        bbPreview.Enabled = false;
                        bbExcel.Enabled = false;
                    }
                    else
                    {
                        OverListDataBind();
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

        private string GetSelStr()
        {
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and a.DepNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.DepNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ApplyMan = (eEmpNo.Text.Trim() != "") ? "   and a.ApplyMan = '" + eEmpNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_RealDay = ((eRealDay_S.Text.Trim() != "") && (eRealDay_E.Text.Trim() != "")) ? "   and a.RealDay between '" + DateTime.Parse(eRealDay_S.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eRealDay_S.Text.Trim()).ToString("MM/dd") + "' and '" + DateTime.Parse(eRealDay_E.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eRealDay_E.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                   ((eRealDay_S.Text.Trim() != "") && (eRealDay_E.Text.Trim() == "")) ? "   and a.RealDay = '" + DateTime.Parse(eRealDay_S.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eRealDay_S.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                   ((eRealDay_S.Text.Trim() == "") && (eRealDay_E.Text.Trim() != "")) ? "   and a.RealDay = '" + DateTime.Parse(eRealDay_E.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eRealDay_E.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vSelect = "select ApplyMan, (select [Name] from Employee where EmpNo = a.ApplyMan) ApplyName, " + Environment.NewLine +
                             "       DepNo, (select[Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                             "       RealDay, " + Environment.NewLine +
                             "       ApplyType, (select ClassTxt from DBDICB where ClassNo = a.ApplyType and FKey = '加班資料檔      OVERDUTY        APPLYTYPE') ApplyType_C, " + Environment.NewLine +
                             "       [Hours], case when ApplyType in ('02', '04') then[Hours] - FeedNum - BackNum else 0 end Over100, " + Environment.NewLine +
                             "       (isnull(FeedNum, 0) + isnull(FrontOver2, 0)) as Over133, " + Environment.NewLine +
                             "       (isnull(BackNum, 0) + isnull(PostOver2, 0)) as Over166, " + Environment.NewLine +
                             "       isnull(PostOver22, 0) as Over266 " + Environment.NewLine +
                             "  from OverDuty a " + Environment.NewLine +
                             " where 1 = 1 " + Environment.NewLine + vWStr_ApplyMan + vWStr_DepNo + vWStr_RealDay +
                             " order by DepNo, ApplyMan, RealDay";
            return vSelect;
        }

        private void OverListDataBind()
        {
            string vSelectStr = GetSelStr();
            sdsEmpOverList.SelectCommand = "";
            sdsEmpOverList.SelectCommand = vSelectStr;
            gridEmpOverList.DataBind();
            if (gridEmpOverList.Rows.Count > 0)
            {
                plShowData.Visible = true;
                bbPreview.Enabled = true;
                bbExcel.Enabled = true;
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OverListDataBind();
        }

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            string vReportName = "加班資料明細表";
            //統計報表
            string vSelectStr = GetSelStr();
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

                ReportDataSource rdsPrint = new ReportDataSource("EmpOverListP", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\EmpOverListP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.Refresh();
                plShowData.Visible = false;
                plReport.Visible = true;

                string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "自 " + eDepNo_S.Text.Trim() + " 站起至 " + eDepNo_E.Text.Trim() + "站止" :
                                   ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() + " 站" :
                                   ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() + " 站" : "全部";
                string vEmpNoStr = (eEmpNo.Text.Trim() != "") ? eEmpNo.Text.Trim() : "";
                string vRealDayStr = ((eRealDay_S.Text.Trim() != "") && (eRealDay_E.Text.Trim() != "")) ? "自 " + eRealDay_S.Text.Trim() + " 起至 " + eRealDay_E.Text.Trim() + " 止" :
                                     ((eRealDay_S.Text.Trim() != "") && (eRealDay_E.Text.Trim() == "")) ? eRealDay_S.Text.Trim() :
                                     ((eRealDay_S.Text.Trim() == "") && (eRealDay_E.Text.Trim() != "")) ? eRealDay_E.Text.Trim() : "";
                string vRecordNote = "預覽報表_加班資料明細表" + Environment.NewLine +
                                     "EmpOverList.aspx" + Environment.NewLine +
                                     "站別：" + vDepNoStr + Environment.NewLine +
                                     "員工工號：" + vEmpNoStr + Environment.NewLine +
                                     "加班日期：" + vRealDayStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
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
            string vFileName = "加班資料明細表";
            int vLinesNo = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelStr();
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
                        vHeaderText = (drExcel.GetName(i) == "ApplyMan") ? "員工編號" :
                                      (drExcel.GetName(i) == "ApplyName") ? "員工姓名" :
                                      (drExcel.GetName(i) == "DepNo") ? "部門代號" :
                                      (drExcel.GetName(i) == "DepName") ? "部門" :
                                      (drExcel.GetName(i) == "RealDay") ? "加班日期" :
                                      (drExcel.GetName(i) == "ApplyType") ? "加班類別代號" :
                                      (drExcel.GetName(i) == "ApplyType_C") ? "加班類別" :
                                      (drExcel.GetName(i) == "Hours") ? "加班時數" :
                                      (drExcel.GetName(i) == "Over100") ? "1倍時數" :
                                      (drExcel.GetName(i) == "Over133") ? "1.33時數" :
                                      (drExcel.GetName(i) == "Over166") ? "1.66時數" :
                                      (drExcel.GetName(i) == "Over266") ? "2.66時數" : drExcel.GetName(i);
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
                            if ((drExcel.GetName(i) == "Hours") ||
                                (drExcel.GetName(i) == "Over100") ||
                                (drExcel.GetName(i) == "Over133") ||
                                (drExcel.GetName(i) == "Over166") ||
                                (drExcel.GetName(i) == "Over266"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if ((drExcel.GetName(i) == "RealDay") && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 1911).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
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
                            string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "自 " + eDepNo_S.Text.Trim() + " 站起至 " + eDepNo_E.Text.Trim() + "站止" :
                                               ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() + " 站" :
                                               ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() + " 站" : "全部";
                            string vEmpNoStr = (eEmpNo.Text.Trim() != "") ? eEmpNo.Text.Trim() : "";
                            string vRealDayStr = ((eRealDay_S.Text.Trim() != "") && (eRealDay_E.Text.Trim() != "")) ? "自 " + eRealDay_S.Text.Trim() + " 起至 " + eRealDay_E.Text.Trim() + " 止" :
                                                 ((eRealDay_S.Text.Trim() != "") && (eRealDay_E.Text.Trim() == "")) ? eRealDay_S.Text.Trim() :
                                                 ((eRealDay_S.Text.Trim() == "") && (eRealDay_E.Text.Trim() != "")) ? eRealDay_E.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_加班資料明細表" + Environment.NewLine +
                                                 "EmpOverList.aspx" + Environment.NewLine +
                                                 "站別：" + vDepNoStr + Environment.NewLine +
                                                 "員工工號：" + vEmpNoStr + Environment.NewLine +
                                                 "加班日期：" + vRealDayStr;
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
            plShowData.Visible = true;
        }
    }
}