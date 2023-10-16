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
    public partial class EmpPayList : System.Web.UI.Page
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
                    if (!IsPostBack)
                    {
                        DateTime vCalDate = DateTime.Today.AddMonths(-1);
                        ePayYear_Search.Text = (vCalDate.Year - 1911).ToString();
                        ePayMonth_Search.SelectedIndex = vCalDate.Month - 1;
                        plShowData.Visible = true;
                        plReport.Visible = false;
                    }
                    else
                    {
                        OpenDataList();
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
            string vPayDate_Start = (Int32.Parse(ePayYear_Search.Text.Trim()) + 1911).ToString("D4") + "/" +
                                         (Int32.Parse(ePayMonth_Search.Text.Trim())).ToString("D2") + "/01";
            string vPayDate_End = PF.GetMonthLastDay(DateTime.Parse(vPayDate_Start), "B");
            string vPayDur = ePayDur_Search.SelectedValue.Trim();
            string vResultStr = "select p.DepNo, (select [Name] from Department where DepNo = p.DepNo) DepName, " + Environment.NewLine +
                                "       p.EmpNo, p.[Name], e.WorkType, " + Environment.NewLine +
                                "       case when e.AssumeDay >= '" + vPayDate_Start + "' then AssumeDay else null end AssumeDay, e.LeaveDay, " + Environment.NewLine +
                                "       case when(isnull(e.StopDay, '" + vPayDate_Start + "') >= '" + vPayDate_Start + "' or e.BeginDay >= '" + vPayDate_Start + "') then e.BeginDay  else null end BeginDay, " + Environment.NewLine +
                                "       case when isnull(e.StopDay, '" + vPayDate_Start + "') >= '" + vPayDate_Start + "' then e.StopDay else null end StopDay, " + Environment.NewLine +
                                "       DATEDIFF(Day, " + Environment.NewLine +
                                "                case when isnull(e.StopDay, e.AssumeDay) >= '" + vPayDate_Start + "' then isnull(e.StopDay, e.AssumeDay) else '" + vPayDate_Start + "' end, " + Environment.NewLine +
                                "                case when isnull(e.LeaveDay, '" + vPayDate_End + "') > '" + vPayDate_End + "' then '" + vPayDate_End + "' else isnull(e.LeaveDay, '" + vPayDate_End + "') end) + 1 as WorkDays, " + Environment.NewLine +
                                "       p.GivCash " + Environment.NewLine +
                                "  from PayRec p left join Employee e on e.EmpNo = p.EmpNo " + Environment.NewLine +
                                " where PayDate between '" + vPayDate_Start + "' and '" + vPayDate_End + "' and p.PayDur = '" + vPayDur + "' " + Environment.NewLine +
                                "   and (e.AssumeDay between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "        or e.LeaveDay between '" + vPayDate_Start + "' and '" + vPayDate_End + "' " + Environment.NewLine +
                                "        or (e.BeginDay <= '" + vPayDate_End + "' and isnull(e.StopDay, '" + vPayDate_End + "') >= '" + vPayDate_Start + "')) " + Environment.NewLine +
                                //"        or  isnull(e.WorkType, '') <> '在職') " + Environment.NewLine +
                                " order by p.DepNo, p.EmpNo ";
            return vResultStr;
        }

        private void OpenDataList()
        {
            sdsEmpPayList.SelectCommand = GetSelectStr();
            gridEmpPayList.DataBind();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenDataList();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vReportName = ePayYear_Search.Text.Trim() + "年" + ePayMonth_Search.Text.Trim() + "月到離員工應發薪資清冊";
            //統計報表
            string vSelectStr = GetSelectStr();
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

                ReportDataSource rdsPrint = new ReportDataSource("EmpPayListP", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\EmpPayListP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("CalYM", ePayYear_Search.Text.Trim() + "年" + ePayMonth_Search.Text.Trim() + "月"));
                rvPrint.LocalReport.Refresh();
                plShowData.Visible = false;
                plReport.Visible = true;

                string vPayYMStr = ePayYear_Search.Text.Trim() + " 年 " + ePayMonth_Search.Text.Trim() + " 月";
                string vPayDurStr = "第 " + ePayDur_Search.Text.Trim() + " 期";
                string vRecordNote = "預覽報表_到離員工應發薪資清冊" + Environment.NewLine +
                                     "EmpPayList.aspx" + Environment.NewLine +
                                     "發薪年月：" + vPayYMStr + Environment.NewLine +
                                     "期別：" + vPayDurStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
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

            HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_Float.DataFormat = format.GetFormat("###,##0.00%");

            string vHeaderText = "";
            string vColumnName = "";
            string vFileName = ePayYear_Search.Text.Trim() + "年" + ePayMonth_Search.Text.Trim() + "月到離員工應發薪資清冊";
            int vLinesNo = 0;
            DateTime vTempDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelectStr();
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
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "單位" :
                                      (drExcel.GetName(i).ToUpper() == "EMPNO") ? "工號" :
                                      (drExcel.GetName(i).ToUpper() == "NAME") ? "姓名" :
                                      (drExcel.GetName(i).ToUpper() == "WORKTYPE") ? "狀態" :
                                      (drExcel.GetName(i).ToUpper() == "ASSUMEDAY") ? "到職日" :
                                      (drExcel.GetName(i).ToUpper() == "LEAVEDAY") ? "離職日" :
                                      (drExcel.GetName(i).ToUpper() == "BEGINDAY") ? "留停起日" :
                                      (drExcel.GetName(i).ToUpper() == "STOPDAY") ? "留停迄日" :
                                      (drExcel.GetName(i).ToUpper() == "WORKDAYS") ? "上班天數" :
                                      (drExcel.GetName(i).ToUpper() == "GIVCASH") ? "應發金額" : drExcel.GetName(i);
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
                            vColumnName = drExcel.GetName(i);
                            if (drExcel[i].ToString() != "" && ((drExcel.GetName(i).ToUpper() == "GIVCASH") || (drExcel.GetName(i).ToUpper() == "WORKDAYS")))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else if (drExcel[i].ToString() != "" && ((drExcel.GetName(i).ToUpper() == "ASSUMEDAY") || (drExcel.GetName(i).ToUpper() == "LEAVEDAY")))
                            {
                                vTempDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vTempDate.Year.ToString("D4") + "/" + vTempDate.ToString("MM/dd"));
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
                            string vPayYMStr = ePayYear_Search.Text.Trim() + " 年 " + ePayMonth_Search.Text.Trim() + " 月";
                            string vPayDurStr = "第 " + ePayDur_Search.Text.Trim() + " 期";
                            string vRecordNote = "匯出檔案_到離員工應發薪資清冊" + Environment.NewLine +
                                                 "EmpPayList.aspx" + Environment.NewLine +
                                                 "發薪年月：" + vPayYMStr + Environment.NewLine +
                                                 "期別：" + vPayDurStr;
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