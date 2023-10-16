using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverBoundsList : System.Web.UI.Page
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
                Thread.CurrentThread.CurrentCulture = Cal;

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
                        string vSQLStr_Temp = "select InSHReport from Department where DepNo = '" + vLoginDepNo + "' ";
                        if (PF.GetValue(vConnStr, vSQLStr_Temp, "InSHReport").ToUpper().Trim() == "V")
                        {
                            eDepNo_Search.Text = vLoginDepNo.Trim();
                            eDepNo_Search.Enabled = false;
                        }
                        else
                        {
                            eDepNo_Search.Text = "";
                            eDepNo_Search.Enabled = true;
                        }
                        DateTime vCalDate = DateTime.Today.AddMonths(-1);
                        eCalYear_Search.Text = (vCalDate.Year - 1911).ToString();
                        eCalMonth_Search.Text = vCalDate.Month.ToString();
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

        private string GetSelectStr()
        {
            string vReturnStr = "";
            if ((eCalYear_Search.Text.Trim() != "") && (eCalMonth_Search.Text.Trim() != ""))
            {
                string vCalYear_Str = eCalYear_Search.Text.Trim();
                string vCalMonth_Str = eCalMonth_Search.Text.Trim();
                DateTime vStartDate = DateTime.Parse(vCalYear_Str + "/" + vCalMonth_Str + "/01");
                DateTime vEndDate = DateTime.Parse(vCalYear_Str + "/" + vCalMonth_Str + "/" + PF.GetMonthDays(vStartDate).ToString());
                string vWStr_BuDate = "           AND a.BuDate between '" + vStartDate.Year.ToString() + "/" + vStartDate.ToString("MM/dd") + "' and '" + vEndDate.Year.ToString() + "/" + vEndDate.ToString("MM/dd") + "' " + Environment.NewLine;
                string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and t.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
                vReturnStr = "SELECT CAST('" + vCalYear_Str + "年" + vCalMonth_Str + "月" + "' AS nvarchar(12)) AS CalYM, " + Environment.NewLine +
                             "       DEPNO, (SELECT [NAME] FROM DEPARTMENT WHERE (DEPNO = t.DEPNO)) AS DepName, " + Environment.NewLine +
                             "       DRIVER, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.DRIVER)) AS DriverName, " + Environment.NewLine +
                             "       SUM(RunItemBounds) AS TotalBounds_R, SUM(LinesBounds) AS TotalBounds_L, SUM(Car_Class) AS TotalBounds_C " + Environment.NewLine +
                             "  FROM ( " + Environment.NewLine +
                             "        SELECT a.DEPNO, a.DRIVER, " + Environment.NewLine +
                             "               CASE WHEN isnull((SELECT HasBounds FROM RunItem r WHERE DepNo = a.DepNo AND RunItemNo = b.RunItem), '') = 'V' THEN 1 ELSE 0 END AS RunItemBounds, " + Environment.NewLine +
                             "               CASE WHEN isnull((SELECT HasBounds FROM Lines l WHERE l.LinesNo = b.LinesNo), '') = 'V' THEN 1 ELSE 0 END AS LinesBounds, " + Environment.NewLine +
                             "               CASE WHEN isnull((SELECT Class FROM Car_InfoA WHERE Car_ID = b.Car_ID), '') = '甲' THEN 1 ELSE 0 END AS Car_Class " + Environment.NewLine +
                             "          FROM RunSheetB AS b LEFT OUTER JOIN RUNSHEETA AS a ON a.ASSIGNNO = b.AssignNo " + Environment.NewLine +
                             "         WHERE (ISNULL(b.ReduceReason, '') = '') " + Environment.NewLine +
                             vWStr_BuDate +
                             "       ) AS t " + Environment.NewLine +
                             " WHERE (1 = 1) " + Environment.NewLine +
                             vWStr_DepNo +
                             " GROUP BY DEPNO, DRIVER " + Environment.NewLine +
                             " ORDER BY DEPNO, DRIVER ";
            }
            return vReturnStr;
        }

        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            if (vSelStr != "")
            {
                sdsShowData.SelectCommand = vSelStr;
                gridShowData.DataBind();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先輸入查核年月！')");
                Response.Write("</" + "Script>");
                eCalYear_Search.Focus();
            }
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Search = eDepNo_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Search + "' ";
            string vDepName_Search = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Search == "")
            {
                vDepName_Search = vDepNo_Search.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Search + "' ";
                vDepNo_Search = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Search.Text = vDepNo_Search.Trim();
            eDepName_Search.Text = vDepName_Search.Trim();
        }

        protected void bbShowData_Click(object sender, EventArgs e)
        {
            OpenData();
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

            string vHeaderText = "";
            int vLinesNo = 0;
            string vSelectStr = GetSelectStr();
            string vFileName = "桃園客運_" + eCalYear_Search.Text.Trim() + "_年_" + eCalMonth_Search.Text.Trim() + "_月_駕駛員趟次清冊";
            string vFieldValue = "";

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
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "CALYM") ? "統計年月" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "場站名" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVER") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALBOUNDS_R") ? "AB班趟次" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALBOUNDS_L") ? "國道趟次" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALBOUNDS_C") ? "大巴趟次" : drExcel.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    vLinesNo++;
                    while (drExcel.Read())
                    {
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            vHeaderText = drExcel.GetName(i).ToUpper();
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if (((vHeaderText == "TOTALBOUNDS_R") ||
                                 (vHeaderText == "TOTALBOUNDS_L") ||
                                 (vHeaderText == "TOTALBOUNDS_C")) && (drExcel[i].ToString() != ""))
                            {
                                //數值欄位
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                vFieldValue = (drExcel[i].ToString().Trim() != "") ? drExcel[i].ToString().Trim() : "0";
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(vFieldValue));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else
                            {
                                //文字欄位
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
                            string vRecordDepNoStr = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_駕駛員大巴／國道／AB班趟次表" + Environment.NewLine +
                                                 "DriverBoundsList.aspx" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "統計年月：" + eCalYear_Search.Text.Trim() + "年" + eDepName_Search.Text.Trim() + "月";
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
    }
}