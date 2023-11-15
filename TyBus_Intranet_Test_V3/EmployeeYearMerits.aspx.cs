using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class EmployeeYearMerits : System.Web.UI.Page
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
        //DateTime vToday = DateTime.Today;

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
                        eCheckYear.Text = (DateTime.Today.Year - 1).ToString();
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

        private string GetSelectStr_Excel(string fYear)
        {
            string vResultStr = "SELECT DEPNO, (SELECT [NAME] FROM DEPARTMENT WHERE (DEPNO = z.DEPNO)) AS DepName, " + Environment.NewLine +
                                "       TITLE, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = z.TITLE)) AS Title_C, " + Environment.NewLine +
                                "       ASSUMEDAY, EMPNO, [NAME], SalaryPoint, " + Environment.NewLine +
                                "       SUM(TruNum) AS TruNum, SUM(EscType4) AS EscType4, SUM(EscType5) AS EscType5, SUM(EscType2) AS EscType2, " + Environment.NewLine +
                                "       SUM(Warning) AS Warning, SUM(Demerit3) AS Demerit3, SUM(Demerit2) AS Demerit2, SUM(Demerit1) AS Demerit1, " + Environment.NewLine +
                                "       SUM(Merit3) AS Merit3, SUM(Merit2) AS Merit2, SUM(Merit1) AS Merit1 " + Environment.NewLine +
                                "  FROM ( " + Environment.NewLine +
                                "       SELECT a.DEPNO, a.TITLE, a.ASSUMEDAY, a.EMPNO, a.[NAME], a.salarypoint, " + Environment.NewLine +
                                "              ISNULL(b.trunum, 0) AS TruNum, ISNULL(b.esctype4, 0) AS EscType4, ISNULL(b.esctype5, 0) AS EscType5, ISNULL(b.esctype2, 0) AS EscType2, " + Environment.NewLine +
                                "              CAST(0 AS float) AS Warning, CAST(0 AS float) AS Demerit3, CAST(0 AS float) AS Demerit2, CAST(0 AS float) AS Demerit1, " + Environment.NewLine +
                                "              CAST(0 AS float) AS Merit3, CAST(0 AS float) AS Merit2, CAST(0 AS float) AS Merit1 " + Environment.NewLine +
                                "         FROM EMPLOYEE AS a LEFT OUTER JOIN MHGS AS b ON b.empno = a.EMPNO " + Environment.NewLine +
                                "        WHERE (a.DEPNO <> '00') AND (a.LEAVEDAY IS NULL) AND (YEAR(b.attmonth) = '" + fYear.Trim() + "') " + Environment.NewLine +
                                "        UNION ALL " + Environment.NewLine +
                                "       SELECT a.DEPNO, a.TITLE, a.ASSUMEDAY, a.EMPNO, a.[NAME], a.SalaryPoint, " + Environment.NewLine +
                                "              CAST(0 AS float) AS TruNum, CAST(0 AS float) AS EscType4, CAST(0 AS float) AS EscType5, CAST(0 AS float) AS EscType2, " + Environment.NewLine +
                                "              ISNULL(c.WARNING, 0) AS Warning, ISNULL(c.DEMERIT3, 0) AS Demerit3, ISNULL(c.DEMERIT2, 0) AS Demerit2, ISNULL(c.DEMERIT1, 0) AS Demerit1, " + Environment.NewLine +
                                "              ISNULL(c.MERIT3, 0) AS Merit3, ISNULL(c.MERIT2, 0) AS Merit2, ISNULL(c.MERIT1, 0) AS Merit1 " + Environment.NewLine +
                                "         FROM EMPLOYEE AS a LEFT OUTER JOIN MERITSA AS c ON c.EMPNO = a.EMPNO " + Environment.NewLine +
                                "        WHERE (a.DEPNO <> '00') AND (a.LEAVEDAY IS NULL) AND (YEAR(c.MERITYEAR) = '" + fYear.Trim() + "') " + Environment.NewLine +
                                "       ) AS z " + Environment.NewLine +
                                " GROUP BY DEPNO, TITLE, ASSUMEDAY, EMPNO, NAME, salarypoint " + Environment.NewLine +
                                " ORDER BY   DEPNO, TITLE ";
            return vResultStr;
        }

        private void ExportExcel(string fSelectStr)
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

            string vHeaderText = "";
            int vLinesNo = 0;
            string vFileName = "桃園客運_" + eCheckYear.Text.Trim() + "年度員工考核參考資料";
            string vFieldValue = "";
            DateTime vDate;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(fSelectStr, connExcel);
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
                                      (drExcel.GetName(i).ToUpper() == "TITLE") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "TITLE_C") ? "職稱" :
                                      (drExcel.GetName(i).ToUpper() == "EMPNO") ? "代號" :
                                      (drExcel.GetName(i).ToUpper() == "NAME") ? "姓名" :
                                      (drExcel.GetName(i).ToUpper() == "ASSUMEDAY") ? "到職日" :
                                      (drExcel.GetName(i).ToUpper() == "SALARYPOINT") ? "俸點" :
                                      (drExcel.GetName(i).ToUpper() == "TRUNUM") ? "曠職" :
                                      (drExcel.GetName(i).ToUpper() == "ESCTYPE4") ? "事假" :
                                      (drExcel.GetName(i).ToUpper() == "ESCTYPE5") ? "病假" :
                                      (drExcel.GetName(i).ToUpper() == "ESCTYPE2") ? "公假" :
                                      (drExcel.GetName(i).ToUpper() == "WARNING") ? "警告" :
                                      (drExcel.GetName(i).ToUpper() == "DEMERIT3") ? "申誡" :
                                      (drExcel.GetName(i).ToUpper() == "DEMERIT2") ? "記過" :
                                      (drExcel.GetName(i).ToUpper() == "DEMERIT1") ? "記大過" :
                                      (drExcel.GetName(i).ToUpper() == "MERIT3") ? "嘉獎" :
                                      (drExcel.GetName(i).ToUpper() == "MERIT2") ? "記功" :
                                      (drExcel.GetName(i).ToUpper() == "MERIT1") ? "記大功" : drExcel.GetName(i);
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
                            if (((vHeaderText == "TRUNUM") ||
                                 (vHeaderText == "SALARYPOINT") ||
                                 (vHeaderText == "ESCTYPE4") ||
                                 (vHeaderText == "ESCTYPE5") ||
                                 (vHeaderText == "ESCTYPE2") ||
                                 (vHeaderText == "WARNING") ||
                                 (vHeaderText == "DEMERIT3") ||
                                 (vHeaderText == "DEMERIT2") ||
                                 (vHeaderText == "DEMERIT1") ||
                                 (vHeaderText == "MERIT3") ||
                                 (vHeaderText == "MERIT2") ||
                                 (vHeaderText == "MERIT1")) && (drExcel[i].ToString() != ""))
                            {
                                //數值欄位
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                vFieldValue = (drExcel[i].ToString().Trim() != "") ? drExcel[i].ToString().Trim() : "0";
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(vFieldValue));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else if ((drExcel.GetName(i).ToUpper() == "ASSUMEDAY") && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vDate.Year.ToString("D4") + "/" + vDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
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
                            string vCheckTearStr = eCheckYear.Text.Trim();
                            string vRecordNote = "匯出檔案_員工考核參考資料查詢" + Environment.NewLine +
                                                 "EmployeeYearMerits.aspx" + Environment.NewLine +
                                                 "查詢年月：" + vCheckTearStr;
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

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            if (eCheckYear.Text.Trim() != "")
            {
                sdsDataList.SelectParameters.Clear();
                string vCheckYear = (Int32.Parse(eCheckYear.Text.Trim()) < 1911) ? (Int32.Parse(eCheckYear.Text.Trim()) + 1911).ToString() : eCheckYear.Text.Trim();
                sdsDataList.SelectParameters.Add(new Parameter("Year", System.Data.DbType.String, vCheckYear));
                gridDataList.DataBind();
                string vCheckTearStr = eCheckYear.Text.Trim();
                string vRecordNote = "查詢資料_員工考核參考資料查詢" + Environment.NewLine +
                                     "EmployeeYearMerits.aspx" + Environment.NewLine +
                                     "查詢年月：" + vCheckTearStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請輸入查詢年度')");
                Response.Write("</" + "Script>");
                eCheckYear.Focus();
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (eCheckYear.Text.Trim() != "")
            {
                string vCheckYear = (Int32.Parse(eCheckYear.Text.Trim()) < 1911) ? (Int32.Parse(eCheckYear.Text.Trim()) + 1911).ToString() : eCheckYear.Text.Trim();
                string vSelectStr = GetSelectStr_Excel(vCheckYear);
                ExportExcel(vSelectStr);
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請輸入查詢年度')");
                Response.Write("</" + "Script>");
                eCheckYear.Focus();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}