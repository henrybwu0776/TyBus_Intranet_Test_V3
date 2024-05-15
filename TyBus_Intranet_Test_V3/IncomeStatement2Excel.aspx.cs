using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;

namespace TyBus_Intranet_Test_V3
{
    public partial class IncomeStatement2Excel : Page
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
                DateTime vToday = DateTime.Today;
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
                        eCalYear.Text = (vToday.Year - 1911).ToString();
                        eCalMonth.Text = vToday.Month.ToString();
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

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExport_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCalYear = (Int32.Parse(eCalYear.Text.Trim()) + 1911).ToString();
            string vCalMonth = Int32.Parse(eCalMonth.Text.Trim()).ToString();
            string vYearStartDate = vCalYear + "/01/01";
            string vStartDate = vCalYear + "/" + vCalMonth + "/01";
            string vEndDate = PF.GetMonthLastDay(DateTime.Parse(vStartDate), "B");
            string vSelectStr_AccountList = "select ClassGroup, ClassCode3, PageM, PCount2, ClassTitle " + Environment.NewLine +
                                            "  from AC_RptSetting " + Environment.NewLine +
                                            " where IsBlankLine = 0 " + Environment.NewLine +
                                            "   and IsCalField = 0 " + Environment.NewLine +
                                            "   and IsShowM = 1 " + Environment.NewLine +
                                            " order by ClassGroup, PageM, PCount2";
            using (SqlConnection connAccountList = new SqlConnection(vConnStr))
            {
                string vSQLStr_Temp = "";
                string vClassGroup = "";
                string vClassCode3 = "";
                string vClassTitle = "";
                int vDataLineNo = 0;
                int vMaxColCount = 0;
                int vPageM = 0;
                int vPCount2 = 0;
                double vAmountM = 0.0;
                double vAmountT = 0.0;
                SqlCommand cmdAccountList = new SqlCommand(vSelectStr_AccountList, connAccountList);
                connAccountList.Open();
                SqlDataReader drAccountList = cmdAccountList.ExecuteReader();

                if (drAccountList.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    vSQLStr_Temp = "select max(PCount2) as MaxPCount " + Environment.NewLine +
                                   "  from AC_RptSetting " + Environment.NewLine +
                                   " where IsShowM = 1 " + Environment.NewLine +
                                   "   and PageM = 1";
                    vMaxColCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "MaxPCount")); //取得第一頁最大的項次
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

                    HSSFDataFormat format_int = (HSSFDataFormat)wbExcel.CreateDataFormat();
                    csData_Int.DataFormat = format_int.GetFormat("###,##0");

                    HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData_Float.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                    HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
                    csData_Float.DataFormat = format_Float.GetFormat("###,##0.00");

                    //string vHeaderText = "";
                    int vLinesNo = 0;
                    int vColNo = 0;
                    string vFileName = eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月份月損益表";
                    string vCalField1 = "";
                    string vCalField2 = "";
                    string vCalField3 = "";
                    string vCalField4 = "";
                    string vCalMode1 = "";
                    string vCalMode2 = "";
                    string vCalMode3 = "";
                    string vFormula1 = "";
                    string vFormula2 = "";

                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("科目");
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("月計金額");
                    wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csTitle;
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("累計金額");
                    wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csTitle;
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("");
                    wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csTitle;
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("科目");
                    wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csTitle;
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("月計金額");
                    wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csTitle;
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("累計金額");
                    wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csTitle;
                    vLinesNo++;

                    while (drAccountList.Read())
                    {
                        vClassGroup = drAccountList[0].ToString().Trim();
                        vClassCode3 = drAccountList[1].ToString().Trim();
                        vPageM = Int32.TryParse(drAccountList[2].ToString().Trim(), out int vTempPageM) ? vTempPageM : 0;
                        vPCount2 = Int32.TryParse(drAccountList[3].ToString().Trim(), out int vTempPCount2) ? vTempPCount2 : 0;
                        vClassTitle = drAccountList[4].ToString().Trim();
                        string vGetAmountStr = "select sum(e.AmountM)as AmountM, sum(e.AmountT) as AmountT " + Environment.NewLine +
                                               "  from( " + Environment.NewLine +
                                               "        select 0.0 as AmountM, sum(isnull(a.debit, 0)) - sum(isnull(a.credit, 0)) as AmountT " + Environment.NewLine +
                                               "          from ACCOUNT a left join AC_CLASS c on SubString(a.Subject, 1, 3) = c.[No] " + Environment.NewLine +
                                               "                         left join AC_SUBJECT b on b.Subject = a.Subject " + Environment.NewLine +
                                               "         where a.Type <> '6' " + Environment.NewLine +
                                               "           and b.RptSetting like '" + vClassCode3 + "%' " + Environment.NewLine +
                                               "           and a.Rec_Date >= '" + vYearStartDate + " 00:00:00' and a.Rec_Date <= '" + vEndDate + " 23:59:59' " + Environment.NewLine +
                                               "           and a.VoidMan is null " + Environment.NewLine +
                                               "         Union all " + Environment.NewLine +
                                               "        select sum(isnull(a.debit, 0)) - sum(isnull(a.credit, 0)) as AmountM, 0.0 as AmountT " + Environment.NewLine +
                                               "          from ACCOUNT a left join AC_CLASS c on SubString(a.Subject, 1, 3) = c.[No] " + Environment.NewLine +
                                               "                         left join AC_SUBJECT b on b.Subject = a.Subject " + Environment.NewLine +
                                               "         where a.Type <> '6' " + Environment.NewLine +
                                               "           and b.RptSetting like '" + vClassCode3 + "%' " + Environment.NewLine +
                                               "           and a.Rec_Date >= '" + vStartDate + " 00:00:00' and a.Rec_Date <= '" + vEndDate + " 23:59:59' " + Environment.NewLine +
                                               "           and a.VoidMan is null " + Environment.NewLine +
                                               "        ) e ";
                        using (SqlConnection connAmount = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdAmount = new SqlCommand(vGetAmountStr, connAmount);
                            connAmount.Open();
                            SqlDataReader drAmount = cmdAmount.ExecuteReader();
                            while (drAmount.Read())
                            {
                                if ((vClassGroup.ToUpper() == "A") || (vClassGroup.ToUpper() == "B"))
                                {
                                    vAmountM = double.TryParse(drAmount[0].ToString().Trim(), out double vTempAmountM) ? vTempAmountM : 0.0;
                                    vAmountT = double.TryParse(drAmount[1].ToString().Trim(), out double vTempAmountT) ? vTempAmountT : 0.0;
                                }
                                else if ((vClassGroup.ToUpper() == "C") || (vClassGroup.ToUpper() == "D"))
                                {
                                    vAmountM = double.TryParse(drAmount[0].ToString().Trim(), out double vTempAmountM) ? vTempAmountM * -1 : 0.0;
                                    vAmountT = double.TryParse(drAmount[1].ToString().Trim(), out double vTempAmountT) ? vTempAmountT * -1 : 0.0;
                                }
                            }
                        }
                        vDataLineNo = (vPageM == 1) ? vPCount2 : vPCount2 + vMaxColCount;
                        if (vDataLineNo == 1)
                        {
                            vLinesNo = vDataLineNo;
                        }
                        while (vLinesNo < vDataLineNo)
                        {
                            if (wsExcel.GetRow(vLinesNo) == null)
                            {
                                wsExcel.CreateRow(vLinesNo);
                            }
                            vLinesNo++;
                        }
                        if ((vClassGroup.ToUpper() == "A") || (vClassGroup.ToUpper() == "B"))
                        {
                            wsExcel.CreateRow(vLinesNo);
                            vColNo = 0;
                        }
                        else
                        {
                            vColNo = 4;
                        }
                        wsExcel.GetRow(vLinesNo).CreateCell(vColNo).SetCellValue(vClassTitle);
                        wsExcel.GetRow(vLinesNo).GetCell(vColNo).CellStyle = csData;
                        wsExcel.GetRow(vLinesNo).CreateCell(vColNo + 1).SetCellValue(vAmountM);
                        wsExcel.GetRow(vLinesNo).GetCell(vColNo + 1).CellStyle = csData_Float;
                        wsExcel.GetRow(vLinesNo).CreateCell(vColNo + 2).SetCellValue(vAmountT);
                        wsExcel.GetRow(vLinesNo).GetCell(vColNo + 2).CellStyle = csData_Float;

                        vLinesNo++;
                    }

                    vSQLStr_Temp = "select ClassGroup, ClassCode3, PageM, PCount2, ClassTitle, CalField1, CalField2, CalField3, CalField4, CalMode1, CalMode2, CalMode3 " + Environment.NewLine +
                                   "  from AC_RptSetting " + Environment.NewLine +
                                   " where IsCalField = 1 " + Environment.NewLine +
                                   "   and IsShowM = 1 " + Environment.NewLine +
                                   " order by PageM,PCount2";
                    using (SqlConnection connCalField = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdCalField = new SqlCommand(vSQLStr_Temp, connCalField);
                        connCalField.Open();
                        SqlDataReader drCalField = cmdCalField.ExecuteReader();
                        while (drCalField.Read())
                        {
                            vClassGroup = drCalField[0].ToString().Trim();
                            vColNo = ((vClassGroup.ToUpper() == "A") || (vClassGroup.ToUpper() == "B")) ? 0 : 4;
                            vClassTitle = drCalField[4].ToString().Trim();
                            vPageM = Int32.TryParse(drCalField[2].ToString().Trim(), out int vTempPageM) ? vTempPageM : 1;
                            vDataLineNo = (Int32.TryParse(drCalField[3].ToString().Trim(), out int vTempPCount) && vPageM == 1) ? vTempPCount :
                                          (Int32.TryParse(drCalField[3].ToString().Trim(), out int vTempPCount2) && vPageM == 2) ? vTempPCount2 + vMaxColCount : 0;
                            vCalField1 = GetExcelLines(drCalField[5].ToString().Trim(), vMaxColCount);
                            vCalField2 = GetExcelLines(drCalField[6].ToString().Trim(), vMaxColCount);
                            vCalField3 = GetExcelLines(drCalField[7].ToString().Trim(), vMaxColCount);
                            vCalField4 = GetExcelLines(drCalField[8].ToString().Trim(), vMaxColCount);
                            vCalMode1 = drCalField[9].ToString().Trim();
                            vCalMode2 = drCalField[10].ToString().Trim();
                            vCalMode3 = drCalField[11].ToString().Trim();
                            vFormula1 = "IF(" + vCalField1 + vCalMode1 + vCalField2 + vCalMode2 + vCalField3 + vCalMode3 + vCalField4 + " > 0, " + vCalField1 + vCalMode1 + vCalField2 + vCalMode2 + vCalField3 + vCalMode3 + vCalField4 + ", 0.0)";
                            vFormula2 = vFormula1.Replace("B", "C");
                            vFormula2 = vFormula2.Replace("F", "G");
                            vFormula2 = vFormula2.Replace("IG", "IF");
                            while (vLinesNo < vDataLineNo)
                            {
                                if (wsExcel.GetRow(vLinesNo) == null)
                                {
                                    wsExcel.CreateRow(vLinesNo);
                                }
                                vLinesNo++;
                            }
                            if (wsExcel.GetRow(vDataLineNo) == null)
                            {
                                wsExcel.CreateRow(vDataLineNo);
                                vLinesNo = vDataLineNo;
                            }
                            wsExcel.GetRow(vDataLineNo).CreateCell(vColNo).SetCellValue(vClassTitle);
                            wsExcel.GetRow(vDataLineNo).GetCell(vColNo).CellStyle = csData;
                            wsExcel.GetRow(vDataLineNo).CreateCell(vColNo + 1).SetCellFormula(vFormula1);
                            wsExcel.GetRow(vDataLineNo).GetCell(vColNo + 1).CellStyle = csData;
                            wsExcel.GetRow(vDataLineNo).CreateCell(vColNo + 2).SetCellFormula(vFormula2);
                            wsExcel.GetRow(vDataLineNo).GetCell(vColNo + 2).CellStyle = csData;
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
                            string vCalBaseDateStr = eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月份";
                            string vRecordNote = "匯出檔案_" + eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月份月損益表" + Environment.NewLine +
                                                 "IncomeStatement2Excel.aspx" + Environment.NewLine +
                                                 "匯出年月：" + vCalBaseDateStr;
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
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('查無月損益表資料！')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        private string GetExcelLines(string fCodeItem, int fMaxCount)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vClassGroup = "";
            string vResultData = "";
            int vPageM = 0;
            string vSQLStr_Temp = "select PCount2, PageM, ClassGroup from AC_RPTSetting where CodeItem= '" + fCodeItem + "' ";

            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vClassGroup = ((drTemp[2].ToString().Trim() == "A") || (drTemp[2].ToString().Trim() == "B")) ? "B" : "F";
                    vPageM = Int32.TryParse(drTemp[1].ToString().Trim(), out int vTempINT) ? vTempINT : 1;
                    vResultData = (vClassGroup != "") ? vClassGroup + (fMaxCount * (vPageM - 1) + Int32.Parse(drTemp[0].ToString().Trim()) + 1).ToString() : "";
                }
            }

            return vResultData.ToString();
        }
    }
}