using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI;

namespace TyBus_Intranet_Test_V3
{
    public partial class PurchaseToExcel : System.Web.UI.Page
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
                        cbIncludeOil_Search.Checked = false;
                        cbIncludeTire_Search.Checked = false;
                        eBuDateE_Search.Text = "";
                        eBuDateS_Search.Text = "";
                    }
                    else
                    {

                    }
                    string vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDateS_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDateS_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDateE_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDateE_Search.Attributes["onClick"] = vDateScript;
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

        private String GetSelectStr()
        {
            string vDateStr_S = (eBuDateS_Search.Text.Trim() != "") ? PF.GetAD(eBuDateS_Search.Text.Trim()) : "";
            string vDateStr_E = (eBuDateE_Search.Text.Trim() != "") ? PF.GetAD(eBuDateE_Search.Text.Trim()) : "";
            string vWStr_SheetDiff = ((cbIncludeOil_Search.Checked == false) && (cbIncludeTire_Search.Checked == false)) ? " where a.SheetDiff = 'M' " + Environment.NewLine :
                                     ((cbIncludeOil_Search.Checked == true) && (cbIncludeTire_Search.Checked == false)) ? " where a.SheetDiff != 'T' " + Environment.NewLine :
                                     ((cbIncludeOil_Search.Checked == false) && (cbIncludeTire_Search.Checked == true)) ? " where a.SheetDiff != 'O' " + Environment.NewLine :
                                     " where isnull(SheetDiff, '') != '' " + Environment.NewLine;
            string vWStr_BuDate = ((vDateStr_S != "") && (vDateStr_E != "")) ? "   and a.BuDate between '" + vDateStr_S + "' and '" + vDateStr_E + "' " + Environment.NewLine :
                                  ((vDateStr_S != "") && (vDateStr_E == "")) ? "   and a.BuDate = '" + vDateStr_S + "' " + Environment.NewLine :
                                  ((vDateStr_S == "") && (vDateStr_E != "")) ? "   and a.BuDate = '" + vDateStr_E + "' " + Environment.NewLine : "";
            string vSelectStr = "select distinct a.PurClass, c.ClassTxt PurClass_C, a.StoreNo, d.[Name] StoreName " + Environment.NewLine +
                                "  from PurchaseA a left join DBDICB c on c.ClassNo = a.PurClass and c.FKey = '材料請購單A     PURCHASEA       PURCLASS' " + Environment.NewLine +
                                "                   left join [Custom] d on d.Code = a.StoreNo " + Environment.NewLine +
                                vWStr_SheetDiff + vWStr_BuDate +
                                " order by a.PurClass, a.StoreNo";
            return vSelectStr;
        }

        private string GetDetailSelectStr(string fStoreNo)
        {
            string vDateStr_S = (eBuDateS_Search.Text.Trim() != "") ? PF.GetAD(eBuDateS_Search.Text.Trim()) : "";
            string vDateStr_E = (eBuDateE_Search.Text.Trim() != "") ? PF.GetAD(eBuDateE_Search.Text.Trim()) : "";
            string vWStr_SheetDiff = ((cbIncludeOil_Search.Checked == false) && (cbIncludeTire_Search.Checked == false)) ? "   and a.SheetDiff = 'M' " + Environment.NewLine :
                                     ((cbIncludeOil_Search.Checked == true) && (cbIncludeTire_Search.Checked == false)) ? "   and a.SheetDiff != 'T' " + Environment.NewLine :
                                     ((cbIncludeOil_Search.Checked == false) && (cbIncludeTire_Search.Checked == true)) ? "   and a.SheetDiff != 'O' " + Environment.NewLine : "";
            string vWStr_BuDate = ((vDateStr_S != "") && (vDateStr_E != "")) ? "   and a.BuDate between '" + vDateStr_S + "' and '" + vDateStr_E + "' " + Environment.NewLine :
                                  ((vDateStr_S != "") && (vDateStr_E == "")) ? "   and a.BuDate = '" + vDateStr_S + "' " + Environment.NewLine :
                                  ((vDateStr_S == "") && (vDateStr_E != "")) ? "   and a.BuDate = '" + vDateStr_E + "' " + Environment.NewLine : "";
            //2023.09.27 使用者表示不需要最後的 "預交貨日" 欄位
            //string vSelectStr = "select b.PurNo, b.[No], b.StockName, b.Quantity, b.Price, b.Amount, b.TradeDate " + Environment.NewLine +
            string vSelectStr = "select b.PurNo, b.[No], b.StockName, b.Quantity, b.Price, b.Amount " + Environment.NewLine +
            "  from PurchaseA a left join PurchaseB b on b.PurNo = a.PurNo " + Environment.NewLine +
                                " where a.StoreNo = '" + fStoreNo + "' " + Environment.NewLine +
                                vWStr_SheetDiff + vWStr_BuDate +
                                " order by a.PurClass, a.StoreNo, b.PurNo, b.[Item]";
            return vSelectStr;
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
            string vFileName = eBuDateS_Search.Text.Trim() + "起至" + eBuDateE_Search.Text.Trim() + "止請購單明細";
            string vStoreNo = "";
            string vWorkSheetName = "";
            int vLinesNo = 0;
            DateTime vTempDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelectStr();
            string vSubSelectStr = "";
            using (SqlConnection connPurchaseA = new SqlConnection(vConnStr))
            {
                SqlCommand cmdPurchaseA = new SqlCommand(vSelectStr, connPurchaseA);
                connPurchaseA.Open();
                SqlDataReader drPurchaseA = cmdPurchaseA.ExecuteReader();
                if (drPurchaseA.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    while (drPurchaseA.Read())
                    {
                        vStoreNo = drPurchaseA["StoreNo"].ToString().Trim();
                        vWorkSheetName = drPurchaseA["StoreName"].ToString().Trim();
                        //新增一個工作表
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(vWorkSheetName);
                        //新增抬頭列
                        vLinesNo = 0;
                        wsExcel.CreateRow(vLinesNo);
                        wsExcel.GetRow(vLinesNo).CreateCell(0, CellType.String);
                        wsExcel.GetRow(vLinesNo).GetCell(0).SetCellValue("廠商：[" + drPurchaseA["StoreNo"].ToString().Trim() + "]" + drPurchaseA["StoreName"].ToString().Trim());
                        wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 5));
                        wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                        vLinesNo++;

                        //取回本頁資料
                        vSubSelectStr = GetDetailSelectStr(vStoreNo);
                        using (SqlConnection connExcel = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdExcel = new SqlCommand(vSubSelectStr, connExcel);
                            connExcel.Open();
                            SqlDataReader drExcel = cmdExcel.ExecuteReader();
                            if (drExcel.HasRows)
                            {
                                wsExcel.CreateRow(vLinesNo);
                                for (int i = 0; i < drExcel.FieldCount; i++)
                                {
                                    vHeaderText = (drExcel.GetName(i).ToUpper() == "PURNO") ? "請購單號" :
                                                  (drExcel.GetName(i).ToUpper() == "NO") ? "料號" :
                                                  (drExcel.GetName(i).ToUpper() == "STOCKNAME") ? "品名" :
                                                  (drExcel.GetName(i).ToUpper() == "QUANTITY") ? "數量" :
                                                  (drExcel.GetName(i).ToUpper() == "PRICE") ? "單價" :
                                                  (drExcel.GetName(i).ToUpper() == "AMOUNT") ? "小計" : "";
                                    //2023.09.27 使用者表示不需要最後的 "預交貨日" 欄位
                                    //(drExcel.GetName(i).ToUpper() == "AMOUNT") ? "小計" :
                                    //(drExcel.GetName(i).ToUpper() == "TRADEDATE") ? "預到貨日" : "";
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
                                        vColumnName = drExcel.GetName(i);
                                        if (drExcel[i].ToString() != "" &&
                                            ((drExcel.GetName(i).ToUpper() == "QUANTITY") ||
                                             (drExcel.GetName(i).ToUpper() == "PRICE") ||
                                             (drExcel.GetName(i).ToUpper() == "AMOUNT")))
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                        }
                                        //2023.09.27 使用者表示不需要最後的 "預交貨日" 欄位
                                        //else if (drExcel[i].ToString() != "" && drExcel.GetName(i).ToUpper() == "TRADEDATE")
                                        //{
                                        //    vTempDate = DateTime.Parse(drExcel[i].ToString());
                                        //    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                        //    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vTempDate.Year.ToString("D4") + "/" + vTempDate.ToString("MM/dd"));
                                        //    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                        //}
                                        else
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                        }
                                    }
                                }
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                wsExcel.GetRow(vLinesNo).CreateCell(4, CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(4).SetCellValue("合計：");
                                wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData;
                                wsExcel.GetRow(vLinesNo).CreateCell(5, CellType.Formula);
                                wsExcel.GetRow(vLinesNo).GetCell(5).SetCellFormula("SUM(F3:F" + vLinesNo.ToString() + ")");
                                wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData_Int;
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
                            string vOther = ((cbIncludeOil_Search.Checked == false) && (cbIncludeTire_Search.Checked == false)) ? "只顯示材料請購單" + Environment.NewLine :
                                            ((cbIncludeOil_Search.Checked == true) && (cbIncludeTire_Search.Checked == false)) ? "包括材料及輪胎請購單" + Environment.NewLine :
                                            ((cbIncludeOil_Search.Checked == false) && (cbIncludeTire_Search.Checked == true)) ? "包括材料及油料請購單" + Environment.NewLine : "全部請購單";
                            string vRecordNote = "匯出檔案_[" + eBuDateS_Search.Text.Trim() + "] 起至 [" + eBuDateE_Search.Text.Trim() + "] 止請購單明細" + Environment.NewLine +
                                                 "PurchaseToExcel.aspx" + Environment.NewLine +
                                                 "日期範圍：[" + eBuDateS_Search.Text.Trim() + "] 起至 [" + eBuDateE_Search.Text.Trim() + "] 止" + Environment.NewLine +
                                                 "其他條件：" + vOther;
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
                    Response.Write("alert('查無資料！')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}