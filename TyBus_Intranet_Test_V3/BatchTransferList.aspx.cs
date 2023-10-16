using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.IO;
using System.Text;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class BatchTransferList : System.Web.UI.Page
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
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
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

        private void SaveToExcel(string fPayAccount, string[] fTempArray)
        {
            string vLinesIndex = "";
            string vTempStr = "";
            string vCompanyName = "";

            string vPayDate;
            string vPayDate_Temp;
            string vPaymentStr = "";
            double vPayment = 0.0;
            double vTotalCount = 0;
            double vTotalPayment = 0;

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

            int vLinesNo = 0;
            wsExcel = (HSSFSheet)wbExcel.CreateSheet("付款明細");
            for (int i = 0; i < fTempArray.Length; i++)
            {
                vTempStr = fTempArray[i];
                vLinesIndex = (vTempStr.Trim() != "") ? vTempStr.Substring(0, 1) : "0";
                if (vLinesIndex == "1") //文字檔第一行
                {
                    wsExcel.CreateRow(vLinesNo);
                    vCompanyName = vTempStr.Substring(48, 68).Trim();
                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue(vCompanyName + "付款明細表");
                    //wsExcel.AutoSizeColumn(1);
                    wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 1, 6));
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("單據類別代號：380:發票, 493:對帳單, 705:提單, 740:空運提單,");
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("自動計算筆數：");
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("日期：");
                    wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("單據類別代號：801:訂單, 802:驗貨單, 803:請款單, 900:其他)");
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("自動計算金額：");
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("製表：");
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue((string)Session["LoginName"]);
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0).SetCellValue("付款帳號");
                    wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("付款日期");
                    wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellValue("付款金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellValue("收款人戶名");
                    wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellValue("收款行代號");
                    wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellValue("收款帳號");
                    wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellValue("收款人統編");
                    wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellValue("付款摘要");
                    wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellValue("通知附言");
                    wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellValue("通知傳真號碼");
                    wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellValue("通知-Email");
                    wsExcel.GetRow(vLinesNo).CreateCell(11).SetCellValue("單據類別");
                    wsExcel.GetRow(vLinesNo).CreateCell(12).SetCellValue("單據號碼");
                    wsExcel.GetRow(vLinesNo).CreateCell(13).SetCellValue("單據金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(14).SetCellValue("實付金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(15).SetCellValue("單據日期");
                    wsExcel.GetRow(vLinesNo).CreateCell(16).SetCellValue("備註");
                    wsExcel.GetRow(vLinesNo).CreateCell(17).SetCellValue("單據類別");
                    wsExcel.GetRow(vLinesNo).CreateCell(18).SetCellValue("單據號碼");
                    wsExcel.GetRow(vLinesNo).CreateCell(19).SetCellValue("單據金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(20).SetCellValue("實付金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(21).SetCellValue("單據日期");
                    wsExcel.GetRow(vLinesNo).CreateCell(22).SetCellValue("備註");
                    wsExcel.GetRow(vLinesNo).CreateCell(23).SetCellValue("單據類別");
                    wsExcel.GetRow(vLinesNo).CreateCell(24).SetCellValue("單據號碼");
                    wsExcel.GetRow(vLinesNo).CreateCell(25).SetCellValue("單據金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(26).SetCellValue("實付金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(27).SetCellValue("單據日期");
                    wsExcel.GetRow(vLinesNo).CreateCell(28).SetCellValue("備註");
                    wsExcel.GetRow(vLinesNo).CreateCell(29).SetCellValue("單據類別");
                    wsExcel.GetRow(vLinesNo).CreateCell(30).SetCellValue("單據號碼");
                    wsExcel.GetRow(vLinesNo).CreateCell(31).SetCellValue("單據金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(32).SetCellValue("實付金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(33).SetCellValue("單據日期");
                    wsExcel.GetRow(vLinesNo).CreateCell(34).SetCellValue("備註");
                    wsExcel.GetRow(vLinesNo).CreateCell(35).SetCellValue("單據類別");
                    wsExcel.GetRow(vLinesNo).CreateCell(36).SetCellValue("單據號碼");
                    wsExcel.GetRow(vLinesNo).CreateCell(37).SetCellValue("單據金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(38).SetCellValue("實付金額");
                    wsExcel.GetRow(vLinesNo).CreateCell(39).SetCellValue("單據日期");
                    wsExcel.GetRow(vLinesNo).CreateCell(40).SetCellValue("備註");
                    vLinesNo++;
                }
                else if (vLinesIndex == "2") //文字檔明細項
                {
                    wsExcel.CreateRow(vLinesNo);
                    //付款帳號
                    wsExcel.GetRow(vLinesNo).CreateCell(0, CellType.String).SetCellValue(fPayAccount);
                    //付款日期
                    vPayDate_Temp = "1" + vTempStr.Substring(19, 6);
                    vPayDate = (Int32.Parse(vPayDate_Temp.Substring(0, 3)) + 1911).ToString("D4") + "/" + (Int32.Parse(vPayDate_Temp.Substring(2, 2))).ToString("D2") + "/" + (Int32.Parse(vPayDate_Temp.Substring(5, 2))).ToString("D2");
                    wsExcel.GetRow(vLinesNo).CreateCell(1, CellType.String).SetCellValue(vPayDate);
                    //付款金額
                    vPaymentStr = vTempStr.Substring(116, 14).Trim();
                    vPayment = (vPaymentStr != "") ? double.Parse(vPaymentStr) / 100.0 : 0;
                    wsExcel.GetRow(vLinesNo).CreateCell(2, CellType.Numeric).SetCellValue(vPayment);
                    //收款人戶名
                    wsExcel.GetRow(vLinesNo).CreateCell(3, CellType.String).SetCellValue(vTempStr.Substring(48, 68));
                    //收款行代號
                    wsExcel.GetRow(vLinesNo).CreateCell(4, CellType.String).SetCellValue(vTempStr.Substring(25, 7));
                    //收款帳號
                    wsExcel.GetRow(vLinesNo).CreateCell(5, CellType.String).SetCellValue(vTempStr.Substring(34, 14));
                    vLinesNo++;
                }
                else if (vLinesIndex == "3") //文字檔尾行
                {
                    vTotalCount = double.Parse(vTempStr.Substring(25, 4));
                    vTotalPayment = double.Parse(vTempStr.Substring(29, 14)) / 100.0;
                    wsExcel.GetRow(2).CreateCell(3, CellType.Numeric).SetCellValue(vTotalCount);
                    wsExcel.GetRow(3).CreateCell(3, CellType.Numeric).SetCellValue(vTotalPayment);
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
                    string vFileName = "UP1001";

                    if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                    {
                        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", Encoding.UTF8));
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
                Response.Write("alert('" + eMessage.Message + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbTrans_Click(object sender, EventArgs e)
        {
            if ((ePayAccount.Text.Trim() != "") && (fuSourceTXT.HasFile))
            {
                string tempStr = "";
                string[] tempTXT;
                string vPayAccount = ePayAccount.Text.Trim();
                string vFileName = fuSourceTXT.FileName;
                string vAppPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vFilePath = vAppPath + vFileName;
                fuSourceTXT.SaveAs(vFilePath);
                if (File.Exists(vFilePath))
                {
                    StreamReader objStreamReader = new StreamReader(vFilePath, Encoding.Default);
                    //將字串內容從頭到尾讀取至緩衝區。
                    string strText = objStreamReader.ReadToEnd();
                    //釋放使用中的資源。
                    objStreamReader.Close();
                    objStreamReader.Dispose();
                    tempStr = strText.Replace(Environment.NewLine, ";");
                    tempTXT = tempStr.Split(';');
                    SaveToExcel(vPayAccount, tempTXT);
                }
            }
            else
            {
                string vErrorMSG = ((ePayAccount.Text.Trim() == "") && (fuSourceTXT.HasFile)) ? "請輸入付款帳號" :
                                   ((ePayAccount.Text.Trim() != "") && (!fuSourceTXT.HasFile)) ? "請先挑選來源 TXT 檔" :
                                   ((ePayAccount.Text.Trim() == "") && (!fuSourceTXT.HasFile)) ? "相關資料均未輸入！" : "";
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + vErrorMSG + "')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}