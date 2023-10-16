using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class StockMonth : System.Web.UI.Page
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

                if (vLoginID != "")
                {
                    if (!IsPostBack)
                    {
                        /*
                        eCalYear_Search.Text = (vToday.Month - 1 < 0) ? (vToday.Year - 1912).ToString() : (vToday.Year - 1911).ToString();
                        eCalMonth_Search.Text = (vToday.Month - 1 < 0) ? "12" : (vToday.Month - 1).ToString();
                        */
                        eCalYear_Search.Text = (vToday.AddMonths(-1).Year - 1911).ToString();
                        eCalMonth_Search.Text = vToday.AddMonths(-1).Month.ToString("D2");
                    }
                    StockDataBind(0);
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

        private string StockDataBind(int CallType)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vWStr_SClass = "";
            switch (rbStockType_Search.SelectedValue)
            {
                case "A":
                    vWStr_SClass = "";
                    break;
                case "M":
                    vWStr_SClass = "   and b.SClass = 'M' " + Environment.NewLine;
                    break;
                case "O":
                    vWStr_SClass = "   and b.SClass = 'O' " + Environment.NewLine;
                    break;
                case "T":
                    vWStr_SClass = "   and b.SClass = 'T' " + Environment.NewLine;
                    break;
                case "MO":
                    vWStr_SClass = "   and b.SClass in ('M', 'O') " + Environment.NewLine;
                    break;
                case "MT":
                    vWStr_SClass = "   and b.SClass in ('M', 'T') " + Environment.NewLine;
                    break;
                case "OT":
                    vWStr_SClass = "   and b.SClass in ('O', 'T') " + Environment.NewLine;
                    break;
                default:
                    vWStr_SClass = "";
                    break;
            }
            string vCalYM = (Int32.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString("D4") + "/" + Int32.Parse(eCalMonth_Search.Text.Trim()) + "/01";
            string vWStr_StockYM = "   and a.StockYM = '" + vCalYM + "' " + Environment.NewLine;

            string vSelectStr = "select a.StockNo, b.[Name], b.SClass, a.PreQty, a.PreUPC, a.Qty, a.UPC " + Environment.NewLine +
                                "  from MStock a left join Stock b on b.[No] = a.StockNo" + Environment.NewLine +
                                " where a.StockNo <> '98' " + Environment.NewLine +
                                vWStr_SClass + vWStr_StockYM +
                                " order by a.StockNo ";
            if (CallType == 0)
            {
                sdsMStockList.SelectCommand = "";
                sdsMStockList.SelectCommand = vSelectStr;
                gridMStockList.DataBind();
            }
            return vSelectStr;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            StockDataBind(0);
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

            string vSQLStr = StockDataBind(1);
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSQLStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    string vHeaderText = "";
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet("結存清冊");
                    //寫入標題列
                    int vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "StockNo") ? "料號" :
                                      (drExcel.GetName(i) == "Name") ? "品名" :
                                      (drExcel.GetName(i) == "SClass") ? "類別" :
                                      (drExcel.GetName(i) == "PreQty") ? "期初數量" :
                                      (drExcel.GetName(i) == "PreUPC") ? "期初金額" :
                                      (drExcel.GetName(i) == "Qty") ? "期末數量" :
                                      (drExcel.GetName(i) == "UPC") ? "期末金額" : drExcel.GetName(i);
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
                            if ((drExcel.GetName(i) == "PreQty") || (drExcel.GetName(i) == "PreUPC") || (drExcel.GetName(i) == "Qty") || (drExcel.GetName(i) == "UPC"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                            }
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                        }
                        vLinesNo++;
                    }
                    try
                    {
                        string vFileName = eCalYear_Search.Text.Trim() + "年" + eCalMonth_Search.Text.Trim() + "月材料結存清冊";
                        var msTarget = new NPOIMemoryStream();
                        msTarget.AllowClose = false;
                        wbExcel.Write(msTarget);
                        msTarget.Flush();
                        msTarget.Seek(0, SeekOrigin.Begin);
                        msTarget.AllowClose = true;

                        if (msTarget.Length > 0)
                        {
                            string vCalYMStr = eCalYear_Search.Text.Trim() + " 年 " + eCalMonth_Search.Text.Trim() + " 月";
                            string vStoreTypeStr = rbStockType_Search.SelectedItem.Text.Trim();
                            string vRecordNote = "匯出檔案_材料月結存清冊" + Environment.NewLine +
                                                 "StockMonth.aspx" + Environment.NewLine +
                                                 "結存年月：" + vCalYMStr + Environment.NewLine +
                                                 "材料類別：" + vStoreTypeStr;
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
                    Response.Write("alert('無符合條件資料！')");
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