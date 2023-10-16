using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class ImportCash : System.Web.UI.Page
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
                //
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
                    if (!IsPostBack)
                    {
                        if (vConnStr == "")
                        {
                            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                        }
                        eBuDate.Text = DateTime.Today.ToShortDateString();
                        plShowData.Visible = false;
                        string vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDate.ClientID;
                        string vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eBuDate.Attributes["onClick"] = vTempDateScript;
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

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            string vExtName = Path.GetExtension(eFilePath.FileName);
            string vTempStr = "";
            DateTime vBuDate = DateTime.Parse(eBuDate.Text.Trim());
            string vDepNo_Temp = "";
            string vDriver = "";
            string vDriverName = "";
            string vCar_No = "";
            string vCar_ID = "";
            string vLinesNo = "";
            string vCashType = "";
            string vCash_1000_Str = "";
            string vCash_500_Str = "";
            string vCash_100_Str = "";
            string vCashAMT_Str = "";
            string vRemark = "";
            int vRCount = 0;
            if (vExtName == ".txt")
            {
                DataTable dtCoinData = new DataTable("CashList");
                dtCoinData.Columns.Add("BuDate", typeof(DateTime));
                dtCoinData.Columns.Add("DepNo", typeof(String));
                dtCoinData.Columns.Add("DepName", typeof(String));
                dtCoinData.Columns.Add("Driver", typeof(String));
                dtCoinData.Columns.Add("DriverName", typeof(String));
                dtCoinData.Columns.Add("Car_No", typeof(String));
                dtCoinData.Columns.Add("Car_ID", typeof(String));
                dtCoinData.Columns.Add("LinesNo", typeof(String));
                dtCoinData.Columns.Add("CashType", typeof(String));
                dtCoinData.Columns.Add("CashType_C", typeof(String));
                dtCoinData.Columns.Add("Cash_1000", typeof(Int32));
                dtCoinData.Columns.Add("Cash_500", typeof(Int32));
                dtCoinData.Columns.Add("Cash_100", typeof(Int32));
                dtCoinData.Columns.Add("CashAMT", typeof(Int32));
                dtCoinData.Columns.Add("BuMan", typeof(String));
                dtCoinData.Columns.Add("BuMan_C", typeof(String));
                dtCoinData.Columns.Add("Remark", typeof(String));

                string vUploadFileName = vUploadPath + eFilePath.FileName;
                eFilePath.SaveAs(vUploadFileName);
                using (StreamReader srTemp = new StreamReader(vUploadFileName, System.Text.Encoding.Default))
                {
                    vTempStr = srTemp.ReadLine();
                    while (vTempStr != null)
                    {
                        if (vConnStr == "")
                        {
                            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                        }
                        DataRow drTemp = dtCoinData.NewRow();
                        drTemp["BuDate"] = vBuDate;
                        string[] vSLTemp = vTempStr.Split(',');
                        vDepNo_Temp = vSLTemp[0].ToString().Trim();
                        drTemp["DepNo"] = vDepNo_Temp;
                        drTemp["DepName"] = PF.GetValue(vConnStr, "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ", "Name");
                        vDriver = vSLTemp[1].ToString().Trim();
                        drTemp["Driver"] = vDriver;
                        vDriverName = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vDriver + "' ", "Name");
                        if (vDriverName == "")
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('數幣機資料有誤，可能造成ERP系統無法轉入！')");
                            Response.Write("</" + "Script>");
                        }
                        drTemp["DriverName"] = vDriverName;
                        vRemark = (PF.GetValue(vConnStr, "select [Name] from Employee where DepNo = '" + vDepNo_Temp + "' and EmpNo = '" + vDriver + "' and LeaveDay is null ", "Name") == "") ?
                                  "駕駛員工號和單位不正確" : "";
                        vCar_No = vSLTemp[2].ToString().Trim();
                        drTemp["Car_No"] = vCar_No;
                        vCar_ID = PF.GetValue(vConnStr, "select Car_ID from Car_InfoA where Car_No = '" + vCar_No + "' ", "Car_ID");
                        if (vCar_ID == "")
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('數幣機資料有誤，可能造成ERP系統無法轉入！')");
                            Response.Write("</" + "Script>");
                        }
                        drTemp["Car_ID"] = vCar_ID;
                        //2019.06.17 課長指示不檢查人車不符的情況，只檢查該人該車所屬單位是否正確
                        //vRCount = Int32.Parse(PF.GetValue(vConnStr, "select count(Car_No) RCount from Car_InfoA where Car_No = '" + vCar_No + "' and CompanyNo = '" + vDepNo + "' and (Driver = '" + vDriver + "' or Driver1 = '" + vDriver + "')", "RCount"));
                        vRCount = Int32.Parse(PF.GetValue(vConnStr, "select count(Car_No) RCount from Car_InfoA where Car_No = '" + vCar_No + "' and CompanyNo = '" + vDepNo_Temp + "' ", "RCount"));
                        vRemark = (vRCount == 0) ? vRemark + " " + "車輛單位或駕駛員不正確" : vRemark;
                        vLinesNo = ((vSLTemp[3].ToString().Trim() != "") && (vSLTemp[3].ToString().Trim() != "00000")) ? vSLTemp[3].ToString().Trim() :
                                   (vSLTemp[3].ToString().Trim() == "00000") ? PF.GetValue(vConnStr, "select ClassTxt from DBDICB where ClassNo = '" + vDepNo_Temp + "' and Fkey = '投現收入轉檔       ImportCash      DepLinesNo' ", "ClassTxt") : "";
                        drTemp["LinesNo"] = vLinesNo;
                        //2019.06.18 加入路線代號檢查
                        vRCount = Int32.Parse(PF.GetValue(vConnStr, "select count(LinesNo) RCount from Lines where LinesNo = '" + vLinesNo + "' ", "RCount"));
                        vRemark = (vRCount == 0) ? vRemark + " " + "路線代號不正確" : vRemark;
                        vCashType = ((vSLTemp[3].ToString().Trim() != "") && (vSLTemp[3].ToString().Trim() != "00000")) ? "3" :
                                    (vSLTemp[3].ToString().Trim() == "00000") ? "1" : "";
                        drTemp["CashType"] = vCashType;
                        drTemp["CashType_C"] = PF.GetValue(vConnStr, "select ClassTxt from DBDICB where ClassNo = '" + vCashType + "' and FKey = '繳銷作業        CashBill        CASHTYPE'", "ClassTxt");
                        vCash_1000_Str = vSLTemp[4].ToString().Trim();
                        drTemp["Cash_1000"] = (vCash_1000_Str != "") ? Int32.Parse(vCash_1000_Str) : 0;
                        vCash_500_Str = vSLTemp[5].ToString().Trim();
                        drTemp["Cash_500"] = (vCash_500_Str != "") ? Int32.Parse(vCash_500_Str) : 0;
                        vCash_100_Str = vSLTemp[6].ToString().Trim();
                        drTemp["Cash_100"] = (vCash_100_Str != "") ? Int32.Parse(vCash_100_Str) : 0;
                        vCashAMT_Str = vSLTemp[7].ToString().Trim();
                        drTemp["CashAMT"] = (vCashAMT_Str != "") ? Int32.Parse(vCashAMT_Str) : 0;
                        drTemp["BuMan"] = vLoginID;
                        drTemp["BuMan_C"] = Session["LoginName"].ToString().Trim();
                        drTemp["Remark"] = vRemark;
                        dtCoinData.Rows.Add(drTemp);
                        vTempStr = srTemp.ReadLine();
                    }
                }
                gridShowData.DataSourceID = null;
                gridShowData.DataSource = dtCoinData;
                gridShowData.DataBind();
                plShowData.Visible = true;
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('只接受純文字檔上傳')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 資料匯入 ERP
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbImport_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vInsertStr = "";
            DateTime vBuDate = DateTime.Parse(eBuDate.Text.Trim());
            DateTime vRSDate = vBuDate.AddDays(-1);
            string vCBYM = (vBuDate.Year - 1911).ToString("D3") + vBuDate.Month.ToString("D2");
            string vMaxCBNo = PF.GetValue(vConnStr, "select MAX(CBNo) MaxNo from CashBill where CBNo like '" + vCBYM + "%' ", "MaxNo");
            int vIndex = (vMaxCBNo != "") ? Int32.Parse(vMaxCBNo.Substring(5)) : 0;
            string vIndexStr = "";
            string vCBNo = "";
            string vCashType = "";
            string vCashDepNo = "";
            string vDriver = "";
            string vLinesNoStr = "";
            string vCar_ID = "";
            double vCashAMT = 0;
            double vRootAMT = 0;
            double vTempAMT = 0;
            double vNormalQty = 0;
            double vOldQty = 0;
            double vTotalAMT = 0;
            string vBuMan = "";
            string vCar_No = "";
            int vCash_1000 = 0;
            int vCash_500 = 0;
            int vCash_100 = 0;
            string vRemark = "";
            string vTempStr = "";
            string vFileName = (vBuDate.Year - 1911).ToString("D3") + vBuDate.ToString("MMdd") + "數幣機輸入資料";
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
            //新增一個工作表
            wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
            //先產生標題列
            wsExcel.CreateRow(vLinesNo);
            for (int i = 0; i < gridShowData.Columns.Count; i++)
            {
                vHeaderText = gridShowData.Columns[i].HeaderText;
                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
            }
            for (int i = 0; i < gridShowData.Rows.Count; i++)
            {
                vRemark = gridShowData.Rows[i].Cells[16].Text.Trim();
                vRemark = vRemark.Replace("&nbsp;", "");
                if (vRemark != "")
                {
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int j = 0; j < gridShowData.Columns.Count; j++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(j);
                        vTempStr = (gridShowData.Rows[i].Cells[j].Text.Trim()).Replace("&nbsp;", "");
                        if (j == 0) //第一個欄位....BuDate 是日期
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(vTempStr);
                        }
                        else if ((j >= 10) && (j <= 13)) //中間四個欄位是數值
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(double.Parse(vTempStr));
                        }
                        else //其他欄位都是文字
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(vTempStr);
                        }
                        wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
                    }
                }
                else
                {
                    vIndexStr = (++vIndex).ToString("D4");
                    vCBNo = vCBYM + vIndexStr;
                    vCashType = gridShowData.Rows[i].Cells[8].Text.Trim();
                    vCashDepNo = gridShowData.Rows[i].Cells[1].Text.Trim();
                    vDriver = gridShowData.Rows[i].Cells[3].Text.Trim();
                    vLinesNoStr = gridShowData.Rows[i].Cells[7].Text.Trim();
                    vCar_ID = gridShowData.Rows[i].Cells[6].Text.Trim();
                    vCar_No = gridShowData.Rows[i].Cells[5].Text.Trim();
                    vCashAMT = double.Parse(gridShowData.Rows[i].Cells[13].Text.Trim());
                    vTotalAMT = vCashAMT + vRootAMT + vTempAMT;
                    vBuMan = vLoginID;
                    vCash_1000 = Int32.Parse(gridShowData.Rows[i].Cells[10].Text.Trim());
                    vCash_500 = Int32.Parse(gridShowData.Rows[i].Cells[11].Text.Trim());
                    vCash_100 = Int32.Parse(gridShowData.Rows[i].Cells[12].Text.Trim());
                    if (vInsertStr == "")
                    {
                        vInsertStr = "insert into CASHBILL (CBNo, BuDate, CashType, DepNo, Driver, LinesNo, Car_ID, " + Environment.NewLine +
                                     "             CashAMT, RootAMT, TempAMT, NormalQty, OldQty, TotalAMT, CBDate, BuMan, Car_No, " + Environment.NewLine +
                                     "             Cash_1000, Cash_500, Cash_100, RSDate)" + Environment.NewLine +
                                     "     values ('" + vCBNo + "', '" + vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd") + "', '" + vCashType + "', " + Environment.NewLine +
                                     "             '" + vCashDepNo + "', '" + vDriver + "', '" + vLinesNoStr + "', '" + vCar_ID + "', " + vCashAMT.ToString() + ", " + Environment.NewLine +
                                     "             " + vRootAMT.ToString() + ", " + vTempAMT.ToString() + ", " + vNormalQty.ToString() + ", " + vOldQty.ToString() + ", " + Environment.NewLine +
                                     "             " + vTotalAMT.ToString() + ", '" + vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd") + "', " + Environment.NewLine +
                                     "             '" + vBuMan + "', '" + vCar_No + "', " + vCash_1000.ToString() + ", " + vCash_500.ToString() + ", " + Environment.NewLine +
                                     "             " + vCash_100.ToString() + ", '" + vRSDate.Year.ToString("D4") + "/" + vRSDate.ToString("MM/dd") + "')";
                    }
                    else
                    {
                        vInsertStr = vInsertStr + Environment.NewLine +
                                     "insert into CASHBILL (CBNo, BuDate, CashType, DepNo, Driver, LinesNo, Car_ID, " + Environment.NewLine +
                                     "             CashAMT, RootAMT, TempAMT, NormalQty, OldQty, TotalAMT, CBDate, BuMan, Car_No, " + Environment.NewLine +
                                     "             Cash_1000, Cash_500, Cash_100, RSDate)" + Environment.NewLine +
                                     "     values ('" + vCBNo + "', '" + vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd") + "', '" + vCashType + "', " + Environment.NewLine +
                                     "             '" + vCashDepNo + "', '" + vDriver + "', '" + vLinesNoStr + "', '" + vCar_ID + "', " + vCashAMT.ToString() + ", " + Environment.NewLine +
                                     "             " + vRootAMT.ToString() + ", " + vTempAMT.ToString() + ", " + vNormalQty.ToString() + ", " + vOldQty.ToString() + ", " + Environment.NewLine +
                                     "             " + vTotalAMT.ToString() + ", '" + vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd") + "', " + Environment.NewLine +
                                     "             '" + vBuMan + "', '" + vCar_No + "', " + vCash_1000.ToString() + ", " + vCash_500.ToString() + ", " + Environment.NewLine +
                                     "             " + vCash_100.ToString() + ", '" + vRSDate.Year.ToString("D4") + "/" + vRSDate.ToString("MM/dd") + "')";
                    }
                }
            }
            if (vInsertStr != "")
            {
                try
                {
                    PF.ExecSQL(vConnStr, vInsertStr);
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

        protected void bbCloseData_Click(object sender, EventArgs e)
        {
            plShowData.Visible = false;
        }
    }
}