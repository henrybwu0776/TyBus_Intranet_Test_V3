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
    public partial class StoreOutAllList : System.Web.UI.Page
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
        DateTime vCalDate = DateTime.Today.AddMonths(-1);

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

                    //事件日期
                    string vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalDate_S.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalDate_S.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCalDate_E.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCalDate_E.Attributes["onClick"] = vCaseDateScript;

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

        /// <summary>
        /// 取回主檔資料的語法
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr_Main()
        {
            DateTime vCalDate_S = (eCalDate_S.Text.Trim() != "") ? DateTime.Parse(eCalDate_S.Text.Trim()) : DateTime.Today;
            string vCalDateStr_S = vCalDate_S.Year.ToString() + "/" + vCalDate_S.ToString("MM/dd");
            DateTime vCalDate_E = (eCalDate_E.Text.Trim() != "") ? DateTime.Parse(eCalDate_E.Text.Trim()) : DateTime.Today;
            string vCalDateStr_E = vCalDate_E.Year.ToString() + "/" + vCalDate_E.ToString("MM/dd");

            string vWStr_Date = ((eCalDate_S.Text.Trim() != "") && (eCalDate_E.Text.Trim() != "")) ? "   and a.[Day] between '" + vCalDateStr_S + "' and '" + vCalDateStr_E + "' " + Environment.NewLine :
                                ((eCalDate_S.Text.Trim() != "") && (eCalDate_E.Text.Trim() == "")) ? "   and a.[Day] = '" + vCalDateStr_S + "' " + Environment.NewLine :
                                ((eCalDate_S.Text.Trim() == "") && (eCalDate_E.Text.Trim() != "")) ? "   and a.[Day] = '" + vCalDateStr_E + "' " + Environment.NewLine : "";
            string vResultStr = "select a.SheetNo, (select ClassTxt from DBDICB where ClassNo = a.Mode and FKey = '領料單A         sheeta          MODE') Mode_C, " + Environment.NewLine +
                                "       a.Seller, (select [Name] from Employee where EmpNo = a.Seller) SellerName, " + Environment.NewLine +
                                "       a.BuMan, (select [Name] from Employee where EmpNo = a.BuMan) BuManName, " + Environment.NewLine +
                                "       a.[Day], a.Remark Remark_A " + Environment.NewLine +
                                "  from SheetA a " + Environment.NewLine +
                                " where a.SheetType = 'E' " + Environment.NewLine +
                                "   and a.SheetClass = '1' " + Environment.NewLine +
                                "   and a.SheetDiff = 'M' " + Environment.NewLine +
                                vWStr_Date +
                                " order by a.SheetNo ";
            return vResultStr;
        }

        /// <summary>
        /// 取回明細資料的語法
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr_Detail()
        {
            DateTime vCalDate_S = (eCalDate_S.Text.Trim() != "") ? DateTime.Parse(eCalDate_S.Text.Trim()) : DateTime.Today;
            string vCalDateStr_S = vCalDate_S.Year.ToString() + "/" + vCalDate_S.ToString("MM/dd");
            DateTime vCalDate_E = (eCalDate_E.Text.Trim() != "") ? DateTime.Parse(eCalDate_E.Text.Trim()) : DateTime.Today;
            string vCalDateStr_E = vCalDate_E.Year.ToString() + "/" + vCalDate_E.ToString("MM/dd");

            string vWStr_Date = ((eCalDate_S.Text.Trim() != "") && (eCalDate_E.Text.Trim() != "")) ? "          and a.[Day] between '" + vCalDateStr_S + "' and '" + vCalDateStr_E + "' " + Environment.NewLine :
                                ((eCalDate_S.Text.Trim() != "") && (eCalDate_E.Text.Trim() == "")) ? "          and a.[Day] = '" + vCalDateStr_S + "' " + Environment.NewLine :
                                ((eCalDate_S.Text.Trim() == "") && (eCalDate_E.Text.Trim() != "")) ? "          and a.[Day] = '" + vCalDateStr_E + "' " + Environment.NewLine : "";
            string vResultStr = "select b.SheetNo, b.Item, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where ClassNo = b.Kind and FKey = '領料單B         sheetb          KIND') Kind_C, " + Environment.NewLine +
                                "       b.Car_ID, b.DepNo, (select[Name] from Department where DepNo = b.DepNo) DepName, " + Environment.NewLine +
                                "       b.[No], (select[Name] from Stock where [No] = b.[No] and SClass = 'M') StockName, " + Environment.NewLine +
                                "       b.PositionB, b.UserQTY, b.Price, b.Amount, b.Remark Remark_B, b.Remark2 Remark_B2 " + Environment.NewLine +
                                "  from SheetB b " + Environment.NewLine +
                                " where b.SheetNo in (" + Environment.NewLine +
                                "       select SheetNo " + Environment.NewLine +
                                "         from SheetA a " + Environment.NewLine +
                                "        where a.SheetType = 'E' " + Environment.NewLine +
                                "          and a.SheetClass = '1' " + Environment.NewLine +
                                "          and a.SheetDiff = 'M' " + Environment.NewLine + vWStr_Date +
                                "       ) " + Environment.NewLine +
                                " order by b.SheetNo, b.Item";
            return vResultStr;
        }

        /// <summary>
        /// 取回所有資料的語法
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr_Total()
        {
            DateTime vCalDate_S = (eCalDate_S.Text.Trim() != "") ? DateTime.Parse(eCalDate_S.Text.Trim()) : DateTime.Today;
            string vCalDateStr_S = vCalDate_S.Year.ToString() + "/" + vCalDate_S.ToString("MM/dd");
            DateTime vCalDate_E = (eCalDate_E.Text.Trim() != "") ? DateTime.Parse(eCalDate_E.Text.Trim()) : DateTime.Today;
            string vCalDateStr_E = vCalDate_E.Year.ToString() + "/" + vCalDate_E.ToString("MM/dd");

            string vWStr_Date = ((eCalDate_S.Text.Trim() != "") && (eCalDate_E.Text.Trim() != "")) ? "   and a.[Day] between '" + vCalDateStr_S + "' and '" + vCalDateStr_E + "' " + Environment.NewLine :
                                ((eCalDate_S.Text.Trim() != "") && (eCalDate_E.Text.Trim() == "")) ? "   and a.[Day] = '" + vCalDateStr_S + "' " + Environment.NewLine :
                                ((eCalDate_S.Text.Trim() == "") && (eCalDate_E.Text.Trim() != "")) ? "   and a.[Day] = '" + vCalDateStr_E + "' " + Environment.NewLine : "";
            string vResultStr = "select a.SheetNo, (select ClassTxt from DBDICB where ClassNo = a.Mode and FKey = '領料單A         sheeta          MODE') Mode_C, " + Environment.NewLine +
                                "       a.Seller, (select[Name] from Employee where EmpNo = a.Seller) SellerName, " + Environment.NewLine +
                                "       a.BuMan, (select[Name] from Employee where EmpNo = a.BuMan) BuManName, " + Environment.NewLine +
                                "       a.[Day], a.Remark Remark_A, " + Environment.NewLine +
                                "       b.Item, (select ClassTxt from DBDICB where ClassNo = b.Kind and FKey = '領料單B         sheetb          KIND') Kind_C, " + Environment.NewLine +
                                "       b.Car_ID, b.DepNo, (select[Name] from Department where DepNo = b.DepNo) DepName, " + Environment.NewLine +
                                "       b.[No], (select[Name] from Stock where [No] = b.[No] and SClass = 'M') StockName, " + Environment.NewLine +
                                "       b.PositionB, b.UserQTY, b.Price, b.Amount, b.Remark Remark_B, b.Remark2 Remark_B2 " + Environment.NewLine +
                                "  from SheetA a left join SheetB b on b.SheetNo = a.SheetNo " + Environment.NewLine +
                                " where a.SheetType = 'E' " + Environment.NewLine +
                                "   and a.SheetClass = '1' " + Environment.NewLine +
                                "   and a.SheetDiff = 'M' " + Environment.NewLine +
                                vWStr_Date +
                                " order by a.SheetNo, b.Item";
            return vResultStr;
        }

        /// <summary>
        /// 轉出 EXCEL
        /// </summary>
        private void ExportExcel()
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
            string vSelStr = "";
            string vFileName = "材料領料單清冊";
            string vSheetName = "";
            DateTime vDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            //先產生主檔的工作表
            using (SqlConnection connExcel_Main = new SqlConnection(vConnStr))
            {
                vSelStr = GetSelectStr_Main();
                SqlCommand cmdExcel_Main = new SqlCommand(vSelStr, connExcel_Main);
                connExcel_Main.Open();
                SqlDataReader drExcel_Main = cmdExcel_Main.ExecuteReader();
                if (drExcel_Main.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    vSheetName = vFileName.Trim() + "--主檔資料";
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel_Main.FieldCount; i++)
                    {
                        vHeaderText = (drExcel_Main.GetName(i).ToUpper() == "SHEETNO") ? "領料單號" :
                                      (drExcel_Main.GetName(i).ToUpper() == "MODE_C") ? "模式" :
                                      (drExcel_Main.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel_Main.GetName(i).ToUpper() == "SELLER") ? "" :
                                      (drExcel_Main.GetName(i).ToUpper() == "SELLERNAME") ? "領料人員" :
                                      (drExcel_Main.GetName(i).ToUpper() == "BUMAN") ? "" :
                                      (drExcel_Main.GetName(i).ToUpper() == "BUMANNAME") ? "開單人員" :
                                      (drExcel_Main.GetName(i).ToUpper() == "DAY") ? "日期" :
                                      (drExcel_Main.GetName(i).ToUpper() == "REMARK_A") ? "摘要" : drExcel_Main.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    while (drExcel_Main.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel_Main.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if ((drExcel_Main.GetName(i).ToUpper() == "DAY") && (drExcel_Main[i].ToString() != ""))
                            {
                                string vTempStr = drExcel_Main[i].ToString();
                                vDate = DateTime.Parse(drExcel_Main[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vDate.Year.ToString("D4") + "/" + vDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel_Main[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                        }
                    }
                }
            }

            //再產生明細的工作表
            using (SqlConnection connExcel_Detail = new SqlConnection(vConnStr))
            {
                vSelStr = GetSelectStr_Detail();
                SqlCommand cmdExcel_Detail = new SqlCommand(vSelStr, connExcel_Detail);
                connExcel_Detail.Open();
                SqlDataReader drExcel_Detail = cmdExcel_Detail.ExecuteReader();
                if (drExcel_Detail.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    vSheetName = vFileName.Trim() + "--明細資料";
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel_Detail.FieldCount; i++)
                    {
                        vHeaderText = (drExcel_Detail.GetName(i).ToUpper() == "SHEETNO") ? "領料單號" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "ITEM") ? "項次" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "KIND_C") ? "型別" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "CAR_ID") ? "牌照號碼" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "DEPNO") ? "" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "NO") ? "料號" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "STOCKNAME") ? "品名" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "POSITIONB") ? "庫位" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "USERQTY") ? "數量" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "PRICE") ? "單價" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "AMOUNT") ? "金額" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "REMARK_B") ? "明細摘要" :
                                      (drExcel_Detail.GetName(i).ToUpper() == "REMARK_B2") ? "明細摘要2" : drExcel_Detail.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    while (drExcel_Detail.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel_Detail.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if (((drExcel_Detail.GetName(i).ToUpper() == "USERQTY") ||
                                 (drExcel_Detail.GetName(i).ToUpper() == "PRICE") ||
                                 (drExcel_Detail.GetName(i).ToUpper() == "AMOUNT")) && (drExcel_Detail[i].ToString() != ""))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel_Detail[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel_Detail[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                        }
                    }
                }
            }

            //最後產生整合資料的工作表
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                vSelStr = GetSelectStr_Total();
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    vSheetName = vFileName.Trim() + "--整合資料";
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "SHEETNO") ? "領料單號" :
                                      (drExcel.GetName(i).ToUpper() == "MODE_C") ? "模式" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "SELLER") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "SELLERNAME") ? "領料人員" :
                                      (drExcel.GetName(i).ToUpper() == "BUMAN") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "BUMANNAME") ? "開單人員" :
                                      (drExcel.GetName(i).ToUpper() == "DAY") ? "日期" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK_A") ? "摘要" :
                                      (drExcel.GetName(i).ToUpper() == "ITEM") ? "項次" :
                                      (drExcel.GetName(i).ToUpper() == "KIND_C") ? "型別" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "牌照號碼" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "NO") ? "料號" :
                                      (drExcel.GetName(i).ToUpper() == "STOCKNAME") ? "品名" :
                                      (drExcel.GetName(i).ToUpper() == "POSITIONB") ? "庫位" :
                                      (drExcel.GetName(i).ToUpper() == "USERQTY") ? "數量" :
                                      (drExcel.GetName(i).ToUpper() == "PRICE") ? "單價" :
                                      (drExcel.GetName(i).ToUpper() == "AMOUNT") ? "金額" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK_B") ? "明細摘要" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK_B2") ? "明細摘要2" : drExcel.GetName(i);
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
                            if ((drExcel.GetName(i).ToUpper() == "DAY") && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vDate.Year.ToString("D4") + "/" + vDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i).ToUpper() == "USERQTY") ||
                                     (drExcel.GetName(i).ToUpper() == "PRICE") ||
                                     (drExcel.GetName(i).ToUpper() == "AMOUNT")) && (drExcel[i].ToString() != ""))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                        }
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
                    string vCAlDateStr = ((eCalDate_S.Text.Trim() != "") && (eCalDate_E.Text.Trim() != "")) ? "從 " + eCalDate_S.Text.Trim() + " 起至 " + eCalDate_E.Text.Trim() + " 止" :
                                         ((eCalDate_S.Text.Trim() != "") && (eCalDate_E.Text.Trim() == "")) ? eCalDate_S.Text.Trim() :
                                         ((eCalDate_S.Text.Trim() == "") && (eCalDate_E.Text.Trim() != "")) ? eCalDate_E.Text.Trim() : "";
                    string vRecordNote = "匯出檔案_修車廠領料單資料匯出" + Environment.NewLine +
                                         "StoreOutAllList.aspx" + Environment.NewLine +
                                         "開單日期：" + vCAlDateStr;
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

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if ((eCalDate_S.Text.Trim() != "") || (eCalDate_E.Text.Trim() != ""))
            {
                ExportExcel();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇日期區間')");
                Response.Write("</" + "Script>");
                eCalDate_S.Focus();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}