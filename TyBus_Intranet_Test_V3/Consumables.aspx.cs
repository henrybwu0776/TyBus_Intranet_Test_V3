using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class Consumables : System.Web.UI.Page
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
                    //GetItemListData();
                    if (!IsPostBack)
                    {
                        plReport.Visible = false;
                        Session["ConsMode"] = "LIST";
                    }
                    else
                    {
                        OpenData();
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

        /*
        /// <summary>
        /// 取回下拉選單的內容
        /// </summary>
        private void GetItemListData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            //取回耗材分類
            string vSelectStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                     " union all" + Environment.NewLine +
                                     "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsType' order by ClassNo";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelectStr_Temp, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    ddlConsType_Search.Items.Clear();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                        ddlConsType_Search.Items.Add(liTemp);
                    }
                }
            }
        } //*/

        /// <summary>
        /// 取回查詢字串
        /// </summary>
        /// <returns></returns>
        private string OpenListStr()
        {
            string vResultStr = "";
            string vWStr_ConsNo = (eConsNo_Search.Text.Trim() != "") ? "   and ConsNo = '" + eConsNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "   and ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Brand = (eBrand_Search.Text.Trim() != "") ? "   and Brand like '%" + eBrand_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            vResultStr = "select c.ConsNo, c.ConsName, d1.ClassTxt ConsType, c.StockQty, d2.ClassTxt ConsUnit, c.AvgPrice, c.StoreLocation, c.Brand " + Environment.NewLine +
                         "  from Consumables c left join DBDICB d1 on d1.ClassNo = c.ConsType and d1.FKey = '耗材庫存        CONSUMABLES     ConsType' " + Environment.NewLine +
                         "                     left join DBDICB d2 on d2.ClassNo = c.ConsUnit and d2.FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                         " where isnull(c.ConsNo, '') != '' " + Environment.NewLine +
                         vWStr_ConsNo +
                         vWStr_ConsName +
                         vWStr_Brand +
                         " order by c.ConsNo";
            return vResultStr;
        }

        private void OpenData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vListStr = OpenListStr();
            sdsShowDataList.SelectCommand = vListStr;
            gridShowList.DataBind();
        }

        /// <summary>
        /// 查詢資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        /// <summary>
        /// 產生盤點單 (不列印，直接出 EXCEL)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintCheckList_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            double vTempFloat = 0.0;
            string vSQLStrTemp = "select ConsNo, isnull(ConsType, '') ConsType, ConsName, isnull(ConsUnit, '') ConsUnit, Brand, " + Environment.NewLine +
                                 "       isnull(ConsColor, '') ConsColor, isnull(ConsSpec, '') ConsSpec, isnull(ConsSpec2, '') ConsSpec2, " + Environment.NewLine +
                                 "       isnull(StockQty, '') StockQty, cast('' as varchar) as InventoryQty, isnull(IsStopUse, 'X') IsStopUse " + Environment.NewLine +
                                 "  from Consumables " + Environment.NewLine +
                                 " order by ConsType, ConsNo";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSQLStrTemp, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //有查到資料才往下執行
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
                    csData_Float.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csData_Float.Alignment = HorizontalAlignment.Right;

                    HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
                    csData_Float.DataFormat = format.GetFormat("##0.00");

                    string vFileName = "桃園汽車客運總務課庫存耗材盤點表";
                    string vHeaderText = "";
                    string vCellValue = "";
                    int vLinesNo = 0;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int vCellCount = 0; vCellCount < drExcel.FieldCount; vCellCount++)
                    {
                        vHeaderText = (drExcel.GetName(vCellCount).ToUpper() == "CONSNO") ? "庫存編號" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "BRAND") ? "廠牌" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "CONSNAME") ? "品名" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "CONSTYPE") ? "類別" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "CONSCOLOR") ? "顏色" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "STOCKQTY") ? "現有庫存量" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "INVENTORYQTY") ? "實際盤點量" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "CONSUNIT") ? "單位" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "CONSSPEC") ? "規格" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "CONSSPEC2") ? "尺寸" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "ISSTOPUSE") ? "已停用" :
                                      drExcel.GetName(vCellCount).Trim();
                        wsExcel.GetRow(vLinesNo).CreateCell(vCellCount).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(vCellCount).CellStyle = csTitle;
                    }
                    //開始取回資料
                    while (drExcel.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            vHeaderText = drExcel.GetName(i).ToUpper();
                            vCellValue = drExcel[i].ToString().Trim();
                            //if (drExcel[i].ToString().Trim() != "")
                            if (vCellValue != "")
                            {
                                if ((drExcel.GetName(i).ToUpper() == "STOCKQTY") || (drExcel.GetName(i).ToUpper() == "INVENTORYQTY"))
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.TryParse(drExcel[i].ToString().Trim(), out vTempFloat) ? vTempFloat : 0.0);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Float;
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                }
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString().Trim());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
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
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "Consumables.aspx";
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
                        eErrorMSG_Main.Text = eMessage.Message;
                        eErrorMSG_Main.Visible = true;
                    }
                }
            }
        }

        /// <summary>
        /// 匯入盤點量
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbUpdateReserve_Search_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "";
            string vSheetNo = "";
            string vSheetNoItems = "";
            string vMaxIndex = "";
            int vIndex = 0;
            string vItems = "";
            string vSheetRemark = DateTime.Today.ToShortDateString() + "_盤點作業";
            string vSheetRemarkB = DateTime.Today.ToShortDateString() + "_盤點數量更新" + Environment.NewLine;
            string vFirstCode = "";
            string vIndex_S = "";
            string vConsNo_Temp = "";
            string vConsName_Temp = "";
            string vBrand_Temp = "";
            string vConsType_Temp = "";
            string vConsUnit_Temp = "";
            string vConsColor_Temp = "";
            string vConsSpec_Temp = "";
            string vConsSpec2_Temp = "";
            string vIsStopUse_Temp = "";
            Double vTempFloat;
            Double vInventoryQty_Temp = 0.0;
            Double vStockQty_Temp = 0.0;
            Double vQtyDift = 0.0;
            int vQtyMode = 1;

            if (fuExcel.FileName != "") //有選擇匯入用的檔案
            {
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回轉入檔的副檔名
                switch (vExtName)
                {
                    //根據不同的副檔名來判斷版本
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        try
                        {
                            HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent); //開啟 EXCEL 檔
                            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0); //取回第一個工作表
                            if (sheetExcel_H.LastRowNum > 0) //判斷工作表有沒有內容
                            {
                                //先寫入表頭
                                vSheetNo = PF.GetConsSheetNo(vConnStr, "SS");
                                using (SqlDataSource dsTempA = new SqlDataSource())
                                {
                                    dsTempA.ConnectionString = vConnStr;
                                    dsTempA.InsertCommand = "insert into ConsSheetA(SheetNo, SheetMode, BuDate, BuMan, SheetStatus, StatusDate, RemarkA)" + Environment.NewLine +
                                                            "values(@SheetNo, 'SS', GetDate(), @BuMan, '998', GetDate(), @RemarkA)";
                                    dsTempA.InsertParameters.Clear();
                                    dsTempA.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                    dsTempA.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                    dsTempA.InsertParameters.Add(new Parameter("RemarkA", DbType.String, vSheetRemark));
                                    dsTempA.Insert();
                                }
                                //準備寫入耗材盤點量
                                for (int i = 0; i <= sheetExcel_H.LastRowNum; i++)
                                {
                                    //逐行讀取資料
                                    HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                    if ((vRowExcel_H != null) && (vRowExcel_H.Cells[0].StringCellValue.Trim() != "庫存編號"))
                                    {
                                        //讀回來的資料列不是空值，或者不是標題列
                                        vConsNo_Temp = vRowExcel_H.Cells[0].StringCellValue.Trim(); //庫存編號
                                        vConsType_Temp = vRowExcel_H.Cells[1].StringCellValue.Trim(); //類別
                                        vConsName_Temp = vRowExcel_H.Cells[2].StringCellValue.Trim(); //品名
                                        vConsUnit_Temp = vRowExcel_H.Cells[3].StringCellValue.Trim(); //單位
                                        vBrand_Temp = vRowExcel_H.Cells[4].StringCellValue.Trim(); //廠牌
                                        vConsColor_Temp = vRowExcel_H.Cells[5].StringCellValue.Trim(); //顏色
                                        vConsSpec_Temp = vRowExcel_H.Cells[6].StringCellValue.Trim(); //規格
                                        vConsSpec2_Temp = vRowExcel_H.Cells[7].StringCellValue.Trim(); //尺寸
                                        vIsStopUse_Temp = (vRowExcel_H.Cells[10].StringCellValue.Trim() == "V") ? vRowExcel_H.Cells[10].StringCellValue.Trim() : "X"; //是否停用
                                        if (vConsNo_Temp == "")
                                        {
                                            //料號欄位是空白，表示是新料號
                                            vFirstCode = vConsType_Temp.Trim() + "-02-";
                                            vSQLStr_Temp = "select max(ConsNo) MaxNo from Consumables where ConsNo like '" + vFirstCode + "%' ";
                                            vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                                            vIndex_S = Int32.TryParse(vMaxIndex.Replace(vFirstCode, ""), out vIndex) ? (vIndex + 1).ToString("D4") : "0001";
                                            vConsNo_Temp = vFirstCode + vIndex_S;
                                            using (SqlDataSource dsConsTemp = new SqlDataSource())
                                            {
                                                dsConsTemp.ConnectionString = vConnStr;
                                                dsConsTemp.InsertCommand = "insert into Consumables " + Environment.NewLine +
                                                                           "       (ConsNo, ConsName, ConsType, Brand, ConsUnit, ConsColor, ConsSpec, ConsSpec2, BuMan, BuDate, IsStopUse) " + Environment.NewLine +
                                                                           "values (@ConsNo, @ConsName, @ConsType, @Brand, @ConsUnit, @ConsColor, @ConsSpec, @ConsSpec2, @BuMan, GetDate(), @IsStopUse)";
                                                dsConsTemp.InsertParameters.Clear();
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsName", DbType.String, vConsName_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsType", DbType.String, vConsType_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("Brand", DbType.String, vBrand_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsUnit", DbType.String, vConsUnit_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsColor", DbType.String, vConsColor_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsSpec", DbType.String, vConsSpec_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsSpec2", DbType.String, vConsSpec2_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                dsConsTemp.InsertParameters.Add(new Parameter("IsStopUse", DbType.String, vIsStopUse_Temp));
                                                dsConsTemp.Insert();
                                            }
                                        }
                                        vStockQty_Temp = Double.TryParse(vRowExcel_H.Cells[8].NumericCellValue.ToString().Trim(), out vTempFloat) ? vTempFloat : 0.0; //庫存量
                                        vInventoryQty_Temp = Double.TryParse(vRowExcel_H.Cells[9].NumericCellValue.ToString().Trim(), out vTempFloat) ? vTempFloat : 0.0; //實際數量
                                        vQtyMode = (vInventoryQty_Temp >= vStockQty_Temp) ? 1 : -1;
                                        vQtyDift = Math.Abs(vInventoryQty_Temp - vStockQty_Temp);
                                        vSQLStr_Temp = "select max(Items) MaxNo from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                                        vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                                        vItems = Int32.TryParse(vMaxIndex, out vIndex) ? (vIndex + 1).ToString("D4") : "0001";
                                        vSheetNoItems = vSheetNo + vItems;
                                        using (SqlDataSource dsTempB = new SqlDataSource())
                                        {
                                            dsTempB.ConnectionString = vConnStr;
                                            dsTempB.InsertCommand = "insert into ConsSheetB(SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                                                    "QtyMode, ItemStatus, BuMan, BuDate)" + Environment.NewLine +
                                                                    "values (@SheetNoItems, @SheetNo, @Items, @ConsNo, 0, @Quantity, 0, @RemarkB, " + Environment.NewLine +
                                                                    "        @QtyMode, '001', @BuMan, GetDate())";
                                            dsTempB.InsertParameters.Clear();
                                            dsTempB.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                                            dsTempB.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                            dsTempB.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                                            dsTempB.InsertParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo_Temp));
                                            dsTempB.InsertParameters.Add(new Parameter("Quantity", DbType.Double, vQtyDift.ToString()));
                                            dsTempB.InsertParameters.Add(new Parameter("RemarkB", DbType.String, vSheetRemarkB));
                                            dsTempB.InsertParameters.Add(new Parameter("QtyMode", DbType.Int32, vQtyMode.ToString()));
                                            dsTempB.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                            dsTempB.Insert();
                                        }
                                    }
                                }
                                string vRecordNote = "匯入庫存量" + Environment.NewLine +
                                                     "盤點單號：" + vSheetNo + Environment.NewLine +
                                                     "IAConsumables.aspx";
                                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            }
                            else
                            {
                                eErrorMSG_Main.Text = "查無資料";
                                eErrorMSG_Main.Visible = true;
                            }
                        }
                        catch (Exception eMessage)
                        {
                            eErrorMSG_Main.Text = eMessage.Message;
                            eErrorMSG_Main.Visible = true;
                        }
                        break;
                    case ".xlsx": //新版 EXCEL (2010 後) 格式
                        try
                        {
                            XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent); //開啟 EXCEL 檔
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0); //取回第一個工作表
                            if (sheetExcel_X.LastRowNum > 0) //判斷工作表有沒有內容
                            {
                                //先寫入表頭
                                vSheetNo = PF.GetConsSheetNo(vConnStr, "SS");
                                using (SqlDataSource dsTempA = new SqlDataSource())
                                {
                                    dsTempA.ConnectionString = vConnStr;
                                    dsTempA.InsertCommand = "insert into ConsSheetA(SheetNo, SheetMode, BuDate, BuMan, SheetStatus, StatusDate, RemarkA)" + Environment.NewLine +
                                                            "values(@SheetNo, 'SS', GetDate(), @BuMan, '998', GetDate(), @RemarkA)";
                                    dsTempA.InsertParameters.Clear();
                                    dsTempA.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                    dsTempA.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                    dsTempA.InsertParameters.Add(new Parameter("RemarkA", DbType.String, vSheetRemark));
                                    dsTempA.Insert();
                                }
                                //準備寫入耗材盤點量
                                for (int i = 0; i <= sheetExcel_X.LastRowNum; i++)
                                {
                                    //逐行讀取資料
                                    XSSFRow vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                    if ((vRowExcel_X != null) && (vRowExcel_X.Cells[0].StringCellValue.Trim() != "庫存編號"))
                                    {
                                        //讀回來的資料列不是空值，或者不是標題列
                                        vConsNo_Temp = vRowExcel_X.Cells[0].StringCellValue.Trim();
                                        vConsType_Temp = vRowExcel_X.Cells[1].StringCellValue.Trim();
                                        vConsName_Temp = vRowExcel_X.Cells[2].StringCellValue.Trim();
                                        vConsUnit_Temp = vRowExcel_X.Cells[3].StringCellValue.Trim();
                                        vBrand_Temp = vRowExcel_X.Cells[4].StringCellValue.Trim();
                                        vConsColor_Temp = vRowExcel_X.Cells[5].StringCellValue.Trim();
                                        vConsSpec_Temp = vRowExcel_X.Cells[6].StringCellValue.Trim();
                                        vConsSpec2_Temp = vRowExcel_X.Cells[7].StringCellValue.Trim();
                                        vIsStopUse_Temp = (vRowExcel_X.Cells[10].StringCellValue.Trim() == "V") ? vRowExcel_X.Cells[10].StringCellValue.Trim() : "X";
                                        if (vConsNo_Temp == "")
                                        {
                                            //料號欄位是空白，表示是新料號
                                            vFirstCode = vConsType_Temp.Trim() + "-02-";
                                            vSQLStr_Temp = "select max(ConsNo) MaxNo from Consumables where ConsNo like '" + vFirstCode + "%' ";
                                            vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                                            vIndex_S = Int32.TryParse(vMaxIndex.Replace(vFirstCode, ""), out vIndex) ? (vIndex + 1).ToString("D4") : "0001";
                                            vConsNo_Temp = vFirstCode + vIndex_S;
                                            using (SqlDataSource dsConsTemp = new SqlDataSource())
                                            {
                                                dsConsTemp.ConnectionString = vConnStr;
                                                dsConsTemp.InsertCommand = "insert into Consumables " + Environment.NewLine +
                                                                           "       (ConsNo, ConsName, ConsType, Brand, ConsUnit, ConsColor, ConsSpec, ConsSpec2, BuMan, BuDate, IsStopUse) " + Environment.NewLine +
                                                                           "values (@ConsNo, @ConsName, @ConsType, @Brand, @ConsUnit, @ConsColor, @ConsSpec, @ConsSpec2, @BuMan, GetDate(), @IsStopUse)";
                                                dsConsTemp.InsertParameters.Clear();
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsName", DbType.String, vConsName_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsType", DbType.String, vConsType_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("Brand", DbType.String, vBrand_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsUnit", DbType.String, vConsUnit_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsColor", DbType.String, vConsColor_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsSpec", DbType.String, vConsSpec_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("ConsSpec2", DbType.String, vConsSpec2_Temp));
                                                dsConsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                dsConsTemp.InsertParameters.Add(new Parameter("IsStopUse", DbType.String, vIsStopUse_Temp));
                                                dsConsTemp.Insert();
                                            }
                                        }
                                        vStockQty_Temp = Double.TryParse(vRowExcel_X.Cells[8].NumericCellValue.ToString().Trim(), out vTempFloat) ? vTempFloat : 0.0;
                                        vInventoryQty_Temp = Double.TryParse(vRowExcel_X.Cells[9].NumericCellValue.ToString().Trim(), out vTempFloat) ? vTempFloat : 0.0;
                                        vQtyMode = (vInventoryQty_Temp > vStockQty_Temp) ? 1 : -1;
                                        vQtyDift = Math.Abs(vInventoryQty_Temp - vStockQty_Temp);
                                        vSQLStr_Temp = "select max(Items) MaxNo from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                                        vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                                        vItems = Int32.TryParse(vMaxIndex, out vIndex) ? (vIndex + 1).ToString("D4") : "0001";
                                        vSheetNoItems = vSheetNo + vItems;
                                        using (SqlDataSource dsTempB = new SqlDataSource())
                                        {
                                            dsTempB.ConnectionString = vConnStr;
                                            dsTempB.InsertCommand = "insert into ConsSheetB(SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                                                    "QtyMode, ItemStatus, BuMan, BuDate)" + Environment.NewLine +
                                                                    "values (@SheetNoItems, @SheetNo, @Items, @ConsNo, 0, @Quantity, 0, @RemarkB, " + Environment.NewLine +
                                                                    "        @QtyMode, '001', @BuMan, GetDate())";
                                            dsTempB.InsertParameters.Clear();
                                            dsTempB.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                                            dsTempB.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                            dsTempB.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                                            dsTempB.InsertParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo_Temp));
                                            dsTempB.InsertParameters.Add(new Parameter("Quantity", DbType.Double, vQtyDift.ToString()));
                                            dsTempB.InsertParameters.Add(new Parameter("RemarkB", DbType.String, vSheetRemarkB));
                                            dsTempB.InsertParameters.Add(new Parameter("QtyMode", DbType.Int32, vQtyMode.ToString()));
                                            dsTempB.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                            dsTempB.Insert();
                                        }
                                    }
                                }
                                string vRecordNote = "匯入庫存量" + Environment.NewLine +
                                                     "盤點單號：" + vSheetNo + Environment.NewLine +
                                                     "IAConsumables.aspx";
                                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            }
                            else
                            {
                                eErrorMSG_Main.Text = "查無資料";
                                eErrorMSG_Main.Visible = true;
                            }
                        }
                        catch (Exception eMessage)
                        {
                            eErrorMSG_Main.Text = eMessage.Message;
                            eErrorMSG_Main.Visible = true;
                        }
                        break;
                    default:
                        eErrorMSG_Main.Text = "您指定的檔案不是 EXCEL 檔，請重新確認檔案！";
                        eErrorMSG_Main.Visible = true;
                        break;
                }
            }
            else
            {
                eErrorMSG_Main.Text = "您沒有選擇要轉入的檔案，請重新確認！";
                eErrorMSG_Main.Visible = true;
            }
        }

        /// <summary>
        /// 離開功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 料號查詢輸入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eConsNo_Search_TextChanged(object sender, EventArgs e)
        {
            lbErrorMSG_ConsName.Text = "";
            lbErrorMSG_ConsName.Visible = false;
            string vTempStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vConsNo = eConsNo_Search.Text.Trim();
            vTempStr = "select ConsName from Consumables where ConsNo = '" + vConsNo + "' ";
            string vConsName = PF.GetValue(vConnStr, vTempStr, "ConsName");
            if (vConsName == "")
            {
                eConsNo_Search.Text = "";
                lbErrorMSG_ConsName.Text = "查無 [ " + vConsNo + " ] 料號";
                lbErrorMSG_ConsName.Visible = true;
            }
        }

        protected void gridShowList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridShowList.PageIndex = e.NewPageIndex;
            gridShowList.DataBind();
        }

        /// <summary>
        /// 資料表格完成繫結
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void fvShowList_DataBound(object sender, EventArgs e)
        {
            DropDownList ddlConsType;
            Label eConsType;
            DropDownList ddlConsUnit;
            Label eConsUnit;
            CheckBox cbIsStopUse;
            Label eIsStopUse;
            CheckBox cbIsInorder;
            Label eIsInorder;
            DropDownList ddlBrand;
            Label eBrand;
            string vSelectType = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                 " union all" + Environment.NewLine +
                                 "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsType' order by ClassNo";
            string vSelectUnit = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                 " union all" + Environment.NewLine +
                                 "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsUnit' order by ClassNo";
            string vSelectBrand = "select cast('' as varchar) BrandCode, cast('' as varchar) BrandName " + Environment.NewLine +
                                  " union all" + Environment.NewLine +
                                  "select BrandCode, BrandName from ConsBrand where BelongGroup = '02' order by BrandCode ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvShowList.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    cbIsStopUse = (CheckBox)fvShowList.FindControl("cbIsStopUse_List");
                    if (cbIsStopUse != null)
                    {
                        cbIsInorder = (CheckBox)fvShowList.FindControl("cbIsInorder_List");
                        eIsStopUse = (Label)fvShowList.FindControl("eIsStopUse_List");
                        eIsInorder = (Label)fvShowList.FindControl("eIsInorder_List");
                        cbIsStopUse.Checked = (eIsStopUse.Text.Trim() == "V") ? true : false;
                        cbIsInorder.Checked = (eIsInorder.Text.Trim() == "V") ? true : false;
                    }
                    break;
                case FormViewMode.Edit:
                    ddlConsType = (DropDownList)fvShowList.FindControl("eConsType_C_Edit");
                    if (ddlConsType != null)
                    {
                        using (SqlConnection connType = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdType = new SqlCommand(vSelectType, connType);
                            connType.Open();
                            SqlDataReader drType = cmdType.ExecuteReader();
                            if (drType.HasRows)
                            {
                                ddlConsType.Items.Clear();
                                while (drType.Read())
                                {
                                    ListItem liTemp = new ListItem(drType["ClassTxt"].ToString().Trim(), drType["ClassNo"].ToString().Trim());
                                    ddlConsType.Items.Add(liTemp);
                                }
                            }
                        }
                        ddlConsUnit = (DropDownList)fvShowList.FindControl("eConsUnit_C_Edit");
                        using (SqlConnection connUnit = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdUnit = new SqlCommand(vSelectUnit, connUnit);
                            connUnit.Open();
                            SqlDataReader drUnit = cmdUnit.ExecuteReader();
                            if (drUnit.HasRows)
                            {
                                ddlConsUnit.Items.Clear();
                                while (drUnit.Read())
                                {
                                    ListItem liTemp = new ListItem(drUnit["ClassTxt"].ToString().Trim(), drUnit["ClassNo"].ToString().Trim());
                                    ddlConsUnit.Items.Add(liTemp);
                                }
                            }
                        }
                        ddlBrand = (DropDownList)fvShowList.FindControl("eBrand_C_Edit");
                        using (SqlConnection connBrand = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdBrand = new SqlCommand(vSelectBrand, connBrand);
                            connBrand.Open();
                            SqlDataReader drBrand = cmdBrand.ExecuteReader();
                            if (drBrand.HasRows)
                            {
                                ddlBrand.Items.Clear();
                                while (drBrand.Read())
                                {
                                    ListItem liTemp = new ListItem(drBrand["BrandCode"].ToString().Trim(), drBrand["BrandName"].ToString().Trim());
                                    ddlBrand.Items.Add(liTemp);
                                }
                            }
                        }
                        eConsType = (Label)fvShowList.FindControl("eConsType_Edit");
                        eConsUnit = (Label)fvShowList.FindControl("eConsUnit_Edit");
                        eBrand = (Label)fvShowList.FindControl("eBrand_Edit");
                        ddlConsType.SelectedIndex = (eConsType.Text.Trim() != "") ? ddlConsType.Items.IndexOf(ddlConsType.Items.FindByValue(eConsType.Text.Trim())) : 0;
                        ddlConsUnit.SelectedIndex = (eConsUnit.Text.Trim() != "") ? ddlConsUnit.Items.IndexOf(ddlConsUnit.Items.FindByValue(eConsUnit.Text.Trim())) : 0;
                        ddlBrand.SelectedIndex = (eBrand.Text.Trim() != "") ? ddlBrand.Items.IndexOf(ddlBrand.Items.FindByValue(eBrand.Text.Trim())) : 0;
                        cbIsStopUse = (CheckBox)fvShowList.FindControl("cbIsStopUse_Edit");
                        cbIsInorder = (CheckBox)fvShowList.FindControl("cbIsInorder_Edit");
                        eIsStopUse = (Label)fvShowList.FindControl("eIsStopUse_Edit");
                        eIsInorder = (Label)fvShowList.FindControl("eIsInorder_Edit");
                        cbIsStopUse.Checked = (eIsStopUse.Text.Trim() == "V") ? true : false;
                        cbIsInorder.Checked = (eIsInorder.Text.Trim() == "V") ? true : false;
                    }
                    break;
                case FormViewMode.Insert:
                    ddlConsType = (DropDownList)fvShowList.FindControl("eConsType_C_INS");
                    if (ddlConsType != null)
                    {
                        using (SqlConnection connType = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdType = new SqlCommand(vSelectType, connType);
                            connType.Open();
                            SqlDataReader drType = cmdType.ExecuteReader();
                            if (drType.HasRows)
                            {
                                ddlConsType.Items.Clear();
                                while (drType.Read())
                                {
                                    ListItem liTemp = new ListItem(drType["ClassTxt"].ToString().Trim(), drType["ClassNo"].ToString().Trim());
                                    ddlConsType.Items.Add(liTemp);
                                }
                            }
                        }
                        ddlConsUnit = (DropDownList)fvShowList.FindControl("eConsUnit_C_INS");
                        using (SqlConnection connUnit = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdUnit = new SqlCommand(vSelectUnit, connUnit);
                            connUnit.Open();
                            SqlDataReader drUnit = cmdUnit.ExecuteReader();
                            if (drUnit.HasRows)
                            {
                                ddlConsUnit.Items.Clear();
                                while (drUnit.Read())
                                {
                                    ListItem liTemp = new ListItem(drUnit["ClassTxt"].ToString().Trim(), drUnit["ClassNo"].ToString().Trim());
                                    ddlConsUnit.Items.Add(liTemp);
                                }
                            }
                        }
                        ddlBrand = (DropDownList)fvShowList.FindControl("eBrand_C_INS");
                        using (SqlConnection connBrand = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdBrand = new SqlCommand(vSelectBrand, connBrand);
                            connBrand.Open();
                            SqlDataReader drBrand = cmdBrand.ExecuteReader();
                            if (drBrand.HasRows)
                            {
                                ddlBrand.Items.Clear();
                                while (drBrand.Read())
                                {
                                    ListItem liTemp = new ListItem(drBrand["BrandCode"].ToString().Trim(), drBrand["BrandName"].ToString().Trim());
                                    ddlBrand.Items.Add(liTemp);
                                }
                            }
                        }
                    }
                    break;
            }
        }

        /// <summary>
        /// 編輯畫面確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            //先定義可以變更的欄位
            Label eConsNo = (Label)fvShowList.FindControl("eConsNo_Edit");
            TextBox eConsName = (TextBox)fvShowList.FindControl("eConsName_Edit");
            Label eConsType = (Label)fvShowList.FindControl("eConsType_Edit");
            Label eConsUnit = (Label)fvShowList.FindControl("eConsUnit_Edit");
            TextBox eConsColor = (TextBox)fvShowList.FindControl("eConsColor_Edit");
            TextBox eConsSpec = (TextBox)fvShowList.FindControl("eConsSpec_Edit");
            TextBox eConsSpec2 = (TextBox)fvShowList.FindControl("eConsSpec2_Edit");
            Label eBrand = (Label)fvShowList.FindControl("eBrand_Edit");
            TextBox eStoreLocation = (TextBox)fvShowList.FindControl("eStoreLocation_Edit");
            Label eIsStopUse = (Label)fvShowList.FindControl("eIsStopUse_Edit");
            Label eIsInorder = (Label)fvShowList.FindControl("eIsInorder_Edit");
            TextBox eRemark = (TextBox)fvShowList.FindControl("eRemark_Edit");
            //決定更新的語法
            string vUpdateStr = "update Consumables " + Environment.NewLine +
                                "   set ConsName = @ConsName, ConsType = @ConsType, ConsUnit = @ConsUnit, ConsColor = @ConsColor,  " + Environment.NewLine +
                                "       ConsSpec = @ConsSpec, ConsSpec2 = @ConsSpec2, Brand = @Brand, StoreLocation = @StoreLocation, " + Environment.NewLine +
                                "       IsStopUse = @IsStopUse, IsInorder = @IsInorder, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                " where ConsNo = @ConsNo ";
            sdsShowDetail.UpdateCommand = vUpdateStr;
            sdsShowDetail.UpdateParameters.Clear();
            sdsShowDetail.UpdateParameters.Add(new Parameter("ConsName", DbType.String, eConsName.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("ConsType", DbType.String, eConsType.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("ConsUnit", DbType.String, eConsUnit.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("ConsColor", DbType.String, eConsColor.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("ConsSpec", DbType.String, eConsSpec.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("ConsSpec2", DbType.String, eConsSpec2.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("Brand", DbType.String, eBrand.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("StoreLocation", DbType.String, eStoreLocation.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("IsStopUse", DbType.String, eIsStopUse.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("IsInorder", DbType.String, eIsInorder.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, eRemark.Text.Trim()));
            sdsShowDetail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
            sdsShowDetail.UpdateParameters.Add(new Parameter("ConsNo", DbType.String, eConsNo.Text.Trim()));
            sdsShowDetail.Update();
            gridShowList.DataBind();
            fvShowList.ChangeMode(FormViewMode.ReadOnly);
            fvShowList.DataBind();
        }

        /// <summary>
        /// 編輯畫面類別下拉選單變更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eConsType_C_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlConsType = (DropDownList)fvShowList.FindControl("eConsType_C_Edit");
            Label eConsType = (Label)fvShowList.FindControl("eConsType_Edit");
            eConsType.Text = ddlConsType.SelectedValue.Trim();
        }

        /// <summary>
        /// 編輯畫面廠商下拉選單變更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eBrand_C_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlBrand = (DropDownList)fvShowList.FindControl("eBrand_C_Edit");
            Label eBrand = (Label)fvShowList.FindControl("eBrand_Edit");
            eBrand.Text = ddlBrand.SelectedValue.Trim();
        }

        protected void eConsUnit_C_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlConsUnit = (DropDownList)fvShowList.FindControl("eConsUnit_C_Edit");
            Label eConsUnit = (Label)fvShowList.FindControl("eConsUnit_Edit");
            eConsUnit.Text = ddlConsUnit.SelectedValue.Trim();
        }

        protected void cbIsStopUse_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsStopUse = (CheckBox)fvShowList.FindControl("cbIsStopUse_Edit");
            Label eIsStopUse = (Label)fvShowList.FindControl("eIsStopUse_Edit");
            eIsStopUse.Text = (cbIsStopUse.Checked) ? "V" : "X";
        }

        protected void cbIsInorder_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsInorder = (CheckBox)fvShowList.FindControl("cbIsInorder_Edit");
            Label eIsInorder = (Label)fvShowList.FindControl("eIsInorder_Edit");
            eIsInorder.Text = (cbIsInorder.Checked) ? "V" : "X";
        }

        /// <summary>
        /// 新增畫面 "確定" 按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            DropDownList eConsType_C = (DropDownList)fvShowList.FindControl("eConsType_C_INS");
            if (eConsType_C != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eConsType = (Label)fvShowList.FindControl("eConsType_INS");
                if (eConsType.Text.Trim() != "") //有指定類別才繼續
                {
                    TextBox eConsName = (TextBox)fvShowList.FindControl("eConsName_INS");
                    if (eConsName.Text.Trim() != "") //品名不是空白才繼續
                    {
                        //決定編號前綴詞
                        string vFirstCode = eConsType.Text.Trim() + "-02-";
                        //找出同類別目前最大編號
                        string vSQLStr = "select MAX(ConsNo) MaxNo from Consumables where ConsNo like '" + vFirstCode + "%' ";
                        string vOldNo = PF.GetValue(vConnStr, vSQLStr, "MaxNo");
                        string vNoStr = vOldNo.Replace(vFirstCode, "");
                        //計算新編號
                        int vIndex;
                        string vIndexStr = Int32.TryParse(vNoStr, out vIndex) ? (vIndex + 1).ToString("D4") : "0001";
                        string vConsNo = vFirstCode + vIndexStr;
                        //漳備寫入用的 SQL 語法
                        vSQLStr = "insert into Consumables " + Environment.NewLine +
                                  "       (ConsNo, ConsName, ConsType, ConsUnit, ConsColor, ConsSpec, ConsSpec2, Brand, StockQty, AVGPrice, " + Environment.NewLine +
                                  "        LastInPrice, StoreLocation, IsStopUse, IsInorder, Remark, BuDate, BuMan) " + Environment.NewLine +
                                  "values (@ConsNo, @ConsName, @ConsType, @ConsUnit, @ConsColor, @ConsSpec, @ConsSpec2, @Brand, 0.0, 0.0, " + Environment.NewLine +
                                  "        0.0, @StoreLocation, @IsStopUse, @IsInorder, @Remark, GetDate(), @BuMan)";
                        sdsShowDetail.InsertCommand = vSQLStr;
                        sdsShowDetail.InsertParameters.Clear();
                        sdsShowDetail.InsertParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo));
                        sdsShowDetail.InsertParameters.Add(new Parameter("ConsName", DbType.String, eConsName.Text.Trim()));
                        sdsShowDetail.InsertParameters.Add(new Parameter("ConsType", DbType.String, eConsType.Text.Trim()));
                        Label eConsUnit = (Label)fvShowList.FindControl("eConsUnit_INS");
                        sdsShowDetail.InsertParameters.Add(new Parameter("ConsUnit", DbType.String, (eConsUnit.Text.Trim() != "") ? eConsUnit.Text.Trim() : String.Empty));
                        TextBox eConsColor = (TextBox)fvShowList.FindControl("eConsColor_INS");
                        sdsShowDetail.InsertParameters.Add(new Parameter("ConsColor", DbType.String, (eConsColor.Text.Trim() != "") ? eConsColor.Text.Trim() : String.Empty));
                        TextBox eConsSpec = (TextBox)fvShowList.FindControl("eConsSpec_INS");
                        sdsShowDetail.InsertParameters.Add(new Parameter("ConsSpec", DbType.String, (eConsSpec.Text.Trim() != "") ? eConsSpec.Text.Trim() : String.Empty));
                        TextBox eConsSpec2 = (TextBox)fvShowList.FindControl("eConsSpec2_INS");
                        sdsShowDetail.InsertParameters.Add(new Parameter("ConsSpec2", DbType.String, (eConsSpec2.Text.Trim() != "") ? eConsSpec2.Text.Trim() : String.Empty));
                        Label eBrand = (Label)fvShowList.FindControl("eBrand_INS");
                        sdsShowDetail.InsertParameters.Add(new Parameter("Brand", DbType.String, (eBrand.Text.Trim() != "") ? eBrand.Text.Trim() : String.Empty));
                        TextBox eStoreLocation = (TextBox)fvShowList.FindControl("eStoreLocation_INS");
                        sdsShowDetail.InsertParameters.Add(new Parameter("StoreLocation", DbType.String, (eStoreLocation.Text.Trim() != "") ? eStoreLocation.Text.Trim() : String.Empty));
                        Label eIsStopUse = (Label)fvShowList.FindControl("eIsStopUse_INS");
                        sdsShowDetail.InsertParameters.Add(new Parameter("IsStopUse", DbType.String, (eIsStopUse.Text.Trim() != "") ? eIsStopUse.Text.Trim() : "X"));
                        Label eIsInorder = (Label)fvShowList.FindControl("eIsInorder_INS");
                        sdsShowDetail.InsertParameters.Add(new Parameter("IsInorder", DbType.String, (eIsInorder.Text.Trim() != "") ? eIsInorder.Text.Trim() : "X"));
                        TextBox eRemark = (TextBox)fvShowList.FindControl("eRemark_INS");
                        sdsShowDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark.Text.Trim() != "") ? eRemark.Text.Trim() : String.Empty));
                        sdsShowDetail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                        sdsShowDetail.Insert();
                        gridShowList.DataBind();
                        fvShowList.ChangeMode(FormViewMode.ReadOnly);
                        fvShowList.DataBind();
                    }
                    else
                    {
                        eErrorMSG_Main.Text = "請輸入品名！";
                        eErrorMSG_Main.Visible = true;
                        eConsName.Focus();
                    }
                }
                else
                {
                    eErrorMSG_Main.Text = "請先選擇類別！";
                    eErrorMSG_Main.Visible = true;
                    eConsType.Focus();
                }
            }
        }

        protected void eConsType_C_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlConsType = (DropDownList)fvShowList.FindControl("eConsType_C_INS");
            Label eConsType = (Label)fvShowList.FindControl("eConsType_INS");
            eConsType.Text = ddlConsType.SelectedValue.Trim();
        }

        protected void eBrand_C_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlBrand = (DropDownList)fvShowList.FindControl("eBrand_C_INS");
            Label eBrand = (Label)fvShowList.FindControl("eBrand_INS");
            eBrand.Text = ddlBrand.SelectedValue.Trim();
        }

        protected void eConsUnit_C_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlConsUnit = (DropDownList)fvShowList.FindControl("eConsUnit_C_INS");
            Label eConsUnit = (Label)fvShowList.FindControl("eConsUnit_INS");
            eConsUnit.Text = ddlConsUnit.SelectedValue.Trim();
        }

        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eConsNo = (Label)fvShowList.FindControl("eConsNo_List");
            if (eConsNo != null)
            {
                string vConsNo = eConsNo.Text.Trim();
                if (vConsNo != "")
                {
                    string vDeleteStr = "delete Consumables where ConsNo = @ConsNo";
                    sdsShowDetail.DeleteCommand = vDeleteStr;
                    sdsShowDetail.DeleteParameters.Clear();
                    sdsShowDetail.DeleteParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo));
                    sdsShowDetail.Delete();
                    gridShowList.DataBind();
                    fvShowList.DataBind();
                }
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {

        }
    }
}