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
    public partial class Consumables : Page
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
        private DataTable dtShowData;

        //資料繫結用的變數
        public string vdConsNo;
        public string vdConsName;
        public string vdConsType;
        public string vdConsUnit;
        public string vdConsColor;
        public string vdConsSpec; //規格
        public string vdStockQty;
        public string vdAvgPrice;
        public string vdLastInDate;
        public string vdLastOutDate;
        public string vdConsSpec2; //尺寸
        public string vdStoreLocation; //庫位
        public string vdIsStopUse;
        public string vdInOrder;
        public string vdBuDate;
        public string vdBuMan;
        public string vdModifyDate;
        public string vdModifyMan;
        public string vdRemark;
        public string vdBrand;
        public string vdLastInPrice;

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
                    GetItemListData();
                    if (!IsPostBack)
                    {
                        plReport.Visible = false;
                        SetInputMode(false);
                        Session["ConsMode"] = "LIST";
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
        /// 設定畫面欄位是否可以編輯
        /// </summary>
        /// <param name="fMode">true：可編輯；false：不可編輯</param>
        private void SetInputMode(bool fMode)
        {
            bbInsert.Visible = !fMode;
            bbEdit.Visible = !fMode;
            bbStopUse.Visible = !fMode;
            bbDelete.Visible = !fMode;
            bbOK.Visible = fMode;
            bbCancel.Visible = fMode;
            foreach (Control cTemp in plShowData.Controls)
            {
                if (cTemp is TextBox)
                {
                    (cTemp as TextBox).Enabled = fMode;
                }
                else if (cTemp is CheckBox)
                {
                    (cTemp as CheckBox).Enabled = fMode;
                }
                else if (cTemp is DropDownList)
                {
                    (cTemp as DropDownList).Enabled = fMode;
                }
            }
        }

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
                    ddlConsType.Items.Clear();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                        ddlConsType_Search.Items.Add(liTemp);
                        ddlConsType.Items.Add(liTemp);
                    }
                }
            }
            //取回耗材單位
            vSelectStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                              " union all" + Environment.NewLine +
                              "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsUnit' order by ClassNo";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelectStr_Temp, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    ddlConsUnit.Items.Clear();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                        ddlConsUnit.Items.Add(liTemp);
                    }
                }
            }
            //取回廠商
            vSelectStr_Temp = "select cast('' as varchar) BrandCode, cast('' as varchar) BrandName " + Environment.NewLine +
                              " union all " + Environment.NewLine +
                              "select BrandCode, BrandName from ConsBrand where BelongGroup in ('00', '02') order by BrandCode ";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelectStr_Temp, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    ddlBrand.Items.Clear();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["BrandName"].ToString().Trim(), drTemp["BrandCode"].ToString().Trim());
                        ddlBrand.Items.Add(liTemp);
                    }
                }
            }

        }

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

        private string OpenListStr()
        {
            string vResultStr = "";
            string vWStr_ConsNo = (eConsNo_Search.Text.Trim() != "") ? "   and ConsNo = '" + eConsNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "   and ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_ConsType = (ddlConsType_Search.SelectedValue.Trim() != "") ? "   and ConsType = '" + ddlConsType_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            vResultStr = "select c.ConsNo, c.ConsName, d1.ClassTxt ConsType, c.StockQty, d2.ClassTxt ConsUnit, c.AvgPrice, c.StoreLocation, c.Brand " + Environment.NewLine +
                         "  from Consumables c left join DBDICB d1 on d1.ClassNo = c.ConsType and d1.FKey = '耗材庫存        CONSUMABLES     ConsType' " + Environment.NewLine +
                         "                     left join DBDICB d2 on d2.ClassNo = c.ConsUnit and d2.FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                         " where isnull(c.ConsNo, '') != '' " + Environment.NewLine +
                         vWStr_ConsNo +
                         vWStr_ConsName +
                         vWStr_ConsType +
                         " order by c.ConsNo";
            return vResultStr;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vListStr = OpenListStr();
            using (SqlConnection connShowList = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daShowList = new SqlDataAdapter(vListStr, connShowList);
                connShowList.Open();
                if (dtShowData != null)
                {
                    dtShowData.Clear();
                }
                else
                {
                    dtShowData = new DataTable();
                }
                daShowList.Fill(dtShowData);
            }
            gvShowList.DataSourceID = "";
            gvShowList.DataSource = dtShowData;
            gvShowList.DataBind();
        }

        /// <summary>
        /// 產生盤點單 (不列印，直接出 EXCEL)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintCheckList_Click(object sender, EventArgs e)
        {
            //Int32 vTempINT = 0;
            double vTempFloat = 0.0;
            string vSQLStrTemp = "select ConsNo, Brand, ConsName, isnull(ConsType, '') ConsType, isnull(ConsColor, '') ConsColor, " + Environment.NewLine +
                                 "       isnull(StockQty, '') StockQty, cast('' as varchar) as InventoryQty, StoreLocation " + Environment.NewLine +
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
                    int vLinesNo = 0;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int vCellCount = 0; vCellCount < drExcel.FieldCount; vCellCount++)
                    {
                        vHeaderText = (drExcel.GetName(vCellCount).ToUpper() == "CONSNO") ? "庫存編號" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "BRAND") ? "廠牌" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "CONSNAME") ? "庫存品項" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "CONSTYPE") ? "庫存類別" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "CONSCOLOR") ? "顏色" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "STOCKQTY") ? "現有庫存量" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "INVENTORYQTY") ? "實際盤點量" :
                                      (drExcel.GetName(vCellCount).ToUpper() == "STORELOCATION") ? "庫存位置" :
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
                            if (drExcel[i].ToString().Trim() != "")
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
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(String.Empty);
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
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
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
            string vSQLStrTemp = "";
            string vSheetNo = "";
            string vRemarkA = "";
            string vConsNo_Temp = "";
            double vStockQTY_Temp = 0.0;
            double vInventory_Temp = 0.0;
            double vQtyDift = 0.0;
            string vRemarkB = "";
            string vItems = "";
            string vSheetNoItems = "";
            Int32 vTempINT = 0;
            Int32 vIndex_Count = 0;
            Int32 vQtyMode = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (fuExcel.FileName != "")
            {
                //有選擇檔案才往下走
                string vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    //根據不同的副檔名來判斷版本
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0); //取回第一個工作表
                        if (sheetExcel_H.LastRowNum > 0)
                        {
                            using (SqlDataSource dsTempA = new SqlDataSource())
                            {
                                vSheetNo = PF.GetIASheetNo(vConnStr, "SS");
                                vRemarkA = DateTime.Today.ToShortDateString() + " 盤點作業";
                                dsTempA.ConnectionString = vConnStr;
                                dsTempA.InsertCommand = "insert into ConsSheetA(SheetNo, SheetMode, BuDate, BuMan, SheetStatus, StatusDate, RemarkA)" + Environment.NewLine +
                                                        "values(@SheetNo, 'SS', GetDate(), @BuMan, '998', GetDate(), @RemarkA)";
                                dsTempA.InsertParameters.Clear();
                                dsTempA.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                dsTempA.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                dsTempA.InsertParameters.Add(new Parameter("RemarkA", DbType.String, vRemarkA));
                                dsTempA.Insert();
                            }
                            for (int i = 0; i <= sheetExcel_H.LastRowNum; i++)
                            {
                                HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                if ((vRowExcel_H != null) && (vRowExcel_H.Cells[0].StringCellValue.Trim() != "庫存編號"))
                                {
                                    vConsNo_Temp = vRowExcel_H.Cells[0].StringCellValue.Trim();
                                    vStockQTY_Temp = Int32.TryParse(vRowExcel_H.Cells[5].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vInventory_Temp = Int32.TryParse(vRowExcel_H.Cells[6].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vQtyMode = ((vInventory_Temp - vStockQTY_Temp) < 0) ? -1 : 1;
                                    vQtyDift = Math.Abs(vInventory_Temp - vStockQTY_Temp);
                                    vSQLStrTemp = "select MAX(Items) MaxItem from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                                    vSQLStrTemp = "select MAX(Items) MaxItem from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                                    vIndex_Count = Int32.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "MaxItem"), out vTempINT) ? vTempINT + 1 : 1;
                                    vItems = vIndex_Count.ToString("D4").Trim();
                                    vSheetNoItems = vSheetNo + vItems;
                                    vRemarkB = DateTime.Today.ToShortDateString() + " 盤點數量更新" + Environment.NewLine;
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
                                        dsTempB.InsertParameters.Add(new Parameter("Quantity", DbType.Int32, vQtyDift.ToString()));
                                        dsTempB.InsertParameters.Add(new Parameter("RemarkB", DbType.String, vRemarkB));
                                        dsTempB.InsertParameters.Add(new Parameter("QtyMode", DbType.Int32, vQtyMode.ToString()));
                                        dsTempB.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                        try
                                        {
                                            dsTempB.Insert();
                                            string vRecordNote = "匯入庫存量" + Environment.NewLine +
                                                                 "IAConsumables.aspx";
                                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                                        }
                                        catch (Exception eMessage)
                                        {
                                            Response.Write("<Script language='Javascript'>");
                                            Response.Write("alert('" + eMessage.Message + "')");
                                            Response.Write("</" + "Script>");
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('查無資料。')");
                            Response.Write("</" + "Script>");
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2010 後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0); //取回第一個工作表
                        if (sheetExcel_X.LastRowNum > 0)
                        {
                            using (SqlDataSource dsTempA = new SqlDataSource())
                            {
                                vSheetNo = PF.GetIASheetNo(vConnStr, "SS");
                                vRemarkA = DateTime.Today.ToShortDateString() + " 盤點作業";
                                dsTempA.ConnectionString = vConnStr;
                                dsTempA.InsertCommand = "insert into ConsSheetA(SheetNo, SheetMode, BuDate, BuMan, SheetStatus, StatusDate, RemarkA)" + Environment.NewLine +
                                                        "values(@SheetNo, 'SS', GetDate(), @BuMan, '998', GetDate(), @RemarkA)";
                                dsTempA.InsertParameters.Clear();
                                dsTempA.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                dsTempA.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                dsTempA.InsertParameters.Add(new Parameter("RemarkA", DbType.String, vRemarkA));
                                dsTempA.Insert();
                            }
                            for (int i = 0; i <= sheetExcel_X.LastRowNum; i++)
                            {
                                HSSFRow vRowExcel_X = (HSSFRow)sheetExcel_X.GetRow(i);
                                if ((vRowExcel_X != null) && (vRowExcel_X.Cells[0].StringCellValue.Trim() != "庫存編號"))
                                {
                                    vConsNo_Temp = vRowExcel_X.Cells[0].StringCellValue.Trim();
                                    vStockQTY_Temp = Int32.TryParse(vRowExcel_X.Cells[5].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vInventory_Temp = Int32.TryParse(vRowExcel_X.Cells[6].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vQtyMode = ((vInventory_Temp - vStockQTY_Temp) < 0) ? -1 : 1;
                                    vQtyDift = Math.Abs(vInventory_Temp - vStockQTY_Temp);
                                    vSQLStrTemp = "select MAX(Items) MaxItem from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                                    vSQLStrTemp = "select MAX(Items) MaxItem from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                                    vIndex_Count = Int32.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "MaxItem"), out vTempINT) ? vTempINT + 1 : 1;
                                    vItems = vIndex_Count.ToString("D4").Trim();
                                    vSheetNoItems = vSheetNo + vItems;
                                    vRemarkB = DateTime.Today.ToShortDateString() + " 盤點數量更新" + Environment.NewLine;
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
                                        dsTempB.InsertParameters.Add(new Parameter("Quantity", DbType.Int32, vQtyDift.ToString()));
                                        dsTempB.InsertParameters.Add(new Parameter("RemarkB", DbType.String, vRemarkB));
                                        dsTempB.InsertParameters.Add(new Parameter("QtyMode", DbType.Int32, vQtyMode.ToString()));
                                        dsTempB.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                        try
                                        {
                                            dsTempB.Insert();
                                            string vRecordNote = "匯入庫存量" + Environment.NewLine +
                                                                 "IAConsumables.aspx";
                                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                                        }
                                        catch (Exception eMessage)
                                        {
                                            Response.Write("<Script language='Javascript'>");
                                            Response.Write("alert('" + eMessage.Message + "')");
                                            Response.Write("</" + "Script>");
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('查無資料。')");
                            Response.Write("</" + "Script>");
                        }
                        break;

                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇盤點結果檔')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 離開
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 關閉報表預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
            plShowData.Visible = true;
        }

        protected void gvShowList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gvShowList.PageIndex = e.NewPageIndex;
            gvShowList.DataBind();
        }

        /// <summary>
        /// 新增耗材
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbInsert_Click(object sender, EventArgs e)
        {
            Session["ConsMode"] = "INS";
            SetInputMode(true);
            foreach (Control cTemp in plShowData.Controls)
            {
                if (cTemp is TextBox)
                {
                    (cTemp as TextBox).Text = "";
                }
                else if (cTemp is CheckBox)
                {
                    (cTemp as CheckBox).Checked = false;
                }
                else if ((cTemp is DropDownList) && ((cTemp as DropDownList).Items.Count > 0))
                {
                    (cTemp as DropDownList).SelectedIndex = 0;
                }
            }
            eIsStopuse.Text = "N";
            eInOrder.Text = "N";
            eBuDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
        }

        /// <summary>
        /// 修改耗材資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbEdit_Click(object sender, EventArgs e)
        {
            Session["ConsMode"] = "EDIT";
            SetInputMode(true);
        }

        /// <summary>
        /// 刪除耗材資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vConsNo = eConsNo.Text.Trim();
            if (vConsNo != "")
            {
                string vTempStr = "delete from Consumables where ConsNo = @ConsNo";
                using (SqlConnection connDEL = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdDEL = new SqlCommand(vTempStr, connDEL);
                    connDEL.Open();
                    cmdDEL.Parameters.Clear();
                    cmdDEL.Parameters.Add(new SqlParameter("ConsNo", vConsNo));
                    cmdDEL.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 設定停用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbStopUse_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vConsNo = eConsNo.Text.Trim();
            if (vConsNo != "")
            {
                string vTempStr = "update Consumables set IsStopuse = 'Y' where ConsNo = @ConsNo";
                using (SqlConnection connStopUse = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdStopUse = new SqlCommand(vTempStr, connStopUse);
                    connStopUse.Open();
                    cmdStopUse.Parameters.Clear();
                    cmdStopUse.Parameters.Add(new SqlParameter("ConsNo", vConsNo));
                    cmdStopUse.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 變更選擇
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gvShowList_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vConsNo = gvShowList.SelectedDataKey.ToString().Trim();
            if (vConsNo != "")
            {
                GetSelectData(vConsNo);
            }
        }

        private void GetSelectData(string fConsNo)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "select ConsNo, ConsName, ConsType, ConsUnit, ConsColor, ConsSpec, ConsSpec2, Brand, " + Environment.NewLine +
                                "       StockQty, AVGPrice, LastInDate, LastOutDate, StoreLocation, IsStopUse, IsInorder, LastInPrice, " + Environment.NewLine +
                                "       BuDate, BuMan, e1.[Name] as BuManName, ModifyDate, ModifyMan, e2.[Name] as ModifyManName, Remark " + Environment.NewLine +
                                "  from Consumables c left join Employee e1 on e1.EmpNo = c.BuMan " + Environment.NewLine +
                                "                     left join Employee e2 on e2.EmpNo = c.ModifyMan " + Environment.NewLine +
                                " where ConsNo = @ConsNo ";
            double vTempNum = 0;
            DateTime vTempDate;
            using (SqlConnection connGetData = new SqlConnection(vConnStr))
            {
                SqlCommand cmdSelectData = new SqlCommand(vSelectStr, connGetData);
                connGetData.Open();
                cmdSelectData.Parameters.Clear();
                cmdSelectData.Parameters.Add(new SqlParameter("ConsNo", fConsNo));
                SqlDataReader drSelectData = cmdSelectData.ExecuteReader();
                if (drSelectData.HasRows)
                {
                    while (drSelectData.Read())
                    {
                        vdConsNo = drSelectData["ConsNo"].ToString().Trim();
                        vdConsName = drSelectData["ConsName"].ToString().Trim();
                        vdConsType = drSelectData["ConsType"].ToString().Trim();
                        ddlConsType.SelectedIndex = ddlConsType.Items.IndexOf(ddlConsType.Items.FindByValue(vdConsType));
                        vdConsUnit = drSelectData["ComsUnit"].ToString().Trim();
                        ddlConsUnit.SelectedIndex = ddlConsUnit.Items.IndexOf(ddlConsUnit.Items.FindByValue(vdConsUnit));
                        vdConsColor = drSelectData["ConsColor"].ToString().Trim();
                        vdConsSpec = drSelectData["ConsSpec"].ToString().Trim();
                        vdConsSpec2 = drSelectData["ConsSpec2"].ToString().Trim();
                        vdBrand = drSelectData["Brand"].ToString().Trim();
                        ddlBrand.SelectedIndex = ddlBrand.Items.IndexOf(ddlBrand.Items.FindByValue(vdBrand));
                        vdStockQty = (double.TryParse(drSelectData["StockQty"].ToString().Trim(), out vTempNum)) ? vTempNum.ToString() : "0";
                        vdAvgPrice = (double.TryParse(drSelectData["AVGPrice"].ToString().Trim(), out vTempNum)) ? vTempNum.ToString() : "0";
                        vdLastInPrice = (double.TryParse(drSelectData["LastInPrice"].ToString().Trim(), out vTempNum)) ? vTempNum.ToString() : "0";
                        vdLastInDate = (DateTime.TryParse(drSelectData["LastInDate"].ToString().Trim(), out vTempDate)) ? vTempDate.ToString("yyyy/MM/dd") : String.Empty;
                        vdLastOutDate = (DateTime.TryParse(drSelectData["LastOutDate"].ToString().Trim(), out vTempDate)) ? vTempDate.ToString("yyyy/MM/dd") : String.Empty;
                        vdBuDate = (DateTime.TryParse(drSelectData["BuDate"].ToString().Trim(), out vTempDate)) ? vTempDate.ToString("yyyy/MM/dd") : String.Empty;
                        vdModifyDate = (DateTime.TryParse(drSelectData["ModifyDate"].ToString().Trim(), out vTempDate)) ? vTempDate.ToString("yyyy/MM/dd") : String.Empty;
                        vdStoreLocation = drSelectData["StoreLocation"].ToString().Trim();
                        vdIsStopUse = drSelectData["IsStopUse"].ToString().Trim();
                        cbIsStopuse.Checked = (vdIsStopUse.ToUpper() == "Y");
                        vdInOrder = drSelectData["IsInorder"].ToString().Trim();
                        cbInOrder.Checked = (vdInOrder.ToUpper() == "Y");
                        vdBuMan = drSelectData["BuManName"].ToString().Trim();
                        vdModifyMan = drSelectData["ModifyManName"].ToString().Trim();
                        vdRemark = drSelectData["Remark"].ToString().Trim();
                    }
                }
                Page.DataBind();
            }
        }

        /// <summary>
        /// 新增資料
        /// </summary>
        private void InsertData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            int vIndexINT = 0;
            string vIndexStr = "";
            string vOldConsNo = "";

            string vConsNo_First = eConsType.Text.Trim().ToUpper() + "-" + eBrand.Text.Trim().ToUpper() + "-";
            vOldConsNo = PF.GetValue(vConnStr, "select top 1 ConsNo from Consumables where ConsNo like '" + vConsNo_First + "%' order by ConsNo DESC", "ConsNo");
            vIndexStr = Int32.TryParse(vOldConsNo.Replace(vConsNo_First, ""), out vIndexINT) ? (vIndexINT + 1).ToString("D4") : "0001";

            string vConsNo_INS = vConsNo_First + vIndexStr;
            string vConsName_INS = eConsName.Text.Trim();
            string vConsType_INS = eConsType.Text.Trim();
            string vConsUnit_INS = eConsUnit.Text.Trim();
            string vConsColor_INS = eConsColor.Text.Trim();
            string vConsSpec_INS = eConsSpec.Text.Trim();
            string vConsSpec2_INS = eConsSpec2.Text.Trim();
            string vBrand_INS = eBrand.Text.Trim();
            string vStoreLocation_INS = eStoreLocation.Text.Trim();
            string vIsStopUse_INS = (eIsStopuse.Text.Trim().ToUpper() == "Y") ? "Y" : "N";
            string vIsInorder_INS = (eInOrder.Text.Trim().ToUpper() == "Y") ? "Y" : "N";
            string vRemark_INS = eRemark.Text.Trim();

            string vInsertStr = "insert into Consumables " + Environment.NewLine +
                                "       (ConsNo, ConsName, ConsType, ConsUnit, ConsColor, ConsSpec, ConsSpec2, Brand, StockQty, AVGPrice, " + Environment.NewLine +
                                "        LastInDate, LastOutDate, StoreLocation, IsStopUse, IsInorder, LastInPrice, BuDate, BuMan, Remark)" + Environment.NewLine +
                                "values (@ConsNo, @ConsName, @ConsType, @ConsUnit, @ConsColor, @ConsSpec, @ConsSpec2, @Brand, 0.0, 0.0, " + Environment.NewLine +
                                "        NULL, NULL, @StoreLocation, @IsStopUse, @IsInorder, 0.0, GetDate(), @BuMan, @Remark)";
            using (SqlConnection connINS = new SqlConnection(vConnStr))
            {
                SqlCommand cmdINS = new SqlCommand(vInsertStr, connINS);
                connINS.Open();
                cmdINS.Parameters.Clear();
                cmdINS.Parameters.Add(new Parameter("ConsNo", DbType.String, vConsNo_INS));
                cmdINS.Parameters.Add(new Parameter("ConsName", DbType.String, vConsName_INS));
                cmdINS.Parameters.Add(new Parameter("ConsType", DbType.String, vConsType_INS));
                cmdINS.Parameters.Add(new Parameter("ConsUnit", DbType.String, vConsUnit_INS));
                cmdINS.Parameters.Add(new Parameter("ConsColor", DbType.String, vConsColor_INS));
                cmdINS.Parameters.Add(new Parameter("ConsSpec", DbType.String, vConsSpec_INS));
                cmdINS.Parameters.Add(new Parameter("ConsSpec2", DbType.String, vConsSpec2_INS));
                cmdINS.Parameters.Add(new Parameter("Brand", DbType.String, vBrand_INS));
                cmdINS.Parameters.Add(new Parameter("StoreLocation", DbType.String, vStoreLocation_INS));
                cmdINS.Parameters.Add(new Parameter("IsStopUse", DbType.String, vIsStopUse_INS));
                cmdINS.Parameters.Add(new Parameter("IsInorder", DbType.String, vIsInorder_INS));
                cmdINS.Parameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                cmdINS.Parameters.Add(new Parameter("Remark", DbType.String, vRemark_INS));
                cmdINS.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 修改資料
        /// </summary>
        private void UpdateData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (eConsNo.Text.Trim() != "")
            {
                string vUpdateStr = "update Consumables set ConsName = @ConsName, ConsUnit = @ConsUnit, ConsColor = @ConsColor, ConsSpec = @ConsSpec, ConsSpec2 = @ConsSpec2, " + Environment.NewLine +
                                    "       Brand = @Brand, StoreLocation = @StoreLocation, IsStopUse = @IsStopUse, IsInorder = @IsInorder, Remark = @Remark,  " + Environment.NewLine +
                                    "       ModifyDate = GetDate(), ModifyMan = @ModifyMan " + Environment.NewLine +
                                    " where ConsNo = @ConsNo ";
                string vConsNo_Edit = eConsNo.Text.Trim();
                string vConsName_Edit = eConsName.Text.Trim();
                string vConsUnit_Edit = eConsUnit.Text.Trim();
                string vConsColor_Edit = eConsColor.Text.Trim();
                string vConsSpec_Edit = eConsSpec.Text.Trim();
                string vConsSpec2_Edit = eConsSpec2.Text.Trim();
                string vBrand_Edit = eBrand.Text.Trim();
                string vStoreLocation_Edit = eStoreLocation.Text.Trim();
                string vIsStopUse_Edit = (eIsStopuse.Text.Trim().ToUpper() == "Y") ? "Y" : "N";
                string vIsInorder_Edit = (eInOrder.Text.Trim().ToUpper() == "Y") ? "Y" : "N";
                string vRemark_Edit = eRemark.Text.Trim();
                using (SqlConnection connUpdate = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdUpdate = new SqlCommand(vUpdateStr, connUpdate);
                    connUpdate.Open();
                    cmdUpdate.Parameters.Clear();
                    cmdUpdate.Parameters.Add(new Parameter("ConsName", DbType.String, vConsName_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("ConsUnit", DbType.String, vConsUnit_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("ConsColor", DbType.String, vConsColor_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("ConsSpec", DbType.String, vConsSpec_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("ConsSpec2", DbType.String, vConsSpec2_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("Brand", DbType.String, vBrand_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("StoreLocation", DbType.String, vStoreLocation_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("IsStopUse", DbType.String, vIsStopUse_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("IsInorder", DbType.String, vIsInorder_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("Remark", DbType.String, vRemark_Edit));
                    cmdUpdate.Parameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    cmdUpdate.Parameters.Add(new Parameter("ConsNo", DbType.String, vConsNo_Edit));
                    cmdUpdate.ExecuteNonQuery();
                }
                CalRealPriceAndQty(vConsNo_Edit);
            }
        }

        /// <summary>
        /// 重新計算指定料件庫存量及平均單價
        /// </summary>
        private void CalRealPriceAndQty(string fConsNo)
        {
            if (fConsNo != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSelectSumStr = "select sum(Quantity * QtyMode) as TotalQty from ConsSheetB where ConsNo = '" + fConsNo + "' ";
                string vTotalQty = PF.GetValue(vConnStr, vSelectSumStr, "TotalQty");
                vSelectSumStr = "select sum(Quantity * Price) / sum(Quantity) as AVGPrice " + Environment.NewLine +
                                "  from ConsSheetB " + Environment.NewLine +
                                " where ConsNo = '" + fConsNo + "' " + Environment.NewLine +
                                "   and SheetNo in (select SheetNo from ConsSheetA where SheetMode = 'SI')";
                string vAVGPrice = PF.GetValue(vConnStr, vSelectSumStr, "AVGPrice");
                vSelectSumStr = "update Consumables set StockQty = " + vTotalQty + ", AvgPrice = " + vAVGPrice + " where ConsNo = '" + fConsNo + "' ";
                PF.ExecSQL(vConnStr, vSelectSumStr);
            }
        }

        /// <summary>
        /// 確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Click(object sender, EventArgs e)
        {
            switch (Session["ConsMode"].ToString().Trim())
            {
                case "INS":
                    InsertData();
                    break;
                case "EDIT":
                    UpdateData();
                    break;
            }
            Session["ConsMode"] = "LIST";
            SetInputMode(false);
            gvShowList.DataBind();
        }

        /// <summary>
        /// 取消
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCancel_Click(object sender, EventArgs e)
        {
            Session["ConsMode"] = "LIST";
            string vConsNo = (gvShowList.SelectedDataKey != null) ? gvShowList.SelectedDataKey.ToString().Trim() : "";
            if (vConsNo != "")
            {
                GetSelectData(vConsNo);
            }
            SetInputMode(false);
            gvShowList.DataBind();
        }

        /// <summary>
        /// 選擇廠商
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlBrand_SelectedIndexChanged(object sender, EventArgs e)
        {
            eBrand.Text = ddlBrand.SelectedValue.Trim();
        }
    }
}