using Amaterasu_Function;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class CarDeviceList : System.Web.UI.Page
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
                        eDepNo_S.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));
                        eDepNo_E.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));
                        eDepNo_S.Text = ((Int32.Parse(vLoginDepNo) >= 11) && (Int32.Parse(vLoginDepNo) <= 30)) ? vLoginDepNo : "";
                    }
                    else
                    {
                        CarDeviceDataBind();
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

        private string GetSelStr(string fSource)
        {
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.CompanyNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and a.CompanyNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.CompanyNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Tran_Type = (eTran_Type.Text.Trim() != "") ? "   and a.Tran_Type = '" + eTran_Type.Text.Trim() + "' " + Environment.NewLine : "";
            string vSelectStr = "";
            if (fSource.ToUpper() == "GRID")
            {
                vSelectStr = "select a.Car_No, a.Car_ID, a.CompanyNo, (select [Name] from Department where DepNo = a.CompanyNo) CompanyName, " + Environment.NewLine +
                             "       a.ProdDate, a.GetLicDate, a.Car_TypeID, a.Canon, " + Environment.NewLine +
                             "       a.Car_Class, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and ClassNo = a.Car_Class) Car_Class_C, " + Environment.NewLine +
                             "       a.Class, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CLASS' and ClassNo = a.Class) Class_C, " + Environment.NewLine +
                             "       a.Point,  (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       POINT' and ClassNo = a.Point) Point_C, " + Environment.NewLine +
                             "       a.CarLevel, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CARLEVEL' and ClassNo = a.CarLevel) CarLevel_C, " + Environment.NewLine +
                             "       a.Exceptional, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       EXCEPTIONAL' and ClassNo = a.Exceptional) Exceptional_C, " + Environment.NewLine +
                             "       a.Tran_Type, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       TRAN_TYPE' and ClassNo = a.Tran_Type) Tran_Type_C, " + Environment.NewLine +
                             "       a.Tran_Date, a.SitQty, a.StandQty, a.BodyClass, a.HorsePower, a.FreezeType, a.Fix3Date, a.NextEDate, a.NCheckTerm, " + Environment.NewLine +
                             "       a.StopDate, a.CancelDate, a.RDDate, a.SaleDate, a.Reday, a.Remark " + Environment.NewLine +
                             "  from Car_InfoA a " + Environment.NewLine +
                             " where (1 = 1) " + Environment.NewLine +
                             vWStr_DepNo + vWStr_Tran_Type +
                             " order by a.CompanyNo, a.Car_ID ";
            }
            else if (fSource.ToUpper() == "EXCEL")
            {
                vSelectStr = "select a.Car_No, a.Car_ID, a.CompanyNo, (select [Name] from Department where DepNo = a.CompanyNo) CompanyName, " + Environment.NewLine +
                             "       a.ProdDate, a.GetLicDate, a.Car_TypeID, a.Canon, " + Environment.NewLine +
                             "       a.Car_Class, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and ClassNo = a.Car_Class) Car_Class_C, " + Environment.NewLine +
                             "       a.Class, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CLASS' and ClassNo = a.Class) Class_C, " + Environment.NewLine +
                             "       a.Point,  (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       POINT' and ClassNo = a.Point) Point_C, " + Environment.NewLine +
                             "       a.CarLevel, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CARLEVEL' and ClassNo = a.CarLevel) CarLevel_C, " + Environment.NewLine +
                             "       a.Exceptional, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       EXCEPTIONAL' and ClassNo = a.Exceptional) Exceptional_C, " + Environment.NewLine +
                             "       a.Tran_Type, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       TRAN_TYPE' and ClassNo = a.Tran_Type) Tran_Type_C, " + Environment.NewLine +
                             "       a.Tran_Date, a.SitQty, a.StandQty, a.BodyClass, a.HorsePower, a.FreezeType, a.Fix3Date, a.NextEDate, a.NCheckTerm, " + Environment.NewLine +
                             "       a.StopDate, a.CancelDate, a.RDDate, a.SaleDate, a.Reday, a.Remark, " + Environment.NewLine +
                             "       b.CK1, b.CK2, b.CK3, b.CK4, b.CK5, b.CK6, b.CK7, b.CK8, b.CK9, b.CK10, " + Environment.NewLine +
                             "       b.CK11, b.CK12, b.CK13, b.CK14, b.CK15, b.CK16, b.CK17, b.CK18, b.CK19, b.CK20, " + Environment.NewLine +
                             "       b.CK21, b.CK22, b.CK23, b.CK24, b.CK25, b.CK26, b.CK27, b.CK28, b.CK29, b.CK30, " + Environment.NewLine +
                             "       b.CK31, b.CK32, b.CK33, b.CK34, b.CK35, b.CK36, b.CK37, b.CK38, b.CK39, b.CK40, b.Remark40, " + Environment.NewLine +
                             "       b.Qty41, b.Unit41, b.Brand41, b.Spec41, b.Mac41, b.Date41, b.Price41, b.Remark41, " + Environment.NewLine +
                             "       b.Qty42, b.Unit42, b.Brand42, b.Spec42, b.Mac42, b.Date42, b.Price42, b.Remark42, " + Environment.NewLine +
                             "       b.Qty43, b.Unit43, b.Brand43, b.Spec43, b.Mac43, b.Date43, b.Price43, b.Remark43, " + Environment.NewLine +
                             "       b.Qty44, b.Unit44, b.Brand44, b.Spec44, b.Mac44, b.Date44, b.Price44, b.Remark44, " + Environment.NewLine +
                             "       b.Qty45, b.Unit45, b.Brand45, b.Spec45, b.Mac45, b.Date45, b.Price45, b.Remark45, " + Environment.NewLine +
                             "       b.Qty46, b.Unit46, b.Brand46, b.Spec46, b.Mac46, b.Date46, b.Price46, b.Remark46, " + Environment.NewLine +
                             "       b.Qty47, b.Unit47, b.Brand47, b.Spec47, b.Mac47, b.Date47, b.Price47, b.Remark47, " + Environment.NewLine +
                             "       b.Qty48, b.Unit48, b.Brand48, b.Spec48, b.Mac48, b.Date48, b.Price48, b.Remark48, " + Environment.NewLine +
                             "       b.Qty49, b.Unit49, b.Brand49, b.Spec49, b.Mac49, b.Date49, b.Price49, b.Remark49, " + Environment.NewLine +
                             "       b.Qty50, b.Unit50, b.Brand50, b.Spec50, b.Mac50, b.Date50, b.Price50, b.Remark50, " + Environment.NewLine +
                             "       b.Qty51, b.Unit51, b.Brand51, b.Spec51, b.Mac51, b.Date51, b.Price51, b.Remark51, " + Environment.NewLine +
                             "       b.Qty52, b.Unit52, b.Brand52, b.Spec52, b.Mac52, b.Date52, b.Price52, b.Remark52, " + Environment.NewLine +
                             "       b.Qty53, b.Unit53, b.Brand53, b.Spec53, b.Mac53, b.Date53, b.Price53, b.Remark53, " + Environment.NewLine +
                             "       b.Qty54, b.Unit54, b.Brand54, b.Spec54, b.Mac54, b.Date54, b.Price54, b.Remark54, " + Environment.NewLine +
                             "       b.Qty55, b.Unit55, b.Brand55, b.Spec55, b.Mac55, b.Date55, b.Price55, b.Remark55, " + Environment.NewLine +
                             "       b.Qty56, b.Unit56, b.Brand56, b.Spec56, b.Mac56, b.Date56, b.Price56, b.Remark56, " + Environment.NewLine +
                             "       b.Qty57, b.Unit57, b.Brand57, b.Spec57, b.Mac57, b.Date57, b.Price57, b.Remark57, " + Environment.NewLine +
                             "       b.Qty58, b.Unit58, b.Brand58, b.Spec58, b.Mac58, b.Date58, b.Price58, b.Remark58, " + Environment.NewLine +
                             "       b.Qty59, b.Unit59, b.Brand59, b.Spec59, b.Mac59, b.Date59, b.Price59, b.Remark59, " + Environment.NewLine +
                             "       b.Qty60, b.Unit60, b.Brand60, b.Spec60, b.Mac60, b.Date60, b.Price60, b.Remark60, " + Environment.NewLine +
                             "       b.Qty61, b.Unit61, b.Brand61, b.Spec61, b.Mac61, b.Date61, b.Price61, b.Remark61, " + Environment.NewLine +
                             "       b.Qty62, b.Unit62, b.Brand62, b.Spec62, b.Mac62, b.Date62, b.Price62, b.Remark62, " + Environment.NewLine +
                             "       b.Qty63, b.Unit63, b.Brand63, b.Spec63, b.Mac63, b.Date63, b.Price63, b.Remark63, " + Environment.NewLine +
                             "       b.Qty64, b.Unit64, b.Brand64, b.Spec64, b.Mac64, b.Date64, b.Price64, b.Remark64, " + Environment.NewLine +
                             "       b.Qty65, b.Unit65, b.Brand65, b.Spec65, b.Mac65, b.Date65, b.Price65, b.Remark65, " + Environment.NewLine +
                             "       b.Qty66, b.Unit66, b.Brand66, b.Spec66, b.Mac66, b.Date66, b.Price66, b.Remark66, " + Environment.NewLine +
                             "       b.Qty67, b.Unit67, b.Brand67, b.Spec67, b.Mac67, b.Date67, b.Price67, b.Remark67, " + Environment.NewLine +
                             "       b.Qty68, b.Unit68, b.Brand68, b.Spec68, b.Mac68, b.Date68, b.Price68, b.Remark68, " + Environment.NewLine +
                             "       b.Qty69, b.Unit69, b.Brand69, b.Spec69, b.Mac69, b.Date69, b.Price69, b.Remark69, " + Environment.NewLine +
                             "       b.Qty70, b.Unit70, b.Brand70, b.Spec70, b.Mac70, b.Date70, b.Price70, b.Remark70 " + Environment.NewLine +
                             "  from Car_InfoA a left join EQUIP b on b.Car_ID = a.Car_ID " + Environment.NewLine +
                             " where (1 = 1) " + Environment.NewLine +
                             vWStr_DepNo + vWStr_Tran_Type +
                             " order by a.CompanyNo, a.Car_ID ";
            }
            return vSelectStr;
        }

        private void CarDeviceDataBind()
        {
            string vSelectStr = GetSelStr("GRID");
            sdsCarDeviceList.SelectCommand = vSelectStr;
            gridCarDeviceList.DataBind();
        }

        protected void ddlTran_Type_SelectedIndexChanged(object sender, EventArgs e)
        {
            eTran_Type.Text = ddlTran_Type.SelectedValue;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            CarDeviceDataBind();
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            //準備匯出EXCEL
            XSSFWorkbook wbExcel = new XSSFWorkbook();
            //Excel 工作表
            XSSFSheet wsExcel;

            //設定標題欄位的格式
            XSSFCellStyle csTitle = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中

            //設定標題欄位的字體格式
            XSSFFont fontTitle = (XSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 12; //字體大小
            csTitle.SetFont(fontTitle);

            //設定資料內容欄位的格式
            XSSFCellStyle csData = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            XSSFFont fontData = (XSSFFont)wbExcel.CreateFont();
            fontData.FontHeightInPoints = 12;
            csData.SetFont(fontData);

            XSSFCellStyle csData_Red = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            XSSFFont fontData_Red = (XSSFFont)wbExcel.CreateFont();
            fontData_Red.FontHeightInPoints = 12; //字體大小
            fontData_Red.Color = IndexedColors.Red.Index; //字體顏色
            csData_Red.SetFont(fontData_Red);

            XSSFCellStyle csData_Int = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            XSSFDataFormat format = (XSSFDataFormat)wbExcel.CreateDataFormat();
            csData_Int.DataFormat = format.GetFormat("###,##0");

            string vHeaderText = "";

            string vFileName = "車輛配備清冊";
            int vLinesNo = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelStr("EXCEL");
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    DateTime vBuDate;
                    int vFieldIndex = 0;
                    string vSQLStr = "";
                    //新增一個工作表
                    wsExcel = (XSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        if ((drExcel.GetName(i).Substring(0, 2).ToUpper() == "CK") && Int32.TryParse(drExcel.GetName(i).Substring(2), out vFieldIndex))
                        {
                            vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_CK" + vFieldIndex.ToString() + "' ";
                            vHeaderText = PF.GetValue(vConnStr, vSQLStr, "Content");
                        }
                        else if ((drExcel.GetName(i).Substring(0, 3).ToUpper() == "QTY") && Int32.TryParse(drExcel.GetName(i).Substring(3), out vFieldIndex))
                        {
                            vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_LB" + vFieldIndex.ToString() + "' ";
                            vHeaderText = PF.GetValue(vConnStr, vSQLStr, "Content") + "數量";
                        }
                        else if ((drExcel.GetName(i).Substring(0, 4).ToUpper() == "UNIT") && Int32.TryParse(drExcel.GetName(i).Substring(4), out vFieldIndex))
                        {
                            vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_LB" + vFieldIndex.ToString() + "' ";
                            vHeaderText = PF.GetValue(vConnStr, vSQLStr, "Content") + "單位";
                        }
                        else if ((drExcel.GetName(i).Substring(0, 5).ToUpper() == "BRAND") && Int32.TryParse(drExcel.GetName(i).Substring(5), out vFieldIndex))
                        {
                            vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_LB" + vFieldIndex.ToString() + "' ";
                            vHeaderText = PF.GetValue(vConnStr, vSQLStr, "Content") + "廠牌";
                        }
                        else if ((drExcel.GetName(i).Substring(0, 4).ToUpper() == "SPEC") && Int32.TryParse(drExcel.GetName(i).Substring(4), out vFieldIndex))
                        {
                            vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_LB" + vFieldIndex.ToString() + "' ";
                            vHeaderText = PF.GetValue(vConnStr, vSQLStr, "Content") + "型號";
                        }
                        else if ((drExcel.GetName(i).Substring(0, 3).ToUpper() == "MAC") && Int32.TryParse(drExcel.GetName(i).Substring(3), out vFieldIndex))
                        {
                            vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_LB" + vFieldIndex.ToString() + "' ";
                            vHeaderText = PF.GetValue(vConnStr, vSQLStr, "Content") + "機號";
                        }
                        else if ((drExcel.GetName(i).Substring(0, 4).ToUpper() == "DATE") && Int32.TryParse(drExcel.GetName(i).Substring(4), out vFieldIndex))
                        {
                            vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_LB" + vFieldIndex.ToString() + "' ";
                            vHeaderText = PF.GetValue(vConnStr, vSQLStr, "Content") + "啟用日";
                        }
                        else if ((drExcel.GetName(i).Substring(0, 5).ToUpper() == "PRICE") && Int32.TryParse(drExcel.GetName(i).Substring(5), out vFieldIndex))
                        {
                            vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_LB" + vFieldIndex.ToString() + "' ";
                            vHeaderText = PF.GetValue(vConnStr, vSQLStr, "Content") + "單價";
                        }
                        else if ((drExcel.GetName(i).Trim().Length > 6) &&
                                 (drExcel.GetName(i).Substring(0, 6).ToUpper() == "REMARK") &&
                                 Int32.TryParse(drExcel.GetName(i).Substring(6), out vFieldIndex))
                        {
                            vSQLStr = "select Content from SysFlag where FormName = 'unSysFlag' and ControlItem = 'a_LB" + vFieldIndex.ToString() + "' ";
                            vHeaderText = PF.GetValue(vConnStr, vSQLStr, "Content") + "備註";
                        }
                        else
                        {
                            vHeaderText = (drExcel.GetName(i).Trim() == "Car_No") ? "車輛代號" :
                                          (drExcel.GetName(i).Trim() == "Car_ID") ? "牌照號碼" :
                                          (drExcel.GetName(i).Trim() == "CompanyNo") ? "單位代碼" :
                                          (drExcel.GetName(i).Trim() == "CompanyName") ? "單位" :
                                          (drExcel.GetName(i).Trim() == "ProdDate") ? "出廠日" :
                                          (drExcel.GetName(i).Trim() == "GetLicDate") ? "領照日" :
                                          (drExcel.GetName(i).Trim() == "Car_TypeID") ? "型式" :
                                          (drExcel.GetName(i).Trim() == "Canon") ? "污染排放標準" :
                                          (drExcel.GetName(i).Trim() == "Car_Class") ? "廠牌代號" :
                                          (drExcel.GetName(i).Trim() == "Car_Class_C") ? "廠牌" :
                                          (drExcel.GetName(i).Trim() == "Class") ? "類別代號" :
                                          (drExcel.GetName(i).Trim() == "Class_C") ? "類別" :
                                          (drExcel.GetName(i).Trim() == "Point") ? "用途代號" :
                                          (drExcel.GetName(i).Trim() == "Point_C") ? "用途" :
                                          (drExcel.GetName(i).Trim() == "CarLevel") ? "車輛等級代號" :
                                          (drExcel.GetName(i).Trim() == "CarLevel_C") ? "車輛等級" :
                                          (drExcel.GetName(i).Trim() == "Exceptional") ? "特殊車種代號" :
                                          (drExcel.GetName(i).Trim() == "Exceptional_C") ? "特殊車種" :
                                          (drExcel.GetName(i).Trim() == "Tran_Type") ? "狀態代號" :
                                          (drExcel.GetName(i).Trim() == "Tran_Type_C") ? "狀態" :
                                          (drExcel.GetName(i).Trim() == "Tran_Date") ? "異動日" :
                                          (drExcel.GetName(i).Trim() == "SitQty") ? "座位數" :
                                          (drExcel.GetName(i).Trim() == "StandQty") ? "站位數" :
                                          (drExcel.GetName(i).Trim() == "BodyClass") ? "車體廠牌" :
                                          (drExcel.GetName(i).Trim() == "HorsePower") ? "馬力" :
                                          (drExcel.GetName(i).Trim() == "FreezeType") ? "冷氣廠牌" :
                                          (drExcel.GetName(i).Trim() == "Fix3Date") ? "下次三級日" :
                                          (drExcel.GetName(i).Trim() == "NextEDate") ? "下次檢驗日" :
                                          (drExcel.GetName(i).Trim() == "NCheckTerm") ? "檢驗期限" :
                                          (drExcel.GetName(i).Trim() == "StopDate") ? "停駛日" :
                                          (drExcel.GetName(i).Trim() == "CancelDate") ? "繳銷日" :
                                          (drExcel.GetName(i).Trim() == "RDDate") ? "補照日" :
                                          (drExcel.GetName(i).Trim() == "SaleDate") ? "出售日" :
                                          (drExcel.GetName(i).Trim() == "REDay") ? "重新領牌日" :
                                          (drExcel.GetName(i).Trim() == "Remark") ? "備註" : drExcel.GetName(i).Trim();
                        }
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
                            if (((drExcel.GetName(i).Trim() == "ProdDate") ||
                                 (drExcel.GetName(i).Trim() == "GetLicDate") ||
                                 (drExcel.GetName(i).Trim() == "Tran_Date") ||
                                 (drExcel.GetName(i).Trim() == "Fix3Date") ||
                                 (drExcel.GetName(i).Trim() == "NextEDate") ||
                                 (drExcel.GetName(i).Trim() == "NCheckTerm") ||
                                 (drExcel.GetName(i).Trim() == "StopDate") ||
                                 (drExcel.GetName(i).Trim() == "CancelDate") ||
                                 (drExcel.GetName(i).Trim() == "RDDate") ||
                                 (drExcel.GetName(i).Trim() == "SaleDate") ||
                                 (drExcel.GetName(i).Trim() == "REDate")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString().Trim());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue((vBuDate.Year - 1911).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i) == "SitQty") ||
                                      (drExcel.GetName(i) == "StandQty")) && (drExcel[i].ToString() != ""))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
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
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;

                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xlsx", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xlsx"));
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