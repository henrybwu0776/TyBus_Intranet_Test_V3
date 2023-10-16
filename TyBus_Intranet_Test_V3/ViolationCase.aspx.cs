using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ViolationCase : System.Web.UI.Page
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
        private string vStationNo = "";
        private DateTime vToday;

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

                UnobtrusiveValidationMode = System.Web.UI.UnobtrusiveValidationMode.None;

                if (vLoginID != "")
                {
                    vToday = DateTime.Today;
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vViolationDateURL_Search = "InputDate_ChineseYears.aspx?TextboxID=" + eViolationDate_Start_Search.ClientID;
                    string vViolationDateScript_Search = "window.open('" + vViolationDateURL_Search + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eViolationDate_Start_Search.Attributes["onClick"] = vViolationDateScript_Search;
                    vViolationDateURL_Search = "InputDate_ChineseYears.aspx?TextboxID=" + eViolationDate_End_Search.ClientID;
                    vViolationDateScript_Search = "window.open('" + vViolationDateURL_Search + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eViolationDate_End_Search.Attributes["onClick"] = vViolationDateScript_Search;

                    if (!IsPostBack)
                    {
                        fuExcel.Visible = (vLoginDepNo == "09");
                        bbImportData.Visible = (vLoginDepNo == "09");
                        vStationNo = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vLoginDepNo + "' and InSHReport = 'V'", "DepNo");

                        eDepNo_Start_Search.Text = vStationNo;
                        eDepNo_End_Search.Text = "";
                        eDepNo_Start_Search.Enabled = (vStationNo == "");
                        eDepNo_End_Search.Enabled = (vStationNo == "");
                    }
                    DataSourceBinded();
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

        private string GetCaseNo(string fCaseNo_TypeA)
        {
            string vResultNo = "";
            string vConnStrTemp = PF.GetConnectionStr(Request.ApplicationPath);
            string vIndex = "";
            string vTempStr = "";
            vTempStr = "select RIGHT(MAX(CaseNo), 3) [Index] from ViolationCase where CaseNo like '" + fCaseNo_TypeA + "%' ";
            vIndex = PF.GetValue(vConnStrTemp, vTempStr, "Index");
            vIndex = (vIndex != "") ? (int.Parse(vIndex) + 1).ToString("D3") : "001";
            vResultNo = fCaseNo_TypeA.Trim() + vIndex;
            return vResultNo;
        }

        /// <summary>
        /// 取回查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            string vWStr_CaseType = (eCaseType_Search.Text.Trim() != "") ? " AND CaseType = '" + eCaseType_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? " AND DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? " AND DepNo like '" + eDepNo_Start_Search.Text.Trim() + "%' " + Environment.NewLine :
                                 ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? " AND DepNo like '" + eDepNo_End_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_CarID = ((eCarID_Start_Search.Text.Trim() != "") && (eCarID_End_Search.Text.Trim() != "")) ? " AND Car_ID between '" + eCarID_Start_Search.Text.Trim() + "' and '" + eCarID_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eCarID_Start_Search.Text.Trim() != "") && (eCarID_End_Search.Text.Trim() == "")) ? " AND Car_ID like '" + eCarID_Start_Search.Text.Trim() + "%' " + Environment.NewLine :
                                 ((eCarID_Start_Search.Text.Trim() == "") && (eCarID_End_Search.Text.Trim() != "")) ? " AND Car_ID like '" + eCarID_End_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_ViolationDate = ((eViolationDate_Start_Search.Text.Trim() != "") && (eViolationDate_End_Search.Text.Trim() != "")) ? " AND ViolationDate between '" + DateTime.Parse(eViolationDate_Start_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' and '" + DateTime.Parse(eViolationDate_End_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start_Search.Text.Trim() != "") && (eViolationDate_End_Search.Text.Trim() == "")) ? " AND ViolationDate = '" + DateTime.Parse(eViolationDate_Start_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                         ((eViolationDate_Start_Search.Text.Trim() == "") && (eViolationDate_End_Search.Text.Trim() != "")) ? " AND ViolationDate = '" + DateTime.Parse(eViolationDate_End_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vResultStr = "SELECT CaseNo, CaseType, " + Environment.NewLine +
                                "       (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = ViolationCase.CaseType) AND (FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, " + Environment.NewLine +
                                "       BuildDate, BuildManName, LinesNo, DepName, Car_ID, DriverName, TicketTitle, PenaltyDep, FineAmount, ViolationDate " + Environment.NewLine +
                                "  FROM ViolationCase " + Environment.NewLine +
                                " WHERE (1 = 1) " + Environment.NewLine +
                                vWStr_CaseType + vWStr_DepNo + vWStr_CarID + vWStr_ViolationDate +
                                " ORDER BY BuildDate DESC, CaseNo";
            return vResultStr;
        }

        /// <summary>
        /// 進行資料繫結
        /// </summary>
        private void DataSourceBinded()
        {
            string vSQLStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = GetSelectStr();
            sdsViolationCase_List.SelectCommand = "";
            sdsViolationCase_List.SelectCommand = vSQLStr;
            gridViolationCaseList.DataSourceID = "sdsViolationCase_List";
            gridViolationCaseList.DataBind();
        }

        protected void ddlCaseType_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eCaseType_Search.Text = ddlCaseType_Search.SelectedValue.Trim();
        }

        protected void bbImportData_Click(object sender, EventArgs e)
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            string vExtName = "";
            string vSheetName = "";
            int SheetCount = 0;

            string vCaseNo = "";
            string vCaseNo_TypeA = "";
            string vCaseType = "";
            string vBuildDate = "";
            string vBuildMan = "";
            string vBuildManName = "";
            string vLinesNo = "";
            string vDepNo = "";
            string vDepName = "";
            string vCar_ID = "";
            string vDriver = "";
            string vDriverName = "";
            string vViolationDate = "";
            string vTicketNo = "";
            string vTicketTitle = "";
            string vPenaltyDep = "";
            string vUndertaker = "";
            string vViolationLocation = "";
            string vViolationRule = "";
            string vViolationNote = "";
            string vFineAmount = "";
            string vPenaltyReason = "";
            string vViolationPoint = "";
            string vPaymentDeadline = "";
            string vPaidDate = "";
            string vRemark = "";
            string vSQLStr = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    case ".xls":
                        //Excel 97-2003
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        SheetCount = wbExcel_H.NumberOfSheets;
                        for (int vSC = 0; vSC < SheetCount; vSC++)
                        {
                            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(vSC);
                            vSheetName = sheetExcel_H.SheetName;
                            vSQLStr = "select ClassNo from DBDICB where ClassTxt = '" + vSheetName + "' and FKey = (CAST('違規記錄檔' AS char(16)) + CAST('ViolationCase' AS char(16)) + 'CaseType')";
                            vCaseType = PF.GetValue(vConnStr, vSQLStr, "ClassNo");
                            if (vCaseType != "")
                            {
                                for (int i = sheetExcel_H.FirstRowNum + 1; i <= sheetExcel_H.LastRowNum; i++)
                                {
                                    HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                    vBuildDate = (vRowTemp_H.GetCell(1).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(1).DateCellValue.ToString("yyyy/MM/dd") : DateTime.Today.ToString("yyyy/MM/dd");
                                    vCaseNo_TypeA = DateTime.Parse(vBuildDate).ToString("yyyyMMdd") + vCaseType;
                                    vCaseNo = GetCaseNo(vCaseNo_TypeA.Trim());
                                    vBuildManName = (vRowTemp_H.GetCell(2).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(2).StringCellValue.Trim() : Session["LoginName"].ToString().Trim();
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vBuildManName + "' and LeaveDay is null";
                                    vBuildMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    vBuildMan = (vBuildMan == "") ? vLoginID : vBuildMan;
                                    vLinesNo = (vRowTemp_H.GetCell(3).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(3).ToString().Trim() : "";
                                    vDepNo = (vRowTemp_H.GetCell(4).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(4).ToString().Trim() : "";
                                    vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo + "'";
                                    vDepName = (vDepNo != "") ? PF.GetValue(vConnStr, vSQLStr, "Name") : "";
                                    vCar_ID = (vRowTemp_H.GetCell(5).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(5).StringCellValue.Trim() : "";
                                    vDriverName = (vRowTemp_H.GetCell(6).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(6).StringCellValue.Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null and Type = '20'";
                                    vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    vViolationDate = (vRowTemp_H.GetCell(7).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(7).DateCellValue.ToString("yyyy/MM/dd") : "";
                                    vTicketNo = (vRowTemp_H.GetCell(8).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(8).StringCellValue.Trim() : "";
                                    vPenaltyDep = (vRowTemp_H.GetCell(9).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(9).StringCellValue.Trim() : "";
                                    vUndertaker = (vRowTemp_H.GetCell(10).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(10).StringCellValue.Trim() : "";
                                    vViolationRule = (vRowTemp_H.GetCell(11).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(11).StringCellValue.Trim() : "";
                                    vFineAmount = (vRowTemp_H.GetCell(12).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(12).NumericCellValue.ToString().Trim() : "0";
                                    vPenaltyReason = (vRowTemp_H.GetCell(13).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(13).StringCellValue.Trim() : "";
                                    vTicketTitle = (vRowTemp_H.GetCell(14).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(14).StringCellValue.Trim() : "";
                                    vViolationLocation = (vRowTemp_H.GetCell(15).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(15).StringCellValue.Trim() : "";
                                    vViolationNote = (vRowTemp_H.GetCell(16).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(16).StringCellValue.Trim() : "";
                                    vViolationPoint = (vRowTemp_H.GetCell(17).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(17).NumericCellValue.ToString().Trim() : "0";
                                    vPaymentDeadline = (vRowTemp_H.GetCell(18).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(18).DateCellValue.ToString("yyyy/MM/dd") : "";
                                    vPaidDate = (vRowTemp_H.GetCell(19).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(19).DateCellValue.ToString("yyyy/MM/dd") : "";
                                    vRemark = (vRowTemp_H.GetCell(20).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(20).ToString().Trim() : "";

                                    sdsImportDataFromExcel.InsertParameters.Clear();
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("CaseType", DbType.String, vCaseType));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, vBuildDate));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vBuildMan));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("BuildManName", DbType.String, vBuildManName));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("DepName", DbType.String, vDepName));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationDate", DbType.Date, vViolationDate));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("TicketNo", DbType.String, vTicketNo));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("TicketTitle", DbType.String, vTicketTitle));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("PenaltyDep", DbType.String, vPenaltyDep));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("Undertaker", DbType.String, vUndertaker));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationLocation", DbType.String, vViolationLocation));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationRule", DbType.String, vViolationRule));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationNote", DbType.String, vViolationNote));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("FineAmount", DbType.Int32, vFineAmount));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("PenaltyReason", DbType.String, vPenaltyReason));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationPoint", DbType.Int32, vViolationPoint));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("PaymentDeadline", DbType.Date, vPaymentDeadline));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("PaidDate", DbType.Date, vPaidDate));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                    sdsImportDataFromExcel.Insert();
                                }
                            }
                        }
                        break;

                    case ".xlsx":
                        //Excel 2010--
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        SheetCount = wbExcel_X.NumberOfSheets;
                        for (int vSC = 0; vSC < SheetCount; vSC++)
                        {
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(vSC);
                            vSheetName = sheetExcel_X.SheetName;
                            vSQLStr = "select ClassNo from DBDICB where ClassTxt = '" + vSheetName + "' and FKey = (CAST('違規記錄檔' AS char(16)) + CAST('ViolationCase' AS char(16)) + 'CaseType')";
                            vCaseType = PF.GetValue(vConnStr, vSQLStr, "ClassNo");
                            if (vCaseType != "")
                            {
                                for (int i = sheetExcel_X.FirstRowNum + 1; i <= sheetExcel_X.LastRowNum; i++)
                                {
                                    XSSFRow vRowTemp_H = (XSSFRow)sheetExcel_X.GetRow(i);
                                    vBuildDate = (vRowTemp_H.GetCell(1).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(1).DateCellValue.ToString("yyyy/MM/dd") : DateTime.Today.ToString("yyyy/MM/dd");
                                    vCaseNo_TypeA = DateTime.Parse(vBuildDate).ToString("yyyyMMdd") + vCaseType;
                                    vCaseNo = GetCaseNo(vCaseNo_TypeA.Trim());
                                    vBuildManName = (vRowTemp_H.GetCell(2).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(2).StringCellValue.Trim() : Session["LoginName"].ToString().Trim();
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vBuildManName + "' and LeaveDay is null";
                                    vBuildMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    vBuildMan = (vBuildMan == "") ? vLoginID : vBuildMan;
                                    vLinesNo = (vRowTemp_H.GetCell(3).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(3).ToString().Trim() : "";
                                    vDepNo = (vRowTemp_H.GetCell(4).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(4).ToString().Trim() : "";
                                    vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo + "'";
                                    vDepName = (vDepNo != "") ? PF.GetValue(vConnStr, vSQLStr, "Name") : "";
                                    vCar_ID = (vRowTemp_H.GetCell(5).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(5).StringCellValue.Trim() : "";
                                    vDriverName = (vRowTemp_H.GetCell(6).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(6).StringCellValue.Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null and Type = '20'";
                                    vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    vViolationDate = (vRowTemp_H.GetCell(7).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(7).DateCellValue.ToString("yyyy/MM/dd") : "";
                                    vTicketNo = (vRowTemp_H.GetCell(8).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(8).StringCellValue.Trim() : "";
                                    vPenaltyDep = (vRowTemp_H.GetCell(9).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(9).StringCellValue.Trim() : "";
                                    vUndertaker = (vRowTemp_H.GetCell(10).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(10).StringCellValue.Trim() : "";
                                    vViolationRule = (vRowTemp_H.GetCell(11).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(11).StringCellValue.Trim() : "";
                                    vFineAmount = (vRowTemp_H.GetCell(12).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(12).NumericCellValue.ToString().Trim() : "0";
                                    vPenaltyReason = (vRowTemp_H.GetCell(13).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(13).StringCellValue.Trim() : "";
                                    vTicketTitle = (vRowTemp_H.GetCell(14).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(14).StringCellValue.Trim() : "";
                                    vViolationLocation = (vRowTemp_H.GetCell(15).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(15).StringCellValue.Trim() : "";
                                    vViolationNote = (vRowTemp_H.GetCell(16).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(16).StringCellValue.Trim() : "";
                                    vViolationPoint = (vRowTemp_H.GetCell(17).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(17).NumericCellValue.ToString().Trim() : "0";
                                    vPaymentDeadline = (vRowTemp_H.GetCell(18).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(18).DateCellValue.ToString("yyyy/MM/dd") : "";
                                    vPaidDate = (vRowTemp_H.GetCell(19).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(19).DateCellValue.ToString("yyyy/MM/dd") : "";
                                    vRemark = (vRowTemp_H.GetCell(20).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(20).ToString().Trim() : "";

                                    sdsImportDataFromExcel.InsertParameters.Clear();
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("CaseType", DbType.String, vCaseType));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, vBuildDate));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vBuildMan));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("BuildManName", DbType.String, vBuildManName));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("DepName", DbType.String, vDepName));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationDate", DbType.Date, vViolationDate));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("TicketNo", DbType.String, vTicketNo));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("TicketTitle", DbType.String, vTicketTitle));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("PenaltyDep", DbType.String, vPenaltyDep));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("Undertaker", DbType.String, vUndertaker));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationLocation", DbType.String, vViolationLocation));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationRule", DbType.String, vViolationRule));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationNote", DbType.String, vViolationNote));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("FineAmount", DbType.Int32, vFineAmount));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("PenaltyReason", DbType.String, vPenaltyReason));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("ViolationPoint", DbType.Int32, vViolationPoint));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("PaymentDeadline", DbType.Date, vPaymentDeadline));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("PaidDate", DbType.Date, vPaidDate));
                                    sdsImportDataFromExcel.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                    sdsImportDataFromExcel.Insert();
                                }
                            }
                        }
                        break;
                }
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            DataSourceBinded();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvViolationCaseMain_DataBound(object sender, EventArgs e)
        {
            switch (fvViolationCaseMain.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    plSearch.Visible = true;
                    /* 以下是用來限制各車站人員登入之後的權限用的...現在都先開放不限制
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vStationNo_FV = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vDepNo + "' and InSHReport = 'V'", "DepNo");
                    Button bbInsert_List = (Button)fvViolationCaseMain.FindControl("bbNew_List");
                    Button bbInsert_Empty = (Button)fvViolationCaseMain.FindControl("bbInsert_Empty");
                    Button bbEdit_List = (Button)fvViolationCaseMain.FindControl("bbEdit");
                    Button bbDelete_List = (Button)fvViolationCaseMain.FindControl("bbDelete");

                    if (bbInsert_List != null)
                    {
                        bbInsert_List.Visible = (vStationNo_FV == "");
                    }
                    if (bbInsert_Empty != null)
                    {
                        bbInsert_Empty.Visible = (vStationNo_FV == "");
                    }
                    if (bbEdit_List != null)
                    {
                        bbEdit_List.Visible = (vStationNo_FV == "");
                    }
                    if (bbDelete_List != null)
                    {
                        bbDelete_List.Visible = (vStationNo_FV == "");
                    }
                    //*/
                    break;

                case FormViewMode.Edit:
                    plSearch.Visible = false;
                    //違規日期
                    TextBox eTempViolationDate_Edit = (TextBox)fvViolationCaseMain.FindControl("eViolationDate_Edit");
                    string vViolationDateURL_Edit = "InputDate_ChineseYears.aspx?TextboxID=" + eTempViolationDate_Edit.ClientID;
                    string vViolationDateScript_Edit = "window.open('" + vViolationDateURL_Edit + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eTempViolationDate_Edit.Attributes["onClick"] = vViolationDateScript_Edit;
                    //繳納期限
                    TextBox eTempPaymentDeadline_Edit = (TextBox)fvViolationCaseMain.FindControl("ePaymentDeadline_Edit");
                    string vPaymentDeadlineURL_Edit = "InputDate_ChineseYears.aspx?TextboxID=" + eTempPaymentDeadline_Edit.ClientID;
                    string vPaymentDeadlineScript_Edit = "window.open('" + vPaymentDeadlineURL_Edit + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eTempPaymentDeadline_Edit.Attributes["onClick"] = vPaymentDeadlineScript_Edit;
                    //繳納日期
                    TextBox eTempPaidDate_Edit = (TextBox)fvViolationCaseMain.FindControl("ePaidDate_Edit");
                    string vPaidDateURL_Edit = "InputDate_ChineseYears.aspx?TextboxID=" + eTempPaidDate_Edit.ClientID;
                    string vPaidDateScript_Edit = "window.open('" + vPaidDateURL_Edit + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eTempPaidDate_Edit.Attributes["onClick"] = vPaidDateScript_Edit;
                    break;

                case FormViewMode.Insert:
                    plSearch.Visible = false;
                    //建檔日
                    TextBox eTempBuildDate_Insert = (TextBox)fvViolationCaseMain.FindControl("eBuildDate_Insert");
                    eTempBuildDate_Insert.Text = (DateTime.Today.Year - 1911).ToString("D3") + "/" + DateTime.Today.ToString("MM/dd");
                    string vBuildDateURL_Insert = "InputDate_ChineseYears.aspx?TextboxID=" + eTempBuildDate_Insert.ClientID;
                    string vBuildDateScript_Insert = "window.open('" + vBuildDateURL_Insert + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eTempBuildDate_Insert.Attributes["onClick"] = vBuildDateScript_Insert;
                    //建檔人
                    Label lbTempBuildMan_Insert = (Label)fvViolationCaseMain.FindControl("eBuildMan_Insert");
                    lbTempBuildMan_Insert.Text = vLoginID;
                    Label lbTempBuildManName_Insert = (Label)fvViolationCaseMain.FindControl("eBuildManName_Insert");
                    lbTempBuildManName_Insert.Text = Session["LoginName"].ToString().Trim();
                    //違規日期
                    TextBox eTempViolationDate_Insert = (TextBox)fvViolationCaseMain.FindControl("eViolationDate_Insert");
                    string vViolationDateURL_Insert = "InputDate_ChineseYears.aspx?TextboxID=" + eTempViolationDate_Insert.ClientID;
                    string vViolationDateScript_Insert = "window.open('" + vViolationDateURL_Insert + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eTempViolationDate_Insert.Attributes["onClick"] = vViolationDateScript_Insert;
                    //繳納期限
                    TextBox eTempPaymentDeadline_Insert = (TextBox)fvViolationCaseMain.FindControl("ePaymentDeadline_Insert");
                    string vPaymentDeadlineURL_Insert = "InputDate_ChineseYears.aspx?TextboxID=" + eTempPaymentDeadline_Insert.ClientID;
                    string vPaymentDeadlineScript_Insert = "window.open('" + vPaymentDeadlineURL_Insert + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eTempPaymentDeadline_Insert.Attributes["onClick"] = vPaymentDeadlineScript_Insert;
                    //繳納日期
                    TextBox eTempPaidDate_Insert = (TextBox)fvViolationCaseMain.FindControl("ePaidDate_Insert");
                    string vPaidDateURL_Insert = "InputDate_ChineseYears.aspx?TextboxID=" + eTempPaidDate_Insert.ClientID;
                    string vPaidDateScript_Insert = "window.open('" + vPaidDateURL_Insert + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eTempPaidDate_Insert.Attributes["onClick"] = vPaidDateScript_Insert;
                    break;
            }
        }

        protected void eCaseType_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eTempCaseType_Edit = (TextBox)fvViolationCaseMain.FindControl("eCaseType_Edit");

            //要變更寫入狀態的欄位
            TextBox eTempPenaltyDep_Edit = (TextBox)fvViolationCaseMain.FindControl("ePenaltyDep_Edit");
            TextBox eTempUndertaker_Edit = (TextBox)fvViolationCaseMain.FindControl("eUndertaker_Edit");
            TextBox eTempLinesNo_Edit = (TextBox)fvViolationCaseMain.FindControl("eLinesNo_Edit");
            TextBox eTempTicketTitle_Edit = (TextBox)fvViolationCaseMain.FindControl("eTicketTitle_Edit");
            TextBox eTempPenaltyReason_Edit = (TextBox)fvViolationCaseMain.FindControl("ePenaltyReason_Edit");
            //TextBox eTempViolationLocation_Edit = (TextBox)fvViolationCaseMain.FindControl("eViolationLocation_Edit");
            TextBox eTempPaymentDeadline_Edit = (TextBox)fvViolationCaseMain.FindControl("ePaymentDeadline_Edit");
            TextBox eTempPaidDate_Edit = (TextBox)fvViolationCaseMain.FindControl("ePaidDate_Edit");
            //TextBox eTempViolationRule_Edit = (TextBox)fvViolationCaseMain.FindControl("eViolationRule_Edit");
            //==================================================================================================
            eTempPenaltyDep_Edit.Enabled = (eTempCaseType_Edit.Text.Trim() == "V02");
            eTempUndertaker_Edit.Enabled = (eTempCaseType_Edit.Text.Trim() == "V02");
            eTempLinesNo_Edit.Enabled = (eTempCaseType_Edit.Text.Trim() == "V02");
            //eTempTicketTitle_Edit.Enabled = (eTempCaseType_Edit.Text.Trim() == "V02");
            eTempPenaltyReason_Edit.Enabled = (eTempCaseType_Edit.Text.Trim() == "V02");

            //eTempViolationLocation_Edit.Enabled = (eTempCaseType_Edit.Text.Trim() == "V01");
            eTempPaymentDeadline_Edit.Enabled = (eTempCaseType_Edit.Text.Trim() == "V01");
            eTempPaidDate_Edit.Enabled = (eTempCaseType_Edit.Text.Trim() == "V01");
            //eTempViolationRule_Edit.Enabled = (eTempCaseType_Edit.Text.Trim() == "V01");
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eTempDepNo_Edit = (TextBox)fvViolationCaseMain.FindControl("eDepNo_Edit");
            Label eTempDepName_Edit = (Label)fvViolationCaseMain.FindControl("eDepName_Edit");
            string vTempDepNo = eTempDepNo_Edit.Text.Trim();
            string vTempDepName = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vTempDepNo + "' ";
            vTempDepName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vTempDepName == "")
            {
                vTempDepName = vTempDepNo;
                vSQLStr = "select DepNo from Department where [Name] = '" + vTempDepName + "' ";
                vTempDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
                eTempDepNo_Edit.Text = vTempDepNo;
                eTempDepName_Edit.Text = vTempDepName;
            }
            else
            {
                eTempDepName_Edit.Text = vTempDepName;
            }
        }

        protected void eDriver_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eTempDriver_Edit = (TextBox)fvViolationCaseMain.FindControl("eDriver_Edit");
            Label eTempDriverName_Edit = (Label)fvViolationCaseMain.FindControl("eDriverName_Edit");
            string vTempDriver = eTempDriver_Edit.Text.Trim();
            string vTempDriverName = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Employee where isnull(LeaveDay,'') = '' and EmpNo = '" + vTempDriver + "' and Type = '20'";
            vTempDriverName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vTempDriverName == "")
            {
                vSQLStr = "select EmpNo from Employee where isnull(LeaveDay, '') = '' and [Name] = '" + vTempDriver + "' and Type = '20'";
                vTempDriverName = vTempDriver;
                vTempDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                eTempDriver_Edit.Text = vTempDriver;
                eTempDriverName_Edit.Text = vTempDriverName;
            }
            else
            {
                eTempDriverName_Edit.Text = vTempDriverName;
            }
        }

        protected void ddlCaseType_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTemp = (DropDownList)fvViolationCaseMain.FindControl("ddlCaseType_Insert");
            TextBox eTempCaseType = (TextBox)fvViolationCaseMain.FindControl("eCaseType_Insert");
            eTempCaseType.Text = ddlTemp.SelectedValue.Trim();

            //要變更寫入狀態的欄位
            TextBox eTempPenaltyDep_Insert = (TextBox)fvViolationCaseMain.FindControl("ePenaltyDep_Insert");
            TextBox eTempUndertaker_Insert = (TextBox)fvViolationCaseMain.FindControl("eUndertaker_Insert");
            TextBox eTempLinesNo_Insert = (TextBox)fvViolationCaseMain.FindControl("eLinesNo_Insert");
            TextBox eTempTicketTitle_Insert = (TextBox)fvViolationCaseMain.FindControl("eTicketTitle_Insert");
            TextBox eTempPenaltyReason_Insert = (TextBox)fvViolationCaseMain.FindControl("ePenaltyReason_Insert");
            //TextBox eTempViolationLocation_Insert = (TextBox)fvViolationCaseMain.FindControl("eViolationLocation_Insert");
            TextBox eTempPaymentDeadline_Insert = (TextBox)fvViolationCaseMain.FindControl("ePaymentDeadline_Insert");
            TextBox eTempPaidDate_Insert = (TextBox)fvViolationCaseMain.FindControl("ePaidDate_Insert");
            //TextBox eTempViolationRule_Insert = (TextBox)fvViolationCaseMain.FindControl("eViolationRule_Insert");

            eTempPenaltyDep_Insert.Enabled = (ddlTemp.SelectedValue.Trim() == "V02");
            eTempUndertaker_Insert.Enabled = (ddlTemp.SelectedValue.Trim() == "V02");
            eTempLinesNo_Insert.Enabled = (ddlTemp.SelectedValue.Trim() == "V02");
            //eTempTicketTitle_Insert.Enabled = (ddlTemp.SelectedValue.Trim() == "V02");
            eTempPenaltyReason_Insert.Enabled = (ddlTemp.SelectedValue.Trim() == "V02");

            //eTempViolationLocation_Insert.Enabled = (ddlTemp.SelectedValue.Trim() == "V01");
            eTempPaymentDeadline_Insert.Enabled = (ddlTemp.SelectedValue.Trim() == "V01");
            eTempPaidDate_Insert.Enabled = (ddlTemp.SelectedValue.Trim() == "V01");
            //eTempViolationRule_Insert.Enabled = (ddlTemp.SelectedValue.Trim() == "V01");
        }

        protected void eDepNo_Insert_TextChanged(object sender, EventArgs e)
        {
            TextBox eTempDepNo_Insert = (TextBox)fvViolationCaseMain.FindControl("eDepNo_Insert");
            Label eTempDepName_Insert = (Label)fvViolationCaseMain.FindControl("eDepName_Insert");
            string vTempDepNo = eTempDepNo_Insert.Text.Trim();
            string vTempDepName = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vTempDepNo + "' ";
            vTempDepName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vTempDepName == "")
            {
                vTempDepName = vTempDepNo;
                vSQLStr = "select DepNo from Department where [Name] = '" + vTempDepName + "' ";
                vTempDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
                eTempDepNo_Insert.Text = vTempDepNo;
                eTempDepName_Insert.Text = vTempDepName;
            }
            else
            {
                eTempDepName_Insert.Text = vTempDepName;
            }
        }

        protected void eDriver_Insert_TextChanged(object sender, EventArgs e)
        {
            TextBox eTempDriver_Insert = (TextBox)fvViolationCaseMain.FindControl("eDriver_Insert");
            Label eTempDriverName_Insert = (Label)fvViolationCaseMain.FindControl("eDriverName_Insert");
            string vTempDriver = eTempDriver_Insert.Text.Trim();
            string vTempDriverName = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Employee where isnull(LeaveDay,'') = '' and EmpNo = '" + vTempDriver + "' and Type = '20'";
            vTempDriverName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vTempDriverName == "")
            {
                vSQLStr = "select EmpNo from Employee where isnull(LeaveDay, '') = '' and [Name] = '" + vTempDriver + "' and Type = '20'";
                vTempDriverName = vTempDriver;
                vTempDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                eTempDriver_Insert.Text = vTempDriver;
                eTempDriverName_Insert.Text = vTempDriverName;
            }
            else
            {
                eTempDriverName_Insert.Text = vTempDriverName;
            }
        }

        protected void bbDelete_Click(object sender, EventArgs e)
        {
            string vDelConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            Label lbCaseNo_Del = (Label)fvViolationCaseMain.FindControl("eCaseNo_List");
            string vCaseNo_Del = lbCaseNo_Del.Text.Trim();
            string vDelCommandText = "INSERT INTO [dbo].[ViolationCaseHistory]" + Environment.NewLine +
                                     "            ([CaseNo],[CaseType],[BuildDate],[BuildMan],[BuildManName],[LinesNo],[DepNo],[DepName],[Car_ID], " + Environment.NewLine +
                                     "             [Driver],[DriverName],[ViolationDate],[TicketNo],[TicketTitle],[PenaltyDep],[Undertaker],[ViolationLocation], " + Environment.NewLine +
                                     "             [ViolationRule],[ViolationNote],[FineAmount],[PenaltyReason],[ViolationPoint],[PaymentDeadline],[PaidDate], " + Environment.NewLine +
                                     "             [Remark],[ModifyType],[ModifyDate],[ModifyMan])" + Environment.NewLine +
                                     "select [CaseNo],[CaseType],[BuildDate],[BuildMan],[BuildManName],[LinesNo],[DepNo],[DepName],[Car_ID], " + Environment.NewLine +
                                     "       [Driver],[DriverName],[ViolationDate],[TicketNo],[TicketTitle],[PenaltyDep],[Undertaker],[ViolationLocation], " + Environment.NewLine +
                                     "       [ViolationRule],[ViolationNote],[FineAmount],[PenaltyReason],[ViolationPoint],[PaymentDeadline],[PaidDate], " + Environment.NewLine +
                                     "       [Remark],'DEL',GetDate(),'" + vLoginID + "' " + Environment.NewLine +
                                     "  from [dbo].[ViolationCase] " + Environment.NewLine +
                                     " where CaseNo = '" + vCaseNo_Del + "' ";
            try
            {
                PF.ExecSQL(vDelConnStr, vDelCommandText);
                sdsViolationCase_Main.DeleteParameters.Clear();
                sdsViolationCase_Main.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Del));
                int vIsDeleted = sdsViolationCase_Main.Delete();
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        protected void sdsViolationCase_Main_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            TextBox tbTemp = (TextBox)fvViolationCaseMain.FindControl("eCaseType_Insert");
            //建檔日
            TextBox tbTemp_BuildDate = (TextBox)fvViolationCaseMain.FindControl("eBuildDate_Insert");
            //違規日期
            TextBox eTemp_ViolationDate = (TextBox)fvViolationCaseMain.FindControl("eViolationDate_Insert");
            //繳納期限
            TextBox eTemp_PaymentDeadline = (TextBox)fvViolationCaseMain.FindControl("ePaymentDeadline_Insert");
            //繳納日期
            TextBox eTemp_PaidDate = (TextBox)fvViolationCaseMain.FindControl("ePaidDate_Insert");


            DateTime vDateResult;
            string vCaseNo_TypeA = (DateTime.TryParse(tbTemp_BuildDate.Text.Trim(), out vDateResult)) ?
                                    vDateResult.Year.ToString("D4") + vDateResult.ToString("MMdd") + tbTemp.Text.Trim() :
                                    DateTime.Today.Year.ToString("D4") + DateTime.Today.ToString("MMdd") + tbTemp.Text.Trim();
            e.Command.Parameters["@CaseNo"].Value = GetCaseNo(vCaseNo_TypeA);
            e.Command.Parameters["@BuildDate"].Value = (tbTemp_BuildDate.Text.Trim() != "") ? DateTime.Parse(tbTemp_BuildDate.Text.Trim()) : DateTime.Today;
            if ((eTemp_ViolationDate != null) && (eTemp_ViolationDate.Text.Trim() != ""))
            {
                e.Command.Parameters["@ViolationDate"].Value = DateTime.Parse(eTemp_ViolationDate.Text.Trim());
            }
            if ((eTemp_PaymentDeadline != null) && (eTemp_PaymentDeadline.Text.Trim() != ""))
            {
                e.Command.Parameters["@PaymentDeadline"].Value = DateTime.Parse(eTemp_PaymentDeadline.Text.Trim());
            }
            if ((eTemp_PaidDate != null) && (eTemp_PaidDate.Text.Trim() != ""))
            {
                e.Command.Parameters["@PaidDate"].Value = DateTime.Parse(eTemp_PaidDate.Text.Trim());
            }
        }

        protected void sdsViolationCase_Main_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            //違規日期
            TextBox eTemp_ViolationDate = (TextBox)fvViolationCaseMain.FindControl("eViolationDate_Edit");
            //繳納期限
            TextBox eTemp_PaymentDeadline = (TextBox)fvViolationCaseMain.FindControl("ePaymentDeadline_Edit");
            //繳納日期
            TextBox eTemp_PaidDate = (TextBox)fvViolationCaseMain.FindControl("ePaidDate_Edit");

            if ((eTemp_ViolationDate != null) && (eTemp_ViolationDate.Text.Trim() != ""))
            {
                e.Command.Parameters["@ViolationDate"].Value = DateTime.Parse(eTemp_ViolationDate.Text.Trim());
            }
            if ((eTemp_PaymentDeadline != null) && (eTemp_PaymentDeadline.Text.Trim() != ""))
            {
                e.Command.Parameters["@PaymentDeadline"].Value = DateTime.Parse(eTemp_PaymentDeadline.Text.Trim());
            }
            if ((eTemp_PaidDate != null) && (eTemp_PaidDate.Text.Trim() != ""))
            {
                e.Command.Parameters["@PaidDate"].Value = DateTime.Parse(eTemp_PaidDate.Text.Trim());
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
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
            string vSelStr = GetSelectStr();
            string vFileName = "違規案件處理清冊";
            DateTime vBuDate;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
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
                        vHeaderText = (drExcel.GetName(i) == "CaseNo") ? "違規序號" :
                                      (drExcel.GetName(i) == "CaseType") ? "" :
                                      (drExcel.GetName(i) == "CaseType_C") ? "違規種類" :
                                      (drExcel.GetName(i) == "ViolationDate") ? "違規日期" :
                                      (drExcel.GetName(i) == "BuildDate") ? "建檔日期" :
                                      (drExcel.GetName(i) == "BuildManName") ? "建檔人" :
                                      (drExcel.GetName(i) == "LinesNo") ? "路線" :
                                      (drExcel.GetName(i) == "DepName") ? "站別" :
                                      (drExcel.GetName(i) == "Car_ID") ? "車號" :
                                      (drExcel.GetName(i) == "DriverName") ? "駕駛員" :
                                      (drExcel.GetName(i) == "TicketTitle") ? "發文主旨" :
                                      (drExcel.GetName(i) == "PenaltyDep") ? "裁罰單位" :
                                      (drExcel.GetName(i) == "FineAmount") ? "裁罰金額" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i) == "ViolationDate") ||
                                 (drExcel.GetName(i) == "BuildDate")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if ((drExcel.GetName(i) == "FineAmount") && (drExcel[i].ToString() != ""))
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
                            string vCaseTypeStr = ddlCaseType_Search.SelectedItem.Text.Trim();
                            string vDepNoStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? "從 " + eDepNo_Start_Search.Text.Trim() + " 起至 " + eDepNo_End_Search.Text.Trim() + " 止" :
                                               ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? eDepNo_Start_Search.Text.Trim() :
                                               ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? eDepNo_End_Search.Text.Trim() : "全部";
                            string vCarIDStr = ((eCarID_Start_Search.Text.Trim() != "") && (eCarID_End_Search.Text.Trim() != "")) ? "從 " + eCarID_Start_Search.Text.Trim() + " 起至 " + eCarID_End_Search.Text.Trim() + " 止" :
                                               ((eCarID_Start_Search.Text.Trim() != "") && (eCarID_End_Search.Text.Trim() == "")) ? eCarID_Start_Search.Text.Trim() :
                                               ((eCarID_Start_Search.Text.Trim() == "") && (eCarID_End_Search.Text.Trim() != "")) ? eCarID_End_Search.Text.Trim() : "全部";
                            string vViolationDateStr = ((eViolationDate_Start_Search.Text.Trim() != "") && (eViolationDate_End_Search.Text.Trim() != "")) ? "從 " + eViolationDate_Start_Search.Text.Trim() + " 起至 " + eViolationDate_End_Search.Text.Trim() + " 止" :
                                                       ((eViolationDate_Start_Search.Text.Trim() != "") && (eViolationDate_End_Search.Text.Trim() == "")) ? eViolationDate_Start_Search.Text.Trim() :
                                                       ((eViolationDate_Start_Search.Text.Trim() == "") && (eViolationDate_End_Search.Text.Trim() != "")) ? eViolationDate_End_Search.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_違規案件處理" + Environment.NewLine +
                                                 "ViolationCase.aspx" + Environment.NewLine +
                                                 "違規類別：" + vCaseTypeStr + Environment.NewLine +
                                                 "站別：" + vDepNoStr + Environment.NewLine +
                                                 "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                                 "違規日期：" + vViolationDateStr;
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
    }
}