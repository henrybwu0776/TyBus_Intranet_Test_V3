using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class DataImport : System.Web.UI.Page
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

        private string vExtName = "";

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

                UnobtrusiveValidationMode = System.Web.UI.UnobtrusiveValidationMode.None;

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        eCalYear.Text = (DateTime.Today.Year - 1911).ToString();
                        eCalMonth.Text = DateTime.Today.AddMonths(1).Month.ToString();
                        eShowError.Visible = false;
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

        protected void bbImport_Click(object sender, EventArgs e)
        {
            eShowError.Visible = false;
            if ((fuExcel.FileName == "") || (ddlTargetTable.SelectedValue.Trim() == ""))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
                ddlTargetTable.Focus();
            }
            else
            {
                switch (ddlTargetTable.SelectedValue)
                {
                    case "T001": //轉入違規案件
                        Import_ViolationCase();
                        break;
                    case "T002": //轉入肇事案件
                        Import_AnecdoteCase();
                        break;
                    case "T003":
                        Import_FixSevNo(); //轉入維修項目代號
                        break;
                    case "T004":
                        Import_CustomService(); //轉入客服案件
                        break;
                    case "T005":
                        Import_CarServicePre(); //轉入保養預排
                        break;
                    case "T006":
                        Import_ContractManager(); //轉入合約書資料
                        break;
                    case "T007":
                        Import_LinesData(); //轉入路線對照
                        break;
                    case "T008":
                        Import_UnusedBounds(); //更新不休假獎金資料
                        break;
                    case "T009":
                        Import_MOUPayData(); //匯入MOU津貼資料
                        break;
                    case "T010":
                        Modify_MOUPayData(); //修改MOU津貼資料
                        break;
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 轉入違規案件
        /// </summary>
        private void Import_ViolationCase()
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            string vSheetName = "";
            int vSheetCount = 0;

            string vCaseNo = "";
            string vCaseNo_TypeA = "";
            string vCaseType = "";
            string vBuildDate = "";
            string vBuildMan = "";
            string vBuildManName = "";
            string vLinesNo = "";
            string vDepNo_Temp = "";
            string vDepName_Temp = "";
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
                        vSheetCount = wbExcel_H.NumberOfSheets;
                        for (int vSC = 0; vSC < vSheetCount; vSC++)
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
                                    vBuildDate = (vRowTemp_H.GetCell(1).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(1).DateCellValue.ToString() : DateTime.Today.ToString("yyyy/MM/dd");
                                    vCaseNo_TypeA = DateTime.Parse(vBuildDate).ToString("yyyyMMdd") + vCaseType;
                                    vCaseNo = PF.GetDataIndex(vConnStr, "ViolationCase", "CaseNo", "B", false, DateTime.Parse(vBuildDate), 8, vCaseType, 4);
                                    vBuildManName = (vRowTemp_H.GetCell(2).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(2).StringCellValue.Trim() : Session["LoginName"].ToString().Trim();
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vBuildManName + "' and LeaveDay is null";
                                    vBuildMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    vBuildMan = (vBuildMan == "") ? vLoginID : vBuildMan;
                                    vLinesNo = (vRowTemp_H.GetCell(3).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(3).ToString().Trim() : "";
                                    vDepNo_Temp = (vRowTemp_H.GetCell(4).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(4).ToString().Trim() : "";
                                    vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "'";
                                    vDepName_Temp = (vDepNo_Temp != "") ? PF.GetValue(vConnStr, vSQLStr, "Name") : "";
                                    vCar_ID = (vRowTemp_H.GetCell(5).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(5).StringCellValue.Trim() : "";
                                    vDriverName = (vRowTemp_H.GetCell(6).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(6).StringCellValue.Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null and Type = '20'";
                                    vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    vViolationDate = (vRowTemp_H.GetCell(7).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(7).DateCellValue.ToString() : "";
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
                                    vPaymentDeadline = (vRowTemp_H.GetCell(18).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(18).DateCellValue.ToString() : "";
                                    vPaidDate = (vRowTemp_H.GetCell(19).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(19).DateCellValue.ToString() : "";
                                    vRemark = (vRowTemp_H.GetCell(20).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(20).ToString().Trim() : "";

                                    using (SqlDataSource sdsTemp = new SqlDataSource())
                                    {
                                        sdsTemp.ConnectionString = vConnStr;
                                        sdsTemp.InsertCommand = "INSERT INTO ViolationCase " + Environment.NewLine +
                                                                "       (CaseNo, CaseType, BuildDate, BuildMan, BuildManName, LinesNo, DepNo, DepName, Car_ID, " + Environment.NewLine +
                                                                "        Driver, DriverName, ViolationDate, TicketNo, TicketTitle, PenaltyDep, Undertaker, " + Environment.NewLine +
                                                                "        ViolationLocation, ViolationRule, ViolationNote, FineAmount, PenaltyReason, ViolationPoint, " + Environment.NewLine +
                                                                "        PaymentDeadline, PaidDate, Remark) " + Environment.NewLine +
                                                                "VALUES (@CaseNo, @CaseType, @BuildDate, @BuildMan, @BuildManName, @LinesNo, @DepNo, @DepName, " + Environment.NewLine +
                                                                "        @Car_ID, @Driver, @DriverName, @ViolationDate, @TicketNo, @TicketTitle, @PenaltyDep, " + Environment.NewLine +
                                                                "        @Undertaker, @ViolationLocation, @ViolationRule, @ViolationNote, @FineAmount, @PenaltyReason, " + Environment.NewLine +
                                                                "        @ViolationPoint, @PaymentDeadline, @PaidDate, @Remark)";
                                        sdsTemp.InsertParameters.Clear();
                                        sdsTemp.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                        sdsTemp.InsertParameters.Add(new Parameter("CaseType", DbType.String, vCaseType));
                                        sdsTemp.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, vBuildDate));
                                        sdsTemp.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vBuildMan));
                                        sdsTemp.InsertParameters.Add(new Parameter("BuildManName", DbType.String, vBuildManName));
                                        sdsTemp.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                        sdsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_Temp));
                                        sdsTemp.InsertParameters.Add(new Parameter("DepName", DbType.String, vDepName_Temp));
                                        sdsTemp.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                        sdsTemp.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                        sdsTemp.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationDate", DbType.Date, vViolationDate));
                                        sdsTemp.InsertParameters.Add(new Parameter("TicketNo", DbType.String, vTicketNo));
                                        sdsTemp.InsertParameters.Add(new Parameter("TicketTitle", DbType.String, vTicketTitle));
                                        sdsTemp.InsertParameters.Add(new Parameter("PenaltyDep", DbType.String, vPenaltyDep));
                                        sdsTemp.InsertParameters.Add(new Parameter("Undertaker", DbType.String, vUndertaker));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationLocation", DbType.String, vViolationLocation));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationRule", DbType.String, vViolationRule));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationNote", DbType.String, vViolationNote));
                                        sdsTemp.InsertParameters.Add(new Parameter("FineAmount", DbType.Int32, vFineAmount));
                                        sdsTemp.InsertParameters.Add(new Parameter("PenaltyReason", DbType.String, vPenaltyReason));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationPoint", DbType.Int32, vViolationPoint));
                                        sdsTemp.InsertParameters.Add(new Parameter("PaymentDeadline", DbType.Date, vPaymentDeadline));
                                        sdsTemp.InsertParameters.Add(new Parameter("PaidDate", DbType.Date, vPaidDate));
                                        sdsTemp.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                        sdsTemp.Insert();
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx":
                        //Excel 2010--
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        vSheetCount = wbExcel_X.NumberOfSheets;
                        for (int vSC = 0; vSC < vSheetCount; vSC++)
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
                                    vBuildDate = (vRowTemp_H.GetCell(1).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(1).DateCellValue.ToString() : DateTime.Today.ToString("yyyy/MM/dd");
                                    vCaseNo_TypeA = DateTime.Parse(vBuildDate).ToString("yyyyMMdd") + vCaseType;
                                    vCaseNo = PF.GetDataIndex(vConnStr, "ViolationCase", "CaseNo", "B", false, DateTime.Parse(vBuildDate), 8, vCaseType, 4);
                                    vBuildManName = (vRowTemp_H.GetCell(2).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(2).StringCellValue.Trim() : Session["LoginName"].ToString().Trim();
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vBuildManName + "' and LeaveDay is null";
                                    vBuildMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    vBuildMan = (vBuildMan == "") ? vLoginID : vBuildMan;
                                    vLinesNo = (vRowTemp_H.GetCell(3).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(3).ToString().Trim() : "";
                                    vDepNo_Temp = (vRowTemp_H.GetCell(4).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(4).ToString().Trim() : "";
                                    vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "'";
                                    vDepName_Temp = (vDepNo_Temp != "") ? PF.GetValue(vConnStr, vSQLStr, "Name") : "";
                                    vCar_ID = (vRowTemp_H.GetCell(5).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(5).StringCellValue.Trim() : "";
                                    vDriverName = (vRowTemp_H.GetCell(6).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(6).StringCellValue.Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null and Type = '20'";
                                    vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    vViolationDate = (vRowTemp_H.GetCell(7).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(7).DateCellValue.ToString() : "";
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
                                    vPaymentDeadline = (vRowTemp_H.GetCell(18).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(18).DateCellValue.ToString() : "";
                                    vPaidDate = (vRowTemp_H.GetCell(19).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(19).DateCellValue.ToString() : "";
                                    vRemark = (vRowTemp_H.GetCell(20).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(20).ToString().Trim() : "";

                                    using (SqlDataSource sdsTemp = new SqlDataSource())
                                    {
                                        sdsTemp.ConnectionString = vConnStr;
                                        sdsTemp.InsertCommand = "INSERT INTO ViolationCase " + Environment.NewLine +
                                                                "       (CaseNo, CaseType, BuildDate, BuildMan, BuildManName, LinesNo, DepNo, DepName, Car_ID, " + Environment.NewLine +
                                                                "        Driver, DriverName, ViolationDate, TicketNo, TicketTitle, PenaltyDep, Undertaker, " + Environment.NewLine +
                                                                "        ViolationLocation, ViolationRule, ViolationNote, FineAmount, PenaltyReason, ViolationPoint, " + Environment.NewLine +
                                                                "        PaymentDeadline, PaidDate, Remark) " + Environment.NewLine +
                                                                "VALUES (@CaseNo, @CaseType, @BuildDate, @BuildMan, @BuildManName, @LinesNo, @DepNo, @DepName, " + Environment.NewLine +
                                                                "        @Car_ID, @Driver, @DriverName, @ViolationDate, @TicketNo, @TicketTitle, @PenaltyDep, " + Environment.NewLine +
                                                                "        @Undertaker, @ViolationLocation, @ViolationRule, @ViolationNote, @FineAmount, @PenaltyReason, " + Environment.NewLine +
                                                                "        @ViolationPoint, @PaymentDeadline, @PaidDate, @Remark)";
                                        sdsTemp.InsertParameters.Clear();
                                        sdsTemp.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                        sdsTemp.InsertParameters.Add(new Parameter("CaseType", DbType.String, vCaseType));
                                        sdsTemp.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, vBuildDate));
                                        sdsTemp.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vBuildMan));
                                        sdsTemp.InsertParameters.Add(new Parameter("BuildManName", DbType.String, vBuildManName));
                                        sdsTemp.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                        sdsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_Temp));
                                        sdsTemp.InsertParameters.Add(new Parameter("DepName", DbType.String, vDepName_Temp));
                                        sdsTemp.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                        sdsTemp.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                        sdsTemp.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationDate", DbType.Date, vViolationDate));
                                        sdsTemp.InsertParameters.Add(new Parameter("TicketNo", DbType.String, vTicketNo));
                                        sdsTemp.InsertParameters.Add(new Parameter("TicketTitle", DbType.String, vTicketTitle));
                                        sdsTemp.InsertParameters.Add(new Parameter("PenaltyDep", DbType.String, vPenaltyDep));
                                        sdsTemp.InsertParameters.Add(new Parameter("Undertaker", DbType.String, vUndertaker));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationLocation", DbType.String, vViolationLocation));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationRule", DbType.String, vViolationRule));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationNote", DbType.String, vViolationNote));
                                        sdsTemp.InsertParameters.Add(new Parameter("FineAmount", DbType.Int32, vFineAmount));
                                        sdsTemp.InsertParameters.Add(new Parameter("PenaltyReason", DbType.String, vPenaltyReason));
                                        sdsTemp.InsertParameters.Add(new Parameter("ViolationPoint", DbType.Int32, vViolationPoint));
                                        sdsTemp.InsertParameters.Add(new Parameter("PaymentDeadline", DbType.Date, vPaymentDeadline));
                                        sdsTemp.InsertParameters.Add(new Parameter("PaidDate", DbType.Date, vPaidDate));
                                        sdsTemp.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                        sdsTemp.Insert();
                                    }
                                }
                            }
                        }
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 轉入肇事案件
        /// </summary>
        private void Import_AnecdoteCase()
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            string vUploadFileName = "";
            //string vExtName = "";
            string vTempIndex = "";
            string vTempStr = "";
            int vTempINT = 0;
            DateTime vTempDate;

            string vCaseNo = "";
            string vDepNo_Temp = "";
            string vDepName_Temp = "";
            string vBuildDate = "";
            string vBuildMan = "";
            string vCar_ID = "";
            string vDriver = "";
            string vDriverName = "";
            string vInsuMan = "";
            string vAncedotalResRatio = "0";
            string vIsNoDeduction = "0";
            string vDeductionDate = "";
            string vRemark = "";
            string vItems = "";
            string vCaseNoItems = "";
            string vRelationship = "";
            string vRelCar_ID = "";
            string vEstimatedAmount = "0";
            string vThirdInsurance = "0";
            string vCompInsurance = "0";
            string vDriverSharing = "0";
            string vCompanySharing = "0";
            string vCarDamageAMT = "0";
            string vPersonDamageAMT = "0";
            string vRelationComp = "0";
            string vReconciliationDate = "";
            string vRemarkB = "";
            string vPassengerInsu = "0";
            string vSQLStr = "";

            if (fuExcel.FileName != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                vExtName = Path.GetExtension(fuExcel.FileName);
                vUploadFileName = vUploadPath + fuExcel.FileName;
                fuExcel.SaveAs(vUploadFileName);
                if (vExtName == ".xls")
                {
                    HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                    HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                    for (int i = sheetExcel_H.FirstRowNum + 2; i <= sheetExcel_H.LastRowNum; i++)
                    {
                        HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                        if (vRowTemp_H.GetCell(1).CellType != CellType.Blank)
                        {
                            vDepNo_Temp = vRowTemp_H.GetCell(1).ToString().Trim();
                            vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
                            vBuildDate = (DateTime.ParseExact(vRowTemp_H.GetCell(3).ToString().Trim(), "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd");
                            vCaseNo = (DateTime.Parse(vBuildDate).Year - 1911).ToString("D3") + DateTime.Parse(vBuildDate).Month.ToString("D2") + "A";
                            vSQLStr = "select MAX(CaseNo) CaseNo from AnecdoteCase where CaseNo like '" + vCaseNo + "%' ";
                            vTempIndex = (PF.GetValue(vConnStr, vSQLStr, "CaseNo")).Replace(vCaseNo, "").Trim();
                            vTempIndex = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D4") : "0001";
                            vCaseNo = vCaseNo + vTempIndex;
                            vBuildMan = vLoginID;
                            vCar_ID = (vRowTemp_H.GetCell(4).CellType != CellType.Blank) ? vRowTemp_H.GetCell(4).ToString().Trim() : "";
                            vDriverName = vRowTemp_H.GetCell(5).ToString().Trim();
                            vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' ";
                            vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                            vInsuMan = (vRowTemp_H.GetCell(8).CellType != CellType.Blank) ? vRowTemp_H.GetCell(8).ToString().Trim() : "";
                            vAncedotalResRatio = (vRowTemp_H.GetCell(18).CellType != CellType.Blank) ? vRowTemp_H.GetCell(18).ToString().Trim().Replace(",", "") : "0";
                            vRemark = "";
                            vDeductionDate = "";
                            if (vRowTemp_H.GetCell(19).CellType != CellType.Blank)
                            {
                                vTempStr = vRowTemp_H.GetCell(19).ToString().Trim();
                                vIsNoDeduction = (vTempStr == "免扣") ? "true" : "false";
                                vTempStr = vTempStr.Replace("免扣", "");
                                //vTempStr = vTempStr.Replace("扣", "");
                                if (DateTime.TryParseExact(vTempStr.Replace("扣", ""), "yyy.MM.dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out vTempDate))
                                {
                                    vDeductionDate = vTempDate.AddYears(1911).ToString("yyyy/MM/dd");
                                    vRemark = "";
                                }
                                else
                                {
                                    vDeductionDate = "";
                                    vRemark = vTempStr;
                                }
                            }
                            else
                            {
                                vIsNoDeduction = "false";
                                vDeductionDate = "";
                            }
                            if (vRemark != "")
                            {
                                vRemark = (vRowTemp_H.GetCell(20).CellType != CellType.Blank) ? vRemark + Environment.NewLine + vRowTemp_H.GetCell(20).ToString().Trim() : vRemark;
                            }
                            else
                            {
                                vRemark = (vRowTemp_H.GetCell(20).CellType != CellType.Blank) ? vRowTemp_H.GetCell(20).ToString().Trim() : "";
                            }

                            using (SqlDataSource dsTemp = new SqlDataSource())
                            {
                                dsTemp.ConnectionString = (vConnStr == "") ? PF.GetConnectionStr(Request.ApplicationPath) : vConnStr;
                                dsTemp.InsertCommand = "INSERT INTO AnecdoteCase (CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, " + Environment.NewLine +
                                                       "                          Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, " + Environment.NewLine +
                                                       "                          DeductionDate, Remark) " + Environment.NewLine +
                                                       "VALUES (@CaseNo, @HasInsurance, @DepNo, @DepName, @BuildDate, @BuildMan, @Car_ID, " + Environment.NewLine +
                                                       "        @Driver, @DriverName, @InsuMan, @AnecdotalResRatio, @IsNoDeduction, @DeductionDate, @Remark)";
                                dsTemp.InsertParameters.Clear();
                                dsTemp.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                dsTemp.InsertParameters.Add(new Parameter("HasInsurance", DbType.Boolean, "false")); //先全部用 "未出險" 轉入
                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_Temp));
                                dsTemp.InsertParameters.Add(new Parameter("DepName", DbType.String, vDepName_Temp));
                                dsTemp.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, vBuildDate));
                                dsTemp.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vBuildMan));
                                dsTemp.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                dsTemp.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                dsTemp.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                dsTemp.InsertParameters.Add(new Parameter("InsuMan", DbType.String, vInsuMan));
                                dsTemp.InsertParameters.Add(new Parameter("AnecdotalResRatio", DbType.Double, vAncedotalResRatio));
                                dsTemp.InsertParameters.Add(new Parameter("IsNoDeduction", DbType.Boolean, vIsNoDeduction));
                                dsTemp.InsertParameters.Add(new Parameter("DeductionDate", DbType.Date, vDeductionDate));
                                dsTemp.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                dsTemp.Insert();
                            }
                        }

                        if (vRowTemp_H.GetCell(6).CellType != CellType.Blank)
                        {
                            vSQLStr = "select MAX(Items) MaxItem from AnecdoteCaseB where CaseNo = '" + vCaseNo + "' ";
                            vTempIndex = PF.GetValue(vConnStr, vSQLStr, "MaxItem");
                            vItems = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D4") : "0001";
                            vCaseNoItems = vCaseNo + vItems;
                            vRelationship = (vRowTemp_H.GetCell(6).CellType != CellType.Blank) ? vRowTemp_H.GetCell(6).ToString().Trim() : "";
                            vRelCar_ID = (vRowTemp_H.GetCell(7).CellType != CellType.Blank) ? vRowTemp_H.GetCell(7).ToString().Trim() : "";
                            vEstimatedAmount = (vRowTemp_H.GetCell(9).CellType != CellType.Blank) ? vRowTemp_H.GetCell(9).ToString().Trim().Replace(",", "") : "0";
                            if (Int32.TryParse(vEstimatedAmount, out vTempINT))
                            {
                                vEstimatedAmount = vTempINT.ToString();
                                vRemarkB = "";
                            }
                            else
                            {
                                vRemarkB = vEstimatedAmount;
                                vEstimatedAmount = "0";
                            }
                            vThirdInsurance = (vRowTemp_H.GetCell(10).CellType != CellType.Blank) ? vRowTemp_H.GetCell(10).ToString().Trim().Replace(",", "") : "0";
                            vCompInsurance = (vRowTemp_H.GetCell(11).CellType != CellType.Blank) ? vRowTemp_H.GetCell(11).ToString().Trim().Replace(",", "") : "0";
                            vDriverSharing = (vRowTemp_H.GetCell(12).CellType != CellType.Blank) ? vRowTemp_H.GetCell(12).ToString().Trim().Replace(",", "") : "0";
                            vCompanySharing = (vRowTemp_H.GetCell(13).CellType != CellType.Blank) ? vRowTemp_H.GetCell(13).ToString().Trim().Replace(",", "") : "0";
                            vCarDamageAMT = (vRowTemp_H.GetCell(14).CellType != CellType.Blank) ? vRowTemp_H.GetCell(14).ToString().Trim().Replace(",", "") : "0";
                            vPersonDamageAMT = (vRowTemp_H.GetCell(15).CellType != CellType.Blank) ? vRowTemp_H.GetCell(15).ToString().Trim().Replace(",", "") : "0";
                            vRelationComp = (vRowTemp_H.GetCell(16).CellType != CellType.Blank) ? vRowTemp_H.GetCell(16).ToString().Trim().Replace(",", "") : "0";
                            vTempStr = (vRowTemp_H.GetCell(17).CellType != CellType.Blank) ? vRowTemp_H.GetCell(17).ToString().Trim() : "";
                            vTempStr = vTempStr.Replace("和解", "");
                            vTempStr = vTempStr.Replace("自行", "");
                            vPassengerInsu = "0"; //乘客險先用 0 轉入
                            vReconciliationDate = (vTempStr != "") ? (DateTime.ParseExact(vTempStr, "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd") : vTempStr;

                            using (SqlDataSource dsTempB = new SqlDataSource())
                            {
                                dsTempB.ConnectionString = (vConnStr == "") ? PF.GetConnectionStr(Request.ApplicationPath) : vConnStr;
                                dsTempB.InsertCommand = "INSERT INTO AnecdoteCaseB (CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, " + Environment.NewLine +
                                                        "                           ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, " + Environment.NewLine +
                                                        "                           PersonDamageAMT, RelationComp, ReconciliationDate, Remark, PassengerInsu) " + Environment.NewLine +
                                                        "VALUES (@CaseNo, @Items, @CaseNoItems, @Relationship, @RelCar_ID, @EstimatedAmount, " + Environment.NewLine +
                                                        "        @ThirdInsurance, @CompInsurance, @DriverSharing, @CompanySharing, @CarDamageAMT, " + Environment.NewLine +
                                                        "        @PersonDamageAMT, @RelationComp, @ReconciliationDate, @Remark, @PassengerInsu)";
                                dsTempB.InsertParameters.Clear();
                                dsTempB.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                dsTempB.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                                dsTempB.InsertParameters.Add(new Parameter("CaseNoItems", DbType.String, vCaseNoItems));
                                dsTempB.InsertParameters.Add(new Parameter("Relationship", DbType.String, vRelationship));
                                dsTempB.InsertParameters.Add(new Parameter("RelCar_ID", DbType.String, vRelCar_ID));
                                dsTempB.InsertParameters.Add(new Parameter("EstimatedAmount", DbType.Int32, vEstimatedAmount));
                                dsTempB.InsertParameters.Add(new Parameter("ThirdInsurance", DbType.Int32, vThirdInsurance));
                                dsTempB.InsertParameters.Add(new Parameter("CompInsurance", DbType.Int32, vCompInsurance));
                                dsTempB.InsertParameters.Add(new Parameter("DriverSharing", DbType.Int32, vDriverSharing));
                                dsTempB.InsertParameters.Add(new Parameter("CompanySharing", DbType.Int32, vCompanySharing));
                                dsTempB.InsertParameters.Add(new Parameter("CarDamageAMT", DbType.Int32, vCarDamageAMT));
                                dsTempB.InsertParameters.Add(new Parameter("PersonDamageAMT", DbType.Int32, vPersonDamageAMT));
                                dsTempB.InsertParameters.Add(new Parameter("RelationComp", DbType.Int32, vRelationComp));
                                dsTempB.InsertParameters.Add(new Parameter("ReconciliationDate", DbType.Date, vReconciliationDate));
                                dsTempB.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemarkB));
                                dsTempB.InsertParameters.Add(new Parameter("PassengerInsu", DbType.Int32, vPassengerInsu));
                                dsTempB.Insert();
                            }
                        }
                    }
                }
                else if (vExtName == ".xlsx")
                {
                    XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                    XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                    for (int i = sheetExcel_X.FirstRowNum + 2; i <= sheetExcel_X.LastRowNum; i++)
                    {
                        XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                        if (vRowTemp_X.GetCell(1).CellType != CellType.Blank)
                        {
                            vDepNo_Temp = vRowTemp_X.GetCell(1).ToString().Trim();
                            vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
                            vBuildDate = (DateTime.ParseExact(vRowTemp_X.GetCell(3).ToString().Trim(), "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd");
                            vCaseNo = (DateTime.Parse(vBuildDate).Year - 1911).ToString("D3") + DateTime.Parse(vBuildDate).Month.ToString("D2") + "A";
                            vSQLStr = "select MAX(CaseNo) CaseNo from AnecdoteCase where CaseNo like '" + vCaseNo + "%' ";
                            vTempIndex = (PF.GetValue(vConnStr, vSQLStr, "CaseNo")).Replace(vCaseNo, "").Trim();
                            vTempIndex = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D4") : "0001";
                            vCaseNo = vCaseNo + vTempIndex;
                            vBuildMan = vLoginID;
                            vCar_ID = (vRowTemp_X.GetCell(4).CellType != CellType.Blank) ? vRowTemp_X.GetCell(4).ToString().Trim() : "";
                            vDriverName = vRowTemp_X.GetCell(5).ToString().Trim();
                            vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' ";
                            vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                            vInsuMan = (vRowTemp_X.GetCell(8).CellType != CellType.Blank) ? vRowTemp_X.GetCell(8).ToString().Trim() : "";
                            vAncedotalResRatio = (vRowTemp_X.GetCell(18).CellType != CellType.Blank) ? vRowTemp_X.GetCell(18).ToString().Trim().Replace(",", "") : "0";
                            vRemark = "";
                            vDeductionDate = "";
                            if (vRowTemp_X.GetCell(19).CellType != CellType.Blank)
                            {
                                vTempStr = vRowTemp_X.GetCell(19).ToString().Trim();
                                vIsNoDeduction = (vTempStr == "免扣") ? "true" : "false";
                                vTempStr = vTempStr.Replace("免扣", "");
                                //vTempStr = vTempStr.Replace("扣", "");
                                if (DateTime.TryParseExact(vTempStr.Replace("扣", ""), "yyy.MM.dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out vTempDate))
                                {
                                    vDeductionDate = vTempDate.AddYears(1911).ToString("yyyy/MM/dd");
                                    vRemark = "";
                                }
                                else
                                {
                                    vDeductionDate = "";
                                    vRemark = vTempStr;
                                }
                            }
                            else
                            {
                                vIsNoDeduction = "false";
                                vDeductionDate = "";
                            }
                            if (vRemark != "")
                            {
                                vRemark = (vRowTemp_X.GetCell(20).CellType != CellType.Blank) ? vRemark + Environment.NewLine + vRowTemp_X.GetCell(20).ToString().Trim() : vRemark;
                            }
                            else
                            {
                                vRemark = (vRowTemp_X.GetCell(20).CellType != CellType.Blank) ? vRowTemp_X.GetCell(20).ToString().Trim() : "";
                            }

                            using (SqlDataSource dsTemp = new SqlDataSource())
                            {
                                dsTemp.ConnectionString = (vConnStr == "") ? PF.GetConnectionStr(Request.ApplicationPath) : vConnStr;
                                dsTemp.InsertCommand = "INSERT INTO AnecdoteCase (CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, " + Environment.NewLine +
                                                       "                          Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, " + Environment.NewLine +
                                                       "                          DeductionDate, Remark) " + Environment.NewLine +
                                                       "VALUES (@CaseNo, @HasInsurance, @DepNo, @DepName, @BuildDate, @BuildMan, @Car_ID, " + Environment.NewLine +
                                                       "        @Driver, @DriverName, @InsuMan, @AnecdotalResRatio, @IsNoDeduction, @DeductionDate, @Remark)";
                                dsTemp.InsertParameters.Clear();
                                dsTemp.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                dsTemp.InsertParameters.Add(new Parameter("HasInsurance", DbType.Boolean, "false")); //先全部用 "未出險" 轉入
                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_Temp));
                                dsTemp.InsertParameters.Add(new Parameter("DepName", DbType.String, vDepName_Temp));
                                dsTemp.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, vBuildDate));
                                dsTemp.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vBuildMan));
                                dsTemp.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                dsTemp.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                dsTemp.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                dsTemp.InsertParameters.Add(new Parameter("InsuMan", DbType.String, vInsuMan));
                                dsTemp.InsertParameters.Add(new Parameter("AnecdotalResRatio", DbType.Double, vAncedotalResRatio));
                                dsTemp.InsertParameters.Add(new Parameter("IsNoDeduction", DbType.Boolean, vIsNoDeduction));
                                dsTemp.InsertParameters.Add(new Parameter("DeductionDate", DbType.Date, vDeductionDate));
                                dsTemp.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                dsTemp.Insert();
                            }
                        }

                        if (vRowTemp_X.GetCell(6).CellType != CellType.Blank)
                        {
                            vSQLStr = "select MAX(Items) MaxItem from AnecdoteCaseB where CaseNo = '" + vCaseNo + "' ";
                            vTempIndex = PF.GetValue(vConnStr, vSQLStr, "MaxItem");
                            vItems = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D4") : "0001";
                            vCaseNoItems = vCaseNo + vItems;
                            vRelationship = (vRowTemp_X.GetCell(6).CellType != CellType.Blank) ? vRowTemp_X.GetCell(6).ToString().Trim() : "";
                            vRelCar_ID = (vRowTemp_X.GetCell(7).CellType != CellType.Blank) ? vRowTemp_X.GetCell(7).ToString().Trim() : "";
                            vEstimatedAmount = (vRowTemp_X.GetCell(9).CellType != CellType.Blank) ? vRowTemp_X.GetCell(9).ToString().Trim().Replace(",", "") : "0";
                            if (Int32.TryParse(vEstimatedAmount, out vTempINT))
                            {
                                vEstimatedAmount = vTempINT.ToString();
                                vRemarkB = "";
                            }
                            else
                            {
                                vRemarkB = vEstimatedAmount;
                                vEstimatedAmount = "0";
                            }
                            vThirdInsurance = (vRowTemp_X.GetCell(10).CellType != CellType.Blank) ? vRowTemp_X.GetCell(10).ToString().Trim().Replace(",", "") : "0";
                            vCompInsurance = (vRowTemp_X.GetCell(11).CellType != CellType.Blank) ? vRowTemp_X.GetCell(11).ToString().Trim().Replace(",", "") : "0";
                            vDriverSharing = (vRowTemp_X.GetCell(12).CellType != CellType.Blank) ? vRowTemp_X.GetCell(12).ToString().Trim().Replace(",", "") : "0";
                            vCompanySharing = (vRowTemp_X.GetCell(13).CellType != CellType.Blank) ? vRowTemp_X.GetCell(13).ToString().Trim().Replace(",", "") : "0";
                            vCarDamageAMT = (vRowTemp_X.GetCell(14).CellType != CellType.Blank) ? vRowTemp_X.GetCell(14).ToString().Trim().Replace(",", "") : "0";
                            vPersonDamageAMT = (vRowTemp_X.GetCell(15).CellType != CellType.Blank) ? vRowTemp_X.GetCell(15).ToString().Trim().Replace(",", "") : "0";
                            vRelationComp = (vRowTemp_X.GetCell(16).CellType != CellType.Blank) ? vRowTemp_X.GetCell(16).ToString().Trim().Replace(",", "") : "0";
                            vTempStr = (vRowTemp_X.GetCell(17).CellType != CellType.Blank) ? vRowTemp_X.GetCell(17).ToString().Trim() : "";
                            vTempStr = vTempStr.Replace("和解", "");
                            vTempStr = vTempStr.Replace("自行", "");
                            vPassengerInsu = "0"; //乘客險先用 0 轉入
                            vReconciliationDate = (vTempStr != "") ? (DateTime.ParseExact(vTempStr, "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd") : vTempStr;

                            using (SqlDataSource dsTempB = new SqlDataSource())
                            {
                                dsTempB.ConnectionString = (vConnStr == "") ? PF.GetConnectionStr(Request.ApplicationPath) : vConnStr;
                                dsTempB.InsertCommand = "INSERT INTO AnecdoteCaseB (CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, " + Environment.NewLine +
                                                        "                           ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, " + Environment.NewLine +
                                                        "                           PersonDamageAMT, RelationComp, ReconciliationDate, Remark, PassengerInsu) " + Environment.NewLine +
                                                        "VALUES (@CaseNo, @Items, @CaseNoItems, @Relationship, @RelCar_ID, @EstimatedAmount, " + Environment.NewLine +
                                                        "        @ThirdInsurance, @CompInsurance, @DriverSharing, @CompanySharing, @CarDamageAMT, " + Environment.NewLine +
                                                        "        @PersonDamageAMT, @RelationComp, @ReconciliationDate, @Remark, @PassengerInsu)";
                                dsTempB.InsertParameters.Clear();
                                dsTempB.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                dsTempB.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                                dsTempB.InsertParameters.Add(new Parameter("CaseNoItems", DbType.String, vCaseNoItems));
                                dsTempB.InsertParameters.Add(new Parameter("Relationship", DbType.String, vRelationship));
                                dsTempB.InsertParameters.Add(new Parameter("RelCar_ID", DbType.String, vRelCar_ID));
                                dsTempB.InsertParameters.Add(new Parameter("EstimatedAmount", DbType.Int32, vEstimatedAmount));
                                dsTempB.InsertParameters.Add(new Parameter("ThirdInsurance", DbType.Int32, vThirdInsurance));
                                dsTempB.InsertParameters.Add(new Parameter("CompInsurance", DbType.Int32, vCompInsurance));
                                dsTempB.InsertParameters.Add(new Parameter("DriverSharing", DbType.Int32, vDriverSharing));
                                dsTempB.InsertParameters.Add(new Parameter("CompanySharing", DbType.Int32, vCompanySharing));
                                dsTempB.InsertParameters.Add(new Parameter("CarDamageAMT", DbType.Int32, vCarDamageAMT));
                                dsTempB.InsertParameters.Add(new Parameter("PersonDamageAMT", DbType.Int32, vPersonDamageAMT));
                                dsTempB.InsertParameters.Add(new Parameter("RelationComp", DbType.Int32, vRelationComp));
                                dsTempB.InsertParameters.Add(new Parameter("ReconciliationDate", DbType.Date, vReconciliationDate));
                                dsTempB.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemarkB));
                                dsTempB.InsertParameters.Add(new Parameter("PassengerInsu", DbType.Int32, vPassengerInsu));
                                dsTempB.Insert();
                            }
                        }
                    }
                }
                else
                {

                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 轉入維修項目代號                                     
        /// </summary>
        private void Import_FixSevNo()
        {
            string vFileName = fuExcel.FileName;
            if (vFileName != "")
            {
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vUploadFileName = vUploadPath + fuExcel.FileName;
                fuExcel.SaveAs(vUploadFileName);
                //string vExtName = Path.GetExtension(vFileName);
                vExtName = Path.GetExtension(vFileName);

                string vSQLStr_Temp = "";
                string vSevNo = "";
                string vSevName = "";
                double vSevHour = 0;
                string vBuDate_Str = DateTime.Today.Year.ToString("D4") + "/" + DateTime.Today.ToString("MM/dd");
                string vBuMan = vLoginID.Trim();

                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }

                if (vExtName == ".xls")
                {
                    HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                    HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                    for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                    {
                        HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                        if (vRowTemp_H != null)
                        {
                            vSevNo = vRowTemp_H.Cells[0].ToString().Trim();
                            if ((vSevNo != "") && (vSevNo != "維修代號"))
                            {
                                vSevName = vRowTemp_H.Cells[1].ToString().Trim();
                                vSevHour = (vRowTemp_H.Cells[2].ToString().Trim() != "") ? vRowTemp_H.Cells[2].NumericCellValue : 0;
                                using (SqlDataSource dsTemp = new SqlDataSource())
                                {
                                    dsTemp.ConnectionString = vConnStr;
                                    vSQLStr_Temp = "select Count(SevNo) RCount from FixSev where SevNo = '" + vSevNo + "' ";
                                    if (PF.GetValue(vConnStr, vSQLStr_Temp, "RCount") == "0")
                                    {
                                        dsTemp.InsertCommand = "INSERT INTO FixSev (SevNo, SevName, SevHour, SevPrice, SevPriceH, SevCost, BuDate, BuMan) " + Environment.NewLine +
                                                               "VALUES (@SevNo, @SevName, @SevHour, @SevPrice, @SevPriceH, @SevCost, @BuDate, @BuMan)";
                                        dsTemp.InsertParameters.Clear();
                                        dsTemp.InsertParameters.Add(new Parameter("SevNo", DbType.String, vSevNo));
                                        dsTemp.InsertParameters.Add(new Parameter("SevName", DbType.String, vSevName));
                                        dsTemp.InsertParameters.Add(new Parameter("SevHour", DbType.Double, vSevHour.ToString()));
                                        dsTemp.InsertParameters.Add(new Parameter("SevPrice", DbType.Double, "0"));
                                        dsTemp.InsertParameters.Add(new Parameter("SevPriceH", DbType.Double, "0"));
                                        dsTemp.InsertParameters.Add(new Parameter("SevCost", DbType.Double, "0"));
                                        dsTemp.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDate_Str));
                                        dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vBuMan));
                                        dsTemp.Insert();
                                    }
                                    else
                                    {
                                        dsTemp.UpdateCommand = "update FixSev set SevName = @SevName, SevHour = @SevHour where SevNo = @SevNo";
                                        dsTemp.UpdateParameters.Clear();
                                        dsTemp.UpdateParameters.Add(new Parameter("SevNo", DbType.String, vSevNo));
                                        dsTemp.UpdateParameters.Add(new Parameter("SevName", DbType.String, vSevName));
                                        dsTemp.UpdateParameters.Add(new Parameter("SevHour", DbType.Double, vSevHour.ToString()));
                                        dsTemp.Update();
                                    }
                                }
                            }
                        }
                    }
                }
                else if (vExtName == ".xlsx")
                {
                    XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                    XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                    for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                    {
                        XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                        if (vRowTemp_X != null)
                        {
                            vSevNo = vRowTemp_X.Cells[0].ToString().Trim();
                            if ((vSevNo != "") && (vSevNo != "維修代號"))
                            {
                                vSevName = vRowTemp_X.Cells[1].ToString().Trim();
                                vSevHour = (vRowTemp_X.Cells[2].ToString().Trim() != "") ? vRowTemp_X.Cells[2].NumericCellValue : 0;
                                using (SqlDataSource dsTemp = new SqlDataSource())
                                {
                                    dsTemp.ConnectionString = vConnStr;
                                    vSQLStr_Temp = "select Count(SevNo) RCount from FixSev where SevNo = '" + vSevNo + "' ";
                                    if (PF.GetValue(vConnStr, vSQLStr_Temp, "RCount") == "0")
                                    {
                                        dsTemp.InsertCommand = "INSERT INTO FixSev (SevNo, SevName, SevHour, SevPrice, SevPriceH, SevCost, BuDate, BuMan) " + Environment.NewLine +
                                                               "VALUES (@SevNo, @SevName, @SevHour, @SevPrice, @SevPriceH, @SevCost, @BuDate, @BuMan)";
                                        dsTemp.InsertParameters.Clear();
                                        dsTemp.InsertParameters.Add(new Parameter("SevNo", DbType.String, vSevNo));
                                        dsTemp.InsertParameters.Add(new Parameter("SevName", DbType.String, vSevName));
                                        dsTemp.InsertParameters.Add(new Parameter("SevHour", DbType.Double, vSevHour.ToString()));
                                        dsTemp.InsertParameters.Add(new Parameter("SevPrice", DbType.Double, "0"));
                                        dsTemp.InsertParameters.Add(new Parameter("SevPriceH", DbType.Double, "0"));
                                        dsTemp.InsertParameters.Add(new Parameter("SevCost", DbType.Double, "0"));
                                        dsTemp.InsertParameters.Add(new Parameter("BuDate", DbType.Date, vBuDate_Str));
                                        dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vBuMan));
                                        dsTemp.Insert();
                                    }
                                    else
                                    {
                                        dsTemp.UpdateCommand = "update FixSev set SevName = @SevName, SevHour = @SevHour where SevNo = @SevNo";
                                        dsTemp.UpdateParameters.Clear();
                                        dsTemp.UpdateParameters.Add(new Parameter("SevNo", DbType.String, vSevNo));
                                        dsTemp.UpdateParameters.Add(new Parameter("SevName", DbType.String, vSevName));
                                        dsTemp.UpdateParameters.Add(new Parameter("SevHour", DbType.Double, vSevHour.ToString()));
                                        dsTemp.Update();
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇要匯入的 EXCEL 檔，再重新執行匯入作業！')");
                Response.Write("</" + "Script>");
                fuExcel.Focus();
            }
        }

        /// <summary>
        /// 轉入客服案件
        /// </summary>
        private void Import_CustomService()
        {
            string vSQLStr = "";
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            //string vExtName = "";
            string vTempIndex = "";
            string vTempStr = "";
            string vServiceNo = "";
            string vBuildDate = "";
            string vBuildTime = "";
            string vBuildMan = "";
            string vServiceType = "";
            string vServiceTypeB = "";
            string vServiceTypeC = "";
            string vLinesNo = "";
            string vCar_ID = "";
            string vDriver = "";
            string vDriverName = "";
            string vBoardTime = "";
            string vBoardStation = "";
            string vGetoffTime = "";
            string vGetoffStation = "";
            string vServiceNote = "";
            string vIsNoContect = "";
            string vCivicName = "";
            string vCivicTelNo = "";
            string vCivicCellPhone = "";
            string vCivicAddress = "";
            string vCivicEMail = "";
            string vAthorityDepNo = "";
            string vAthorityDepNote = "";
            string vIsReplied = "";
            string vRemark = "";
            string vIsPending = "";
            string vAssignDate = "";
            string vAssignMan = "";
            int SheetCount = 0;
            DateTime dtTemp;

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
                            for (int i = sheetExcel_H.FirstRowNum + 1; i <= sheetExcel_H.LastRowNum; i++)
                            {
                                HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                if ((vRowTemp_H.GetCell(0).CellType != CellType.Blank) && (vRowTemp_H.GetCell(2).CellType != CellType.Blank))
                                {
                                    //開單日期時間
                                    vBuildDate = (vRowTemp_H.GetCell(0).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(0).DateCellValue.ToString() : "";
                                    //vBuildTime = (vRowTemp_H.GetCell(1).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(1).DateCellValue.ToString("hh:mm") : "";
                                    vBuildTime = DateTime.TryParse(vRowTemp_H.GetCell(1).ToString().Trim(), out dtTemp) ? dtTemp.ToString("hh:mm") : "";
                                    //開單人
                                    vTempStr = (vRowTemp_H.GetCell(2).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(2).ToString().Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vTempStr + "' and LeaveDay is null ";
                                    vBuildMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    //反映事項 (大類)
                                    vTempStr = (vRowTemp_H.GetCell(3).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(3).ToString().Trim() : "";
                                    vSQLStr = "select TypeLevel1 from CustomServiceType where TypeStep = '1' and TypeText = '" + vTempStr + "' ";
                                    vServiceType = PF.GetValue(vConnStr, vSQLStr, "TypeLevel1");
                                    //單號
                                    vTempStr = DateTime.Parse(vBuildDate).ToString("yyyyMMdd") + vServiceType;
                                    vSQLStr = "select Max(ServiceNo) MaxNo from CustomService where ServiceNo like '" + vTempStr + "%' ";
                                    vTempIndex = (PF.GetValue(vConnStr, vSQLStr, "MaxNo")).Replace(vTempStr, "");
                                    vTempIndex = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D3") : "001";
                                    vServiceNo = vTempStr + vTempIndex;
                                    //客訴事項 (中類)
                                    vTempStr = (vRowTemp_H.GetCell(4).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(4).ToString().Trim() : "";
                                    if (vTempStr != "")
                                    {
                                        vSQLStr = "select TypeLevel2 from CustomServiceType where TypeStep = '2' and TypeText = '" + vTempStr + "' ";
                                        vServiceTypeB = PF.GetValue(vConnStr, vSQLStr, "TypeLevel2");
                                    }
                                    else
                                    {
                                        vServiceTypeB = "";
                                    }
                                    //客訴類別 (小類)
                                    vTempStr = (vRowTemp_H.GetCell(5).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(5).ToString().Trim() : "";
                                    if (vTempStr != "")
                                    {
                                        vSQLStr = "select TypeLevel3 from CustomServiceType where TypeStep = '3' and TypeText = '" + vTempStr + "' ";
                                        vServiceTypeC = PF.GetValue(vConnStr, vSQLStr, "TypeLevel3");
                                    }
                                    else
                                    {
                                        vServiceTypeC = "";
                                    }
                                    //路線
                                    vLinesNo = (vRowTemp_H.GetCell(6).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(6).ToString().Trim() : "";
                                    //牌照號碼
                                    vCar_ID = (vRowTemp_H.GetCell(7).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(7).ToString().Trim() : "";
                                    //駕駛員
                                    vDriverName = (vRowTemp_H.GetCell(8).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(8).ToString().Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null and [Type] = '20' ";
                                    vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    //上下車時間及站牌
                                    vBoardTime = (vRowTemp_H.GetCell(9).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(9).ToString().Trim() : "";
                                    vGetoffTime = (vRowTemp_H.GetCell(10).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(10).ToString().Trim() : "";
                                    vBoardStation = (vRowTemp_H.GetCell(11).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(11).ToString().Trim() : "";
                                    vGetoffStation = (vRowTemp_H.GetCell(12).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(12).ToString().Trim() : "";
                                    //情況概述
                                    vServiceNote = (vRowTemp_H.GetCell(13).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(13).ToString().Trim() : "";
                                    //民眾連絡資料
                                    vCivicName = (vRowTemp_H.GetCell(14).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(14).ToString().Trim() : "";
                                    vCivicTelNo = (vRowTemp_H.GetCell(15).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(15).ToString().Trim() : "";
                                    vCivicCellPhone = (vRowTemp_H.GetCell(16).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(16).ToString().Trim() : "";
                                    vCivicAddress = (vRowTemp_H.GetCell(17).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(17).ToString().Trim() : "";
                                    vCivicEMail = (vRowTemp_H.GetCell(18).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(18).ToString().Trim() : "";
                                    vIsNoContect = ((vCivicName == "") && (vCivicTelNo == "") && (vCivicCellPhone == "") && (vCivicAddress == "") && (vCivicEMail == "")) ? "true" : "false";
                                    //權責單位
                                    vTempStr = (vRowTemp_H.GetCell(19).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(19).ToString().Trim() : "";
                                    vSQLStr = "select DepNo from Department where [Name] = '" + vTempStr + "' ";
                                    vAthorityDepNo = (PF.GetValue(vConnStr, vSQLStr, "DepNo") != "") ? PF.GetValue(vConnStr, vSQLStr, "DepNo") : vTempStr;
                                    //是否回覆
                                    vIsReplied = (vRowTemp_H.GetCell(20).ToString().Trim() != "XXXXXXXX") ? (vRowTemp_H.GetCell(20).ToString().Trim() == "是") ? "true" : "false" : "false";
                                    vRemark = (vRowTemp_H.GetCell(21).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(21).ToString().Trim() : "";
                                    //是否分發待查
                                    vIsPending = (vRowTemp_H.GetCell(22).ToString().Trim() != "XXXXXXXX") ? (vRowTemp_H.GetCell(22).ToString().Trim() == "是") ? "true" : "false" : "false";
                                    vAssignDate = (vRowTemp_H.GetCell(23).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(23).DateCellValue.ToString() : "";
                                    vTempStr = (vRowTemp_H.GetCell(24).ToString().Trim() != "XXXXXXXX") ? vRowTemp_H.GetCell(24).ToString().Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vTempStr + "' and LeaveDay is null ";
                                    vAssignMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");

                                    odsCustomServiceImport.InsertMethod = "InsertQuery_Import";
                                    odsCustomServiceImport.InsertParameters.Clear();

                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceNo", DbType.String, vServiceNo));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, vBuildDate));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BuildTime", DbType.String, vBuildTime));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vBuildMan));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceType", DbType.String, vServiceType));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceTypeB", DbType.String, vServiceTypeB));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceTypeC", DbType.String, vServiceTypeC));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BoardTime", DbType.String, vBoardTime));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BoardStation", DbType.String, vBoardStation));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("GetoffTime", DbType.String, vGetoffTime));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("GetoffStation", DbType.String, vGetoffStation));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceNote", DbType.String, vServiceNote));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("IsNoContect", DbType.Boolean, vIsNoContect));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicName", DbType.String, vCivicName));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicTelNo", DbType.String, vCivicTelNo));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicCellPhone", DbType.String, vCivicCellPhone));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicAddress", DbType.String, vCivicAddress));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicEMail", DbType.String, vCivicEMail));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("AthorityDepNo", DbType.String, vAthorityDepNo));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("AthorityDepNote", DbType.String, vAthorityDepNote));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("IsReplied", DbType.Boolean, vIsReplied));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("IsPending", DbType.Boolean, vIsPending));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("AssignDate", DbType.Date, vAssignDate));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("AssignMan", DbType.String, vAssignMan));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("IsClosed", DbType.Boolean, "false"));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CloseDate", DbType.Date, ""));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CloseMan", DbType.String, ""));

                                    odsCustomServiceImport.Insert();
                                }
                            }
                        }
                        break;
                    case ".xlsx":
                        //Excel 2010 或更新版本
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        SheetCount = wbExcel_X.NumberOfSheets;
                        for (int vSC = 0; vSC < SheetCount; vSC++)
                        {
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(vSC);
                            for (int i = sheetExcel_X.FirstRowNum + 1; i <= sheetExcel_X.LastRowNum; i++)
                            {
                                XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                if ((vRowTemp_X.GetCell(0).CellType != CellType.Blank) && (vRowTemp_X.GetCell(2).CellType != CellType.Blank))
                                {
                                    //開單日期時間
                                    vBuildDate = (vRowTemp_X.GetCell(0).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(0).DateCellValue.ToString() : "";
                                    //vBuildTime = (vRowTemp_X.GetCell(1).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(1).DateCellValue.ToString("hh:mm") : "";
                                    vBuildTime = DateTime.TryParse(vRowTemp_X.GetCell(1).ToString().Trim(), out dtTemp) ? dtTemp.ToString("hh:mm") : "";
                                    //開單人
                                    vTempStr = (vRowTemp_X.GetCell(2).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(2).ToString().Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vTempStr + "' and LeaveDay is null ";
                                    vBuildMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    //反映事項 (大類)
                                    vTempStr = (vRowTemp_X.GetCell(3).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(3).ToString().Trim() : "";
                                    vSQLStr = "select TypeLevel1 from CustomServiceType where TypeStep = '1' and TypeText = '" + vTempStr + "' ";
                                    vServiceType = PF.GetValue(vConnStr, vSQLStr, "TypeLevel1");
                                    //單號
                                    vTempStr = DateTime.Parse(vBuildDate).ToString("yyyyMMdd") + vServiceType;
                                    vSQLStr = "select Max(ServiceNo) MaxNo from CustomService where ServiceNo like '" + vTempStr + "%' ";
                                    vTempIndex = (PF.GetValue(vConnStr, vSQLStr, "MaxNo")).Replace(vTempStr, "");
                                    vTempIndex = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D3") : "001";
                                    vServiceNo = vTempStr + vTempIndex;
                                    //客訴事項 (中類)
                                    vTempStr = (vRowTemp_X.GetCell(4).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(4).ToString().Trim() : "";
                                    if (vTempStr != "")
                                    {
                                        vSQLStr = "select TypeLevel2 from CustomServiceType where TypeStep = '2' and TypeText = '" + vTempStr + "' ";
                                        vServiceTypeB = PF.GetValue(vConnStr, vSQLStr, "TypeLevel2");
                                    }
                                    else
                                    {
                                        vServiceTypeB = "";
                                    }
                                    //客訴類別 (小類)
                                    vTempStr = (vRowTemp_X.GetCell(5).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(5).ToString().Trim() : "";
                                    if (vTempStr != "")
                                    {
                                        vSQLStr = "select TypeLevel3 from CustomServiceType where TypeStep = '3' and TypeText = '" + vTempStr + "' ";
                                        vServiceTypeC = PF.GetValue(vConnStr, vSQLStr, "TypeLevel3");
                                    }
                                    else
                                    {
                                        vServiceTypeC = "";
                                    }
                                    //路線
                                    vLinesNo = (vRowTemp_X.GetCell(6).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(6).ToString().Trim() : "";
                                    //牌照號碼
                                    vCar_ID = (vRowTemp_X.GetCell(7).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(7).ToString().Trim() : "";
                                    //駕駛員
                                    vDriverName = (vRowTemp_X.GetCell(8).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(8).ToString().Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' and LeaveDay is null and [Type] = '20' ";
                                    vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                    //上下車時間及站牌
                                    vBoardTime = (vRowTemp_X.GetCell(9).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(9).ToString().Trim() : "";
                                    vGetoffTime = (vRowTemp_X.GetCell(10).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(10).ToString().Trim() : "";
                                    vBoardStation = (vRowTemp_X.GetCell(11).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(11).ToString().Trim() : "";
                                    vGetoffStation = (vRowTemp_X.GetCell(12).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(12).ToString().Trim() : "";
                                    //情況概述
                                    vServiceNote = (vRowTemp_X.GetCell(13).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(13).ToString().Trim() : "";
                                    //民眾連絡資料
                                    vCivicName = (vRowTemp_X.GetCell(14).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(14).ToString().Trim() : "";
                                    vCivicTelNo = (vRowTemp_X.GetCell(15).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(15).ToString().Trim() : "";
                                    vCivicCellPhone = (vRowTemp_X.GetCell(16).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(16).ToString().Trim() : "";
                                    vCivicAddress = (vRowTemp_X.GetCell(17).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(17).ToString().Trim() : "";
                                    vCivicEMail = (vRowTemp_X.GetCell(18).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(18).ToString().Trim() : "";
                                    vIsNoContect = ((vCivicName == "") && (vCivicTelNo == "") && (vCivicCellPhone == "") && (vCivicAddress == "") && (vCivicEMail == "")) ? "true" : "false";
                                    //權責單位
                                    vTempStr = (vRowTemp_X.GetCell(19).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(19).ToString().Trim() : "";
                                    vSQLStr = "select DepNo from Department where [Name] = '" + vTempStr + "' ";
                                    vAthorityDepNo = (PF.GetValue(vConnStr, vSQLStr, "DepNo") != "") ? PF.GetValue(vConnStr, vSQLStr, "DepNo") : vTempStr;
                                    //是否回覆
                                    vIsReplied = (vRowTemp_X.GetCell(20).ToString().Trim() != "XXXXXXXX") ? (vRowTemp_X.GetCell(20).ToString().Trim() == "是") ? "true" : "false" : "false";
                                    vRemark = (vRowTemp_X.GetCell(21).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(21).ToString().Trim() : "";
                                    //是否分發待查
                                    vIsPending = (vRowTemp_X.GetCell(22).ToString().Trim() != "XXXXXXXX") ? (vRowTemp_X.GetCell(22).ToString().Trim() == "是") ? "true" : "false" : "false";
                                    vAssignDate = (vRowTemp_X.GetCell(23).ToString().Trim().Replace("XXXXXXXX", "") != "") ? vRowTemp_X.GetCell(23).DateCellValue.ToString() : "";
                                    vTempStr = (vRowTemp_X.GetCell(24).ToString().Trim() != "XXXXXXXX") ? vRowTemp_X.GetCell(24).ToString().Trim() : "";
                                    vSQLStr = "select EmpNo from Employee where [Name] = '" + vTempStr + "' and LeaveDay is null ";
                                    vAssignMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");

                                    odsCustomServiceImport.InsertMethod = "InsertQuery_Import";
                                    odsCustomServiceImport.InsertParameters.Clear();

                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceNo", DbType.String, vServiceNo));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, vBuildDate));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BuildTime", DbType.String, vBuildTime));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vBuildMan));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceType", DbType.String, vServiceType));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceTypeB", DbType.String, vServiceTypeB));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceTypeC", DbType.String, vServiceTypeC));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BoardTime", DbType.String, vBoardTime));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("BoardStation", DbType.String, vBoardStation));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("GetoffTime", DbType.String, vGetoffTime));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("GetoffStation", DbType.String, vGetoffStation));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("ServiceNote", DbType.String, vServiceNote));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("IsNoContect", DbType.Boolean, vIsNoContect));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicName", DbType.String, vCivicName));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicTelNo", DbType.String, vCivicTelNo));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicCellPhone", DbType.String, vCivicCellPhone));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicAddress", DbType.String, vCivicAddress));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CivicEMail", DbType.String, vCivicEMail));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("AthorityDepNo", DbType.String, vAthorityDepNo));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("AthorityDepNote", DbType.String, vAthorityDepNote));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("IsReplied", DbType.Boolean, vIsReplied));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("IsPending", DbType.Boolean, vIsPending));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("AssignDate", DbType.Date, vAssignDate));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("AssignMan", DbType.String, vAssignMan));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("IsClosed", DbType.Boolean, "false"));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CloseDate", DbType.Date, ""));
                                    odsCustomServiceImport.InsertParameters.Add(new Parameter("CloseMan", DbType.String, ""));

                                    odsCustomServiceImport.Insert();
                                }
                            }
                        }
                        break;
                }
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('轉入完成！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 轉入保養預排
        /// </summary>
        private void Import_CarServicePre()
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            //string vExtName = "";
            string vSQLStr_Temp = "";
            string vGetTempData = "";
            string vTempStr = "";
            int vNewIndex = 0;

            //單據用變數
            string vPreSheetNo = "";
            string vDepNo_Temp = "";
            string vCar_ID_Temp = "";
            string vCarData_Str = "";
            string vDriver_Temp = "";
            string vServiceType = "";
            string vServiceDate_Str = "";
            DateTime vServiceDate;
            string vRemark_Temp = "";
            string vCalYM = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                vCalYM = eCalYear.Text.Trim() + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2");
                switch (vExtName)
                {
                    case ".xls":
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if ((vRowTemp_H != null) && (vRowTemp_H.Cells.Count > 0))
                            {
                                vPreSheetNo = vRowTemp_H.Cells[0].ToString().Trim();
                                if (vPreSheetNo != "預排單號")
                                {
                                    vCar_ID_Temp = vRowTemp_H.Cells[3].ToString().Trim();
                                    if (vCar_ID_Temp != "")
                                    {
                                        vSQLStr_Temp = "select (isnull(CompanyNo, '') + ',' + isnull(Driver, '')) CarData " + Environment.NewLine +
                                                       "  from Car_InfoA where Car_ID = '" + vCar_ID_Temp + "' ";
                                        vCarData_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "CarData");
                                        if (vCarData_Str.Trim() == "")
                                        {
                                            vCar_ID_Temp = vCar_ID_Temp.Replace("-U7", "-U-7");
                                            vSQLStr_Temp = "select (isnull(CompanyNo, '') + ',' + isnull(Driver, '')) CarData " + Environment.NewLine +
                                                           "  from Car_InfoA where Car_ID = '" + vCar_ID_Temp + "' ";
                                            vCarData_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "CarData");
                                        }
                                        string[] aCarData = vCarData_Str.Split(',');
                                        vDepNo_Temp = (vRowTemp_H.Cells[1].ToString().Trim() != "") ? vRowTemp_H.Cells[1].ToString().Trim() : aCarData[0];
                                        vDepNo_Temp = (vDepNo_Temp != "") ? "'" + vDepNo_Temp + "'" : "NULL";
                                        vDriver_Temp = (vRowTemp_H.Cells[4].ToString().Trim() != "") ?
                                                       Int32.Parse(vRowTemp_H.Cells[4].ToString().Trim()).ToString("D6") :
                                                       aCarData[1];
                                        vDriver_Temp = (vDriver_Temp != "") ? "'" + vDriver_Temp + "'" : "NULL";
                                        vSQLStr_Temp = "select ClassNo from DBDICB where FKey = '工作單A         FixworkA        SERVICE' and ClassTxt = '" + vRowTemp_H.Cells[7].ToString().Trim() + "' ";
                                        vServiceType = PF.GetValue(vConnStr, vSQLStr_Temp, "ClassNo");
                                        vServiceDate_Str = vRowTemp_H.Cells[9].ToString().Trim();
                                        vServiceDate = DateTime.Parse(vServiceDate_Str);
                                        vRemark_Temp = (vRowTemp_H.Cells[10].ToString().Trim() != "") ? "'" + vRowTemp_H.Cells[10].ToString().Trim() + "'" : "NULL";
                                        if (vPreSheetNo == "")
                                        {
                                            vGetTempData = "select MAX(ServicePreNo) MaxIndex from CarServicePre where ServicePreNo like '" + vCalYM + "%' ";
                                            vTempStr = PF.GetValue(vConnStr, vGetTempData, "MaxIndex");
                                            vNewIndex = (vTempStr != "") ? Int32.Parse(vTempStr.Replace(vCalYM, "")) + 1 : 1;
                                            vPreSheetNo = vCalYM + vNewIndex.ToString("D4");
                                            vSQLStr_Temp = "insert into CarServicePre " + Environment.NewLine +
                                                           "       (ServicePreNo, DepNo, Car_ID, Driver, ServiceType, " + Environment.NewLine +
                                                           "        LastDate, ServiceDate, OriDate, Remark, BuMan, BuDate, IsPassed)" + Environment.NewLine +
                                                           "values ('" + vPreSheetNo + "', " + vDepNo_Temp + ", '" + vCar_ID_Temp + "', " + vDriver_Temp + ", " + Environment.NewLine +
                                                           "        '" + vServiceType.Trim() + "', NULL, " + Environment.NewLine +
                                                           "        '" + vServiceDate.Year.ToString() + "/" + vServiceDate.ToString("MM/dd") + "', " + Environment.NewLine +
                                                           "        '" + vServiceDate.Year.ToString() + "/" + vServiceDate.ToString("MM/dd") + "', " + Environment.NewLine +
                                                           "        " + vRemark_Temp + ", '" + vLoginID + "', GetDate(), 'X')";
                                        }
                                        else
                                        {
                                            vSQLStr_Temp = "update CarServicePre " + Environment.NewLine +
                                                           "  set ServiceDate = '" + vServiceDate_Str + "', Remark = " + vRemark_Temp + ", ModifyMan = '" + vLoginID + "', ModifyDate = GetDate() " + Environment.NewLine +
                                                           "where ServicePreNo = '" + vPreSheetNo + "' ";
                                        }
                                        try
                                        {
                                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                                        }
                                        catch (Exception eMessage)
                                        {
                                            //Response.Write("<Script language='Javascript'>");
                                            //Response.Write("alert('" + eMessage.Message + Environment.NewLine + vTempStr + "')");
                                            //Response.Write("</" + "Script>");
                                            lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    case ".xlsx":
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if ((vRowTemp_X != null) && (vRowTemp_X.Cells.Count > 0))
                            {
                                vPreSheetNo = vRowTemp_X.Cells[0].ToString().Trim();
                                if (vPreSheetNo != "預排單號")
                                {
                                    vCar_ID_Temp = vRowTemp_X.Cells[3].ToString().Trim();
                                    if (vCar_ID_Temp != "")
                                    {
                                        vSQLStr_Temp = "select (isnull(CompanyNo, '') + ',' + isnull(Driver, '')) CarData " + Environment.NewLine +
                                                       "  from Car_InfoA where Car_ID = '" + vCar_ID_Temp + "' ";
                                        vCarData_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "CarData");
                                        if (vCarData_Str.Trim() == "")
                                        {
                                            vCar_ID_Temp = vCar_ID_Temp.Replace("-U7", "-U-7");
                                            vSQLStr_Temp = "select (isnull(CompanyNo, '') + ',' + isnull(Driver, '')) CarData " + Environment.NewLine +
                                                           "  from Car_InfoA where Car_ID = '" + vCar_ID_Temp + "' ";
                                            vCarData_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "CarData");
                                        }
                                        string[] aCarData = vCarData_Str.Split(',');
                                        vDepNo_Temp = (vRowTemp_X.Cells[1].ToString().Trim() != "") ? vRowTemp_X.Cells[1].ToString().Trim() :
                                                      (aCarData.Length > 0) ? aCarData[0] : "";
                                        vDepNo_Temp = (vDepNo_Temp != "") ? "'" + vDepNo_Temp + "'" : "NULL";
                                        vDriver_Temp = (vRowTemp_X.Cells[4].ToString().Trim() != "") ?
                                                       Int32.Parse(vRowTemp_X.Cells[4].ToString().Trim()).ToString("D6") :
                                                       (aCarData.Length > 0) ? aCarData[1] : "";
                                        vDriver_Temp = (vDriver_Temp != "") ? "'" + vDriver_Temp + "'" : "NULL";
                                        vSQLStr_Temp = "select ClassNo from DBDICB where FKey = '工作單A         FixworkA        SERVICE' and ClassTxt = '" + vRowTemp_X.Cells[7].ToString().Trim() + "' ";
                                        vServiceType = PF.GetValue(vConnStr, vSQLStr_Temp, "ClassNo");
                                        vServiceDate_Str = vRowTemp_X.Cells[9].ToString().Trim();
                                        vServiceDate = DateTime.Parse(vServiceDate_Str);
                                        vRemark_Temp = (vRowTemp_X.Cells[10].ToString().Trim() != "") ? "'" + vRowTemp_X.Cells[10].ToString().Trim() + "'" : "NULL";
                                        if (vPreSheetNo == "")
                                        {
                                            vGetTempData = "select MAX(ServicePreNo) MaxIndex from CarServicePre where ServicePreNo like '" + vCalYM + "%' ";
                                            vTempStr = PF.GetValue(vConnStr, vGetTempData, "MaxIndex");
                                            vNewIndex = (vTempStr != "") ? Int32.Parse(vTempStr.Replace(vCalYM, "")) + 1 : 1;
                                            vPreSheetNo = vCalYM + vNewIndex.ToString("D4");
                                            vSQLStr_Temp = "insert into CarServicePre " + Environment.NewLine +
                                                           "       (ServicePreNo, DepNo, Car_ID, Driver, ServiceType, " + Environment.NewLine +
                                                           "        LastDate, ServiceDate, OriDate, Remark, BuMan, BuDate, IsPassed)" + Environment.NewLine +
                                                           "values ('" + vPreSheetNo + "', " + vDepNo_Temp + ", '" + vCar_ID_Temp + "', " + vDriver_Temp + ", " + Environment.NewLine +
                                                           "        '" + vServiceType + "', NULL, " + Environment.NewLine +
                                                           "        '" + vServiceDate.Year.ToString() + "/" + vServiceDate.ToString("MM/dd") + "', " + Environment.NewLine +
                                                           "        '" + vServiceDate.Year.ToString() + "/" + vServiceDate.ToString("MM/dd") + "', " + Environment.NewLine +
                                                           "        " + vRemark_Temp + ", '" + vLoginID + "', GetDate(), 'X')";
                                        }
                                        else
                                        {
                                            vSQLStr_Temp = "update CarServicePre " + Environment.NewLine +
                                                           "  set ServiceDate = '" + vServiceDate_Str + "', Remark = " + vRemark_Temp + ", ModifyMan = '" + vLoginID + "', ModifyDate = GetDate() " + Environment.NewLine +
                                                           "where ServicePreNo = '" + vPreSheetNo + "' ";
                                        }
                                        try
                                        {
                                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                                        }
                                        catch (Exception eMessage)
                                        {
                                            //Response.Write("<Script language='Javascript'>");
                                            //Response.Write("alert('" + eMessage.Message + Environment.NewLine + vTempStr + "')");
                                            //Response.Write("</" + "Script>");
                                            lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                        }
                                    }
                                }
                            }
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
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 轉入合約書資料
        /// </summary>
        private void Import_ContractManager()
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            //string vExtName = "";
            //string vGetTempData = "";
            string vTempStr = "";
            int vNewIndex = 0;
            int vSheetCount = 0;
            string vSheetName = "";
            bool vIsPass = false;
            string vErrorStr = "";
            string vAssignYM = "";
            string vOldIndexNo = "";

            //欄位內容變數
            string vIndexNo = "";
            string vWorkDep = "";
            string vWorkDepName = "";
            string vResponsibleDep = "";
            string vResponsibleDepName = "";
            string vContractNo = "";
            string vInvoiceNo = "";
            string vCustomName = "";
            string vContractText = "";
            string vAssignDateStr = "";
            DateTime vAssignDate;
            string vEffectiveDateSStr = "";
            DateTime vEffectiveDateS;
            string vEffectiveDateEStr = "";
            DateTime vEffectiveDateE;
            string vIsClose = "";
            string vAmountStr = "";
            int vAmount = 0;
            string vTaxStr = "";
            int vTax = 0;
            string vTotalAmountStr = "";
            int vTotalAmount = 0;
            string vStampDutyStr = "";
            int vStampDuty = 0;
            string vRemark = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;
                    case ".xls": //EXCEL 2003 之前版本
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        vSheetCount = wbExcel_H.NumberOfSheets; //取得工作表數量
                        for (int vSC = 0; vSC < vSheetCount; vSC++)
                        {
                            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(vSC);
                            vSheetName = sheetExcel_H.SheetName;
                            for (int i = sheetExcel_H.FirstRowNum + 1; i <= sheetExcel_H.LastRowNum; i++)
                            {
                                HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                vIsPass = true;
                                vErrorStr = "";
                                if ((vRowTemp_H != null) && (vRowTemp_H.Cells.Count > 0))
                                {
                                    vWorkDepName = vRowTemp_H.Cells[0].ToString().Trim();
                                    if ((vWorkDepName != "") && (vWorkDepName != "承辦單位"))
                                    {
                                        vContractNo = vRowTemp_H.Cells[2].ToString().Trim(); //合約編號
                                        vErrorStr = ((vErrorStr != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + " 筆記錄無合約編號" :
                                                    ((vErrorStr == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + " 筆記錄無合約編號" :
                                                    vErrorStr;
                                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vWorkDepName + "' ";
                                        vWorkDep = PF.GetValue(vConnStr, vTempStr, "DepNo"); //承辦單位
                                        if ((vWorkDepName != "") && (vWorkDep.Trim() == ""))
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "承辦單位內容有誤_[" + vWorkDepName + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄承辦單位內容有誤_[" + vWorkDepName + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "承辦單位內容有誤_[" + vWorkDepName + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄承辦單位內容有誤_[" + vWorkDepName + "]" : "";
                                            vWorkDepName = "";
                                        }
                                        else if ((vWorkDepName != "") && (vWorkDep != ""))
                                        {
                                            vWorkDep = "'" + vWorkDep + "'";
                                        }
                                        else
                                        {
                                            vWorkDep = "NULL";
                                        }

                                        vResponsibleDepName = vRowTemp_H.Cells[1].ToString().Trim();
                                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vResponsibleDepName + "' ";
                                        vResponsibleDep = PF.GetValue(vConnStr, vTempStr, "DepNo"); //主責單位
                                        if ((vResponsibleDepName != "") && (vResponsibleDep.Trim() == ""))
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "主責單位內容有誤_[" + vResponsibleDepName + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄主責單位內容有誤_[" + vResponsibleDepName + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "主責單位內容有誤_[" + vResponsibleDepName + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄主責單位內容有誤_[" + vResponsibleDepName + "]" : "";
                                            vResponsibleDepName = "";
                                        }
                                        else if ((vResponsibleDepName != "") && (vResponsibleDep != ""))
                                        {
                                            vResponsibleDep = "'" + vResponsibleDep + "'";
                                        }
                                        else
                                        {
                                            vResponsibleDep = "NULL";
                                        }
                                        vInvoiceNo = (vRowTemp_H.Cells[3].ToString().Trim().Length > 10) ?
                                                     vRowTemp_H.Cells[3].ToString().Trim().Substring(0, 10) :
                                                     vRowTemp_H.Cells[3].ToString().Trim(); //對方統一編號
                                        vInvoiceNo = (vInvoiceNo != "") ? "'" + vInvoiceNo + "'" : "NULL";
                                        vCustomName = (vRowTemp_H.Cells[4].ToString().Trim().Length > 64) ?
                                                      vRowTemp_H.Cells[4].ToString().Trim().Substring(0, 64) :
                                                      vRowTemp_H.Cells[4].ToString().Trim(); //對方名稱
                                        vCustomName = (vCustomName != "") ? "'" + vCustomName + "'" : "NULL";
                                        vContractText = vRowTemp_H.Cells[5].ToString().Trim(); //工程地點、內容或名稱
                                        vContractText = (vContractText != "") ? "'" + vContractText + "'" : "NULL";
                                        vAssignDateStr = vRowTemp_H.Cells[6].ToString().Trim().Replace(".", "/"); //立契 (據) 日期
                                        if ((vAssignDateStr != "") && (DateTime.TryParse(vAssignDateStr, out vAssignDate)))
                                        {
                                            vAssignDate = (vAssignDate.Year < 1911) ? vAssignDate.AddYears(1911) : vAssignDate;
                                            vAssignDateStr = "'" + vAssignDate.Year.ToString() + "/" + vAssignDate.ToString("MM/dd") + "'";
                                            vAssignYM = vAssignDate.Year.ToString() + vAssignDate.Month.ToString("D2");
                                            vTempStr = "select max(IndexNo) MaxIndex from ContractManager where IndexNo like '" + vAssignYM + "%' ";
                                            vOldIndexNo = PF.GetValue(vConnStr, vTempStr, "MaxIndex");
                                            vNewIndex = (vOldIndexNo != "") ? Int32.Parse((vOldIndexNo.Replace(vAssignYM + "CM", ""))) + 1 : 1;
                                            vIndexNo = vAssignYM + "CM" + vNewIndex.ToString("D4");
                                        }
                                        else if (vAssignDateStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "立契 (據) 日期有誤_[" + vAssignDateStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄立契 (據) 日期有誤_[" + vAssignDateStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "立契 (據) 日期有誤_[" + vAssignDateStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄立契 (據) 日期有誤_[" + vAssignDateStr + "]" : "";
                                            vAssignDateStr = "";
                                        }
                                        else if (vAssignDateStr == "")
                                        {
                                            vAssignDate = DateTime.Today;
                                            vAssignDateStr = "NULL";
                                            vAssignYM = vAssignDate.Year.ToString() + vAssignDate.Month.ToString("D2");
                                            vTempStr = "select max(IndexNo) MaxIndex from ContractManager where IndexNo like '" + vAssignYM + "%' ";
                                            vOldIndexNo = PF.GetValue(vConnStr, vTempStr, "MaxIndex");
                                            vNewIndex = (vOldIndexNo != "") ? Int32.Parse((vOldIndexNo.Replace(vAssignYM + "CM", ""))) + 1 : 1;
                                            vIndexNo = vAssignYM + "CM" + vNewIndex.ToString("D4");
                                        }
                                        vEffectiveDateSStr = vRowTemp_H.Cells[7].ToString().Trim().Replace(".", "/"); //合約書期間 (起)
                                        if ((vEffectiveDateSStr != "") && (DateTime.TryParse(vEffectiveDateSStr, out vEffectiveDateS)))
                                        {
                                            vEffectiveDateS = (vEffectiveDateS.Year < 1911) ? vEffectiveDateS.AddYears(1911) : vEffectiveDateS;
                                            vEffectiveDateSStr = "'" + vEffectiveDateS.Year.ToString() + "/" + vEffectiveDateS.ToString("MM/dd") + "'";
                                        }
                                        else if (vEffectiveDateSStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "合約書期間 (起) 有誤_[" + vEffectiveDateSStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合約書期間 (起) 有誤_[" + vEffectiveDateSStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "合約書期間 (起) 有誤_[" + vEffectiveDateSStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合約書期間 (起) 有誤_[" + vEffectiveDateSStr + "]" : "";
                                            vEffectiveDateSStr = "";
                                        }
                                        else
                                        {
                                            vEffectiveDateSStr = "NULL";
                                        }
                                        vEffectiveDateEStr = vRowTemp_H.Cells[8].ToString().Trim().Replace(".", "/"); //合約書期間 (迄)
                                        if ((vEffectiveDateEStr != "") && (DateTime.TryParse(vEffectiveDateEStr, out vEffectiveDateE)))
                                        {
                                            vEffectiveDateE = (vEffectiveDateE.Year < 1911) ? vEffectiveDateE.AddYears(1911) : vEffectiveDateE;
                                            vEffectiveDateEStr = "'" + vEffectiveDateE.Year.ToString() + "/" + vEffectiveDateE.ToString("MM/dd") + "'";
                                        }
                                        else if (vEffectiveDateEStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "合約書期間 (迄) 有誤_[" + vEffectiveDateEStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合約書期間 (迄) 有誤_[" + vEffectiveDateEStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "合約書期間 (迄) 有誤_[" + vEffectiveDateEStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合約書期間 (迄) 有誤_[" + vEffectiveDateEStr + "]" : "";
                                            vEffectiveDateEStr = "";
                                        }
                                        else
                                        {
                                            vEffectiveDateEStr = "NULL";
                                        }
                                        vIsClose = vRowTemp_H.Cells[9].ToString().Trim().Replace("Y", "V");
                                        vIsClose = (vIsClose != "") ? "'" + vIsClose + "'" : "'X'";
                                        //vAmountStr = vRowTemp_H.Cells[10].ToString().Trim().Replace(",", "");
                                        vAmountStr = vRowTemp_H.Cells[10].NumericCellValue.ToString();
                                        if ((vAmountStr != "") && Int32.TryParse(vAmountStr, out vAmount))
                                        {
                                            vAmountStr = vAmount.ToString();
                                        }
                                        else if (vAmountStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "憑證金額有誤_[" + vAmountStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄憑證金額有誤_[" + vAmountStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "憑證金額有誤_[" + vAmountStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄憑證金額有誤_[" + vAmountStr + "]" : "";
                                            vAmountStr = "";
                                        }
                                        else
                                        {
                                            vAmountStr = "NULL";
                                        }
                                        //vTaxStr = vRowTemp_H.Cells[11].ToString().Trim().Replace(",", "");
                                        vTaxStr = vRowTemp_H.Cells[11].NumericCellValue.ToString();
                                        if ((vTaxStr != "") && Int32.TryParse(vTaxStr, out vTax))
                                        {
                                            vTaxStr = vTax.ToString();
                                        }
                                        else if (vTaxStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "憑證稅額有誤_[" + vTaxStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄憑證稅額有誤_[" + vTaxStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "憑證稅額有誤_[" + vTaxStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄憑證稅額有誤_[" + vTaxStr + "]" : "";
                                            vTaxStr = "";
                                        }
                                        else
                                        {
                                            vTaxStr = "NULL";
                                        }
                                        //vTotalAmountStr = vRowTemp_H.Cells[12].ToString().Trim().Replace(",", "");
                                        vTotalAmountStr = vRowTemp_H.Cells[12].NumericCellValue.ToString();
                                        if ((vTotalAmountStr != "") && Int32.TryParse(vTotalAmountStr, out vTotalAmount))
                                        {
                                            vTotalAmountStr = vTotalAmount.ToString();
                                        }
                                        else if (vTotalAmountStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "合計金額有誤_[" + vTotalAmountStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合計金額有誤_[" + vTotalAmountStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "合計金額有誤_[" + vTotalAmountStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合計金額有誤_[" + vTotalAmountStr + "]" : "";
                                            vTotalAmountStr = "";
                                        }
                                        else
                                        {
                                            vTotalAmountStr = "NULL";
                                        }
                                        vStampDutyStr = vRowTemp_H.Cells[13].ToString().Trim().Replace(",", "");
                                        if ((vStampDutyStr != "") && Int32.TryParse(vStampDutyStr, out vStampDuty))
                                        {
                                            vStampDutyStr = vStampDuty.ToString();
                                        }
                                        else if (vStampDutyStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "印花稅額有誤_[" + vStampDutyStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄印花稅額有誤_[" + vStampDutyStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "印花稅額有誤_[" + vStampDutyStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄印花稅額有誤_[" + vStampDutyStr + "]" : "";
                                            vStampDutyStr = "";
                                        }
                                        else
                                        {
                                            vStampDutyStr = "NULL";
                                        }
                                        vRemark = vRowTemp_H.Cells[14].ToString().Trim();
                                        vRemark = (vRemark != "") ? "'" + vRemark + "'" : "NULL";
                                        //寫入資料
                                        if (vIsPass)
                                        {
                                            vTempStr = "insert into ContractManager (IndexNo, ContractNo, WorkDep, ResponsibleDep, " + Environment.NewLine +
                                                       "       InvoiceNo, CustomName, ContractText, AssignDate, EffectiveDateS, EffectiveDateE, " + Environment.NewLine +
                                                       "       IsClose, Amount, Tax, TotalAmount, StampDuty, Remark, BuMan, BuDate)" + Environment.NewLine +
                                                       "values('" + vIndexNo + "', '" + vContractNo + "', " + vWorkDep + ", " + vResponsibleDep + ", " +
                                                       vInvoiceNo + ", " + vCustomName + ", " + vContractText + ", " + vAssignDateStr + ", " + vEffectiveDateSStr + ", " + vEffectiveDateEStr + ", " +
                                                       vIsClose + ", " + vAmountStr + ", " + vTaxStr + ", " + vTotalAmountStr + ", " + vStampDutyStr + ", " + vRemark + ", '" + vLoginID + "', GetDate())";
                                            try
                                            {
                                                PF.ExecSQL(vConnStr, vTempStr);
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
                                            eShowError.Text = (vErrorStr != "") ? eShowError.Text + Environment.NewLine + vErrorStr : eShowError.Text;
                                            eShowError.Visible = (vErrorStr.Length > 0);
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    case ".xlsx": //EXCEL 2010 之後版本
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        vSheetCount = wbExcel_X.NumberOfSheets; //取得工作表數量
                        for (int vSC = 0; vSC < vSheetCount; vSC++)
                        {
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(vSC);
                            vSheetName = sheetExcel_X.SheetName;
                            for (int i = sheetExcel_X.FirstRowNum + 1; i <= sheetExcel_X.LastRowNum; i++)
                            {
                                XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                if ((vRowTemp_X != null) && (vRowTemp_X.Cells.Count > 0))
                                {
                                    vIsPass = true;
                                    vErrorStr = "";
                                    vWorkDepName = vRowTemp_X.Cells[0].ToString().Trim();
                                    if ((vWorkDepName != "") && (vWorkDepName != "承辦單位"))
                                    {
                                        vContractNo = vRowTemp_X.Cells[2].ToString().Trim(); //合約編號
                                        vErrorStr = ((vErrorStr != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + " 筆記錄無合約編號" :
                                                    ((vErrorStr == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + " 筆記錄無合約編號" :
                                                    vErrorStr;
                                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vWorkDepName + "' ";
                                        vWorkDep = PF.GetValue(vConnStr, vTempStr, "DepNo"); //承辦單位
                                        if ((vWorkDepName != "") && (vWorkDep.Trim() == ""))
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "承辦單位內容有誤_[" + vWorkDepName + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄承辦單位內容有誤_[" + vWorkDepName + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "承辦單位內容有誤_[" + vWorkDepName + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄承辦單位內容有誤_[" + vWorkDepName + "]" : "";
                                            vWorkDepName = "";
                                        }
                                        else if ((vWorkDepName != "") && (vWorkDep != ""))
                                        {
                                            vWorkDep = "'" + vWorkDep + "'";
                                        }
                                        else
                                        {
                                            vWorkDep = "NULL";
                                        }

                                        vResponsibleDepName = vRowTemp_X.Cells[1].ToString().Trim();
                                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vResponsibleDepName + "' ";
                                        vResponsibleDep = PF.GetValue(vConnStr, vTempStr, "DepNo"); //主責單位
                                        if ((vResponsibleDepName != "") && (vResponsibleDep.Trim() == ""))
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "主責單位內容有誤_[" + vResponsibleDepName + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄主責單位內容有誤_[" + vResponsibleDepName + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "主責單位內容有誤_[" + vResponsibleDepName + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄主責單位內容有誤_[" + vResponsibleDepName + "]" : "";
                                            vResponsibleDepName = "";
                                        }
                                        else if ((vResponsibleDepName != "") && (vResponsibleDep != ""))
                                        {
                                            vResponsibleDep = "'" + vResponsibleDep + "'";
                                        }
                                        else
                                        {
                                            vResponsibleDep = "NULL";
                                        }
                                        vInvoiceNo = (vRowTemp_X.Cells[3].ToString().Trim().Length > 10) ?
                                                     vRowTemp_X.Cells[3].ToString().Trim().Substring(0, 10) :
                                                     vRowTemp_X.Cells[3].ToString().Trim(); //對方統一編號
                                        vInvoiceNo = (vInvoiceNo != "") ? "'" + vInvoiceNo + "'" : "NULL";
                                        vCustomName = (vRowTemp_X.Cells[4].ToString().Trim().Length > 64) ?
                                                      vRowTemp_X.Cells[4].ToString().Trim().Substring(0, 64) :
                                                      vRowTemp_X.Cells[4].ToString().Trim(); //對方名稱
                                        vCustomName = (vCustomName != "") ? "'" + vCustomName + "'" : "NULL";
                                        vContractText = vRowTemp_X.Cells[5].ToString().Trim(); //工程地點、內容或名稱
                                        vContractText = (vContractText != "") ? "'" + vContractText + "'" : "NULL";
                                        vAssignDateStr = vRowTemp_X.Cells[6].ToString().Trim().Replace(".", "/"); //立契 (據) 日期
                                        if ((vAssignDateStr != "") && (DateTime.TryParse(vAssignDateStr, out vAssignDate)))
                                        {
                                            vAssignDate = (vAssignDate.Year < 1911) ? vAssignDate.AddYears(1911) : vAssignDate;
                                            vAssignDateStr = "'" + vAssignDate.Year.ToString() + "/" + vAssignDate.ToString("MM/dd") + "'";
                                            vAssignYM = vAssignDate.Year.ToString() + vAssignDate.Month.ToString("D2");
                                            vTempStr = "select max(IndexNo) MaxIndex from ContractManager where IndexNo like '" + vAssignYM + "%' ";
                                            vOldIndexNo = PF.GetValue(vConnStr, vTempStr, "MaxIndex");
                                            vNewIndex = (vOldIndexNo != "") ? Int32.Parse((vOldIndexNo.Replace(vAssignYM + "CM", ""))) + 1 : 1;
                                            vIndexNo = vAssignYM + "CM" + vNewIndex.ToString("D4");
                                        }
                                        else if (vAssignDateStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "立契 (據) 日期有誤_[" + vAssignDateStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄立契 (據) 日期有誤_[" + vAssignDateStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "立契 (據) 日期有誤_[" + vAssignDateStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄立契 (據) 日期有誤_[" + vAssignDateStr + "]" : "";
                                            vAssignDateStr = "";
                                        }
                                        else if (vAssignDateStr == "")
                                        {
                                            vAssignDate = DateTime.Today;
                                            vAssignDateStr = "NULL";
                                            vAssignYM = vAssignDate.Year.ToString() + vAssignDate.Month.ToString("D2");
                                            vTempStr = "select max(IndexNo) MaxIndex from ContractManager where IndexNo like '" + vAssignYM + "%' ";
                                            vOldIndexNo = PF.GetValue(vConnStr, vTempStr, "MaxIndex");
                                            vNewIndex = (vOldIndexNo != "") ? Int32.Parse((vOldIndexNo.Replace(vAssignYM + "CM", ""))) + 1 : 1;
                                            vIndexNo = vAssignYM + "CM" + vNewIndex.ToString("D4");
                                        }
                                        vEffectiveDateSStr = vRowTemp_X.Cells[7].ToString().Trim().Replace(".", "/"); //合約書期間 (起)
                                        if ((vEffectiveDateSStr != "") && (DateTime.TryParse(vEffectiveDateSStr, out vEffectiveDateS)))
                                        {
                                            vEffectiveDateS = (vEffectiveDateS.Year < 1911) ? vEffectiveDateS.AddYears(1911) : vEffectiveDateS;
                                            vEffectiveDateSStr = "'" + vEffectiveDateS.Year.ToString() + "/" + vEffectiveDateS.ToString("MM/dd") + "'";
                                        }
                                        else if (vEffectiveDateSStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "合約書期間 (起) 有誤_[" + vEffectiveDateSStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合約書期間 (起) 有誤_[" + vEffectiveDateSStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "合約書期間 (起) 有誤_[" + vEffectiveDateSStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合約書期間 (起) 有誤_[" + vEffectiveDateSStr + "]" : "";
                                            vEffectiveDateSStr = "";
                                        }
                                        else
                                        {
                                            vEffectiveDateSStr = "NULL";
                                        }
                                        vEffectiveDateEStr = vRowTemp_X.Cells[8].ToString().Trim().Replace(".", "/"); //合約書期間 (迄)
                                        if ((vEffectiveDateEStr != "") && (DateTime.TryParse(vEffectiveDateEStr, out vEffectiveDateE)))
                                        {
                                            vEffectiveDateE = (vEffectiveDateE.Year < 1911) ? vEffectiveDateE.AddYears(1911) : vEffectiveDateE;
                                            vEffectiveDateEStr = "'" + vEffectiveDateE.Year.ToString() + "/" + vEffectiveDateE.ToString("MM/dd") + "'";
                                        }
                                        else if (vEffectiveDateEStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "合約書期間 (迄) 有誤_[" + vEffectiveDateEStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合約書期間 (迄) 有誤_[" + vEffectiveDateEStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "合約書期間 (迄) 有誤_[" + vEffectiveDateEStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合約書期間 (迄) 有誤_[" + vEffectiveDateEStr + "]" : "";
                                            vEffectiveDateEStr = "";
                                        }
                                        else
                                        {
                                            vEffectiveDateEStr = "NULL";
                                        }
                                        vIsClose = vRowTemp_X.Cells[9].ToString().Trim().Replace("Y", "V");
                                        vIsClose = (vIsClose != "") ? "'" + vIsClose + "'" : "'X'";
                                        //vAmountStr = vRowTemp_X.Cells[10].ToString().Trim().Replace(",", "");
                                        vAmountStr = vRowTemp_X.Cells[10].NumericCellValue.ToString();
                                        if ((vAmountStr != "") && Int32.TryParse(vAmountStr, out vAmount))
                                        {
                                            vAmountStr = vAmount.ToString();
                                        }
                                        else if (vAmountStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "憑證金額有誤_[" + vAmountStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄憑證金額有誤_[" + vAmountStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "憑證金額有誤_[" + vAmountStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄憑證金額有誤_[" + vAmountStr + "]" : "";
                                            vAmountStr = "";
                                        }
                                        else
                                        {
                                            vAmountStr = "NULL";
                                        }
                                        //vTaxStr = vRowTemp_X.Cells[11].ToString().Trim().Replace(",", "");
                                        vTaxStr = vRowTemp_X.Cells[11].NumericCellValue.ToString();
                                        if ((vTaxStr != "") && Int32.TryParse(vTaxStr, out vTax))
                                        {
                                            vTaxStr = vTax.ToString();
                                        }
                                        else if (vTaxStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "憑證稅額有誤_[" + vTaxStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄憑證稅額有誤_[" + vTaxStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "憑證稅額有誤_[" + vTaxStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄憑證稅額有誤_[" + vTaxStr + "]" : "";
                                            vTaxStr = "";
                                        }
                                        else
                                        {
                                            vTaxStr = "NULL";
                                        }
                                        //vTotalAmountStr = vRowTemp_X.Cells[12].ToString().Trim().Replace(",", "");
                                        vTotalAmountStr = vRowTemp_X.Cells[12].NumericCellValue.ToString();
                                        if ((vTotalAmountStr != "") && Int32.TryParse(vTotalAmountStr, out vTotalAmount))
                                        {
                                            vTotalAmountStr = vTotalAmount.ToString();
                                        }
                                        else if (vTotalAmountStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "合計金額有誤_[" + vTotalAmountStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合計金額有誤_[" + vTotalAmountStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "合計金額有誤_[" + vTotalAmountStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄合計金額有誤_[" + vTotalAmountStr + "]" : "";
                                            vTotalAmountStr = "";
                                        }
                                        else
                                        {
                                            vTotalAmountStr = "NULL";
                                        }
                                        //vStampDutyStr = vRowTemp_X.Cells[13].ToString().Trim().Replace(",", "");
                                        vStampDutyStr = vRowTemp_X.Cells[13].NumericCellValue.ToString();
                                        if ((vStampDutyStr != "") && Int32.TryParse(vStampDutyStr, out vStampDuty))
                                        {
                                            vStampDutyStr = vStampDuty.ToString();
                                        }
                                        else if (vStampDutyStr != "")
                                        {
                                            vIsPass = false;
                                            vErrorStr = ((vErrorStr.Trim() != "") && (vContractNo != "")) ? vErrorStr + Environment.NewLine + vSheetName + "_合約編號：" + vContractNo + "印花稅額有誤_[" + vStampDutyStr + "]" :
                                                        ((vErrorStr.Trim() != "") && (vContractNo == "")) ? vErrorStr + Environment.NewLine + vSheetName + "_第 " + (i + 1).ToString() + "筆記錄印花稅額有誤_[" + vStampDutyStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo != "")) ? vSheetName + "_合約編號：" + vContractNo + "印花稅額有誤_[" + vStampDutyStr + "]" :
                                                        ((vErrorStr.Trim() == "") && (vContractNo == "")) ? vSheetName + "_第 " + (i + 1).ToString() + "筆記錄印花稅額有誤_[" + vStampDutyStr + "]" : "";
                                            vStampDutyStr = "";
                                        }
                                        else
                                        {
                                            vStampDutyStr = "NULL";
                                        }
                                        vRemark = vRowTemp_X.Cells[14].ToString().Trim();
                                        vRemark = (vRemark != "") ? "'" + vRemark + "'" : "NULL";
                                        //寫入資料
                                        if (vIsPass)
                                        {
                                            vTempStr = "insert into ContractManager (IndexNo, ContractNo, WorkDep, ResponsibleDep, " + Environment.NewLine +
                                                       "       InvoiceNo, CustomName, ContractText, AssignDate, EffectiveDateS, EffectiveDateE, " + Environment.NewLine +
                                                       "       IsClose, Amount, Tax, TotalAmount, StampDuty, Remark, BuMan, BuDate)" + Environment.NewLine +
                                                       "values('" + vIndexNo + "', '" + vContractNo + "', " + vWorkDep + ", " + vResponsibleDep + ", " +
                                                       vInvoiceNo + ", " + vCustomName + ", " + vContractText + ", " + vAssignDateStr + ", " + vEffectiveDateSStr + ", " + vEffectiveDateEStr + ", " +
                                                       vIsClose + ", " + vAmountStr + ", " + vTaxStr + ", " + vTotalAmountStr + ", " + vStampDutyStr + ", " + vRemark + ", '" + vLoginID + "', GetDate())";
                                            try
                                            {
                                                PF.ExecSQL(vConnStr, vTempStr);
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
                                            eShowError.Text = (vErrorStr != "") ? eShowError.Text + Environment.NewLine + vErrorStr : eShowError.Text;
                                            eShowError.Visible = (vErrorStr.Length > 0);
                                        }
                                    }
                                }
                            }
                        }
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 轉入路線對照
        /// </summary>
        private void Import_LinesData()
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            //string vExtName = "";
            string vTempStr = "";
            int vRC = 0;
            string vMaxIndexNoStr = "";
            string vErrorStr = "";

            string vIndexNo = "";
            string vMaxIndexNo = "";
            string vIndexNoHead = "";
            string vTicketLineNo = "";
            string vERPLineNo = "";
            string vLineGovNo = "";
            string vDepNo = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;

                    case ".xls": //EXCEL 2003 之前版本
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet vSheet_H = (HSSFSheet)wbExcel_H.GetSheet("分支路線展開");
                        if (vSheet_H.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_H.FirstRowNum + 1; vRCount <= vSheet_H.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    HSSFRow vRowTemp_H = (HSSFRow)vSheet_H.GetRow(vRCount);
                                    vTicketLineNo = vRowTemp_H.GetCell(2).ToString().Trim();
                                    vERPLineNo = (vRowTemp_H.GetCell(4).ToString().Trim() != "-") ? Int32.Parse(vRowTemp_H.GetCell(4).ToString().Trim()).ToString("D5") : vERPLineNo;
                                    vLineGovNo = vRowTemp_H.GetCell(0).ToString().Trim();

                                    vTempStr = "select count(LinesNo) RCount from Lines where LinesNo = '" + vERPLineNo + "' ";
                                    vRC = Int32.Parse(PF.GetValue(vConnStr, vTempStr, "RCount"));
                                    if (vRC == 0)
                                    {
                                        vErrorStr += (Environment.NewLine + "系統內找不到ERP路線代號 [" + vERPLineNo + "]，請重新確認！");
                                    }
                                    else if (vRC == 1)
                                    {
                                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vRowTemp_H.GetCell(10).ToString().Trim() + "' ";
                                        vDepNo = PF.GetValue(vConnStr, vTempStr, "DepNo");
                                        vTempStr = "select count(TicketLinesNo) RCount from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vRC = Int32.Parse(PF.GetValue(vConnStr, vTempStr, "RCount"));
                                        if (vRC == 0)
                                        {
                                            vIndexNoHead = DateTime.Today.ToString("yyyyMM") + "ERPLines";
                                            vTempStr = "select Max(IndexNo) MaxIndex from LinesNoChart where IndexNo like '" + vIndexNoHead + "%' ";
                                            vMaxIndexNoStr = PF.GetValue(vConnStr, vTempStr, "MaxIndex").Replace(vIndexNoHead, "");
                                            vMaxIndexNo = (vMaxIndexNoStr != "") ? (Int32.Parse(vMaxIndexNoStr) + 1).ToString("D4") : "0001";
                                            vIndexNo = vIndexNoHead.Trim() + vMaxIndexNo;
                                            vTempStr = "INSERT INTO LinesNoChart " + Environment.NewLine +
                                                       "       (IndexNo, DepNo, ERPLinesNo, GOVLinesNo, TicketLinesNo, BuMan, BuDate) " + Environment.NewLine +
                                                       "VALUES ('" + vIndexNo + "', '" + vDepNo + "', '" + vERPLineNo + "', '" + vLineGovNo + "', '" + vTicketLineNo + "', '" + vLoginID + "', GetDate())";
                                            PF.ExecSQL(vConnStr, vTempStr);
                                        }
                                        else if (vRC == 1)
                                        {
                                            vTempStr = "select IndexNo from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                            vIndexNo = PF.GetValue(vConnStr, vTempStr, "IndexNo");
                                            vTempStr = "UPDATE LinesNoChart " + Environment.NewLine +
                                                       "   SET DepNo = '" + vDepNo + "', ERPLinesNo = '" + vERPLineNo + "', GOVLinesNo = '" + vLineGovNo + "', " + Environment.NewLine +
                                                       "       TicketLinesNo = '" + vTicketLineNo + "', ModifyMan = '" + vLoginID + "', ModifyDate = GetDate() " + Environment.NewLine +
                                                       " WHERE (IndexNo = '" + vIndexNo + "' )";
                                            PF.ExecSQL(vConnStr, vTempStr);
                                        }
                                        else
                                        {

                                        }
                                        vTempStr = "update Lines " + Environment.NewLine +
                                                   "   set TicketLineNo = '" + Int32.Parse(vTicketLineNo).ToString("D6") + "', LinesGOVNo = '" + vLineGovNo + "', ApprovedDepNo = '" + vDepNo + "' " + Environment.NewLine +
                                                   " where LinesNo = '" + vERPLineNo + "' ";
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                    else
                                    {

                                    }
                                }
                                catch (Exception eMessage)
                                {
                                    //Response.Write("<Script language='Javascript'>");
                                    //Response.Write("alert('" + eMessage.Message + Environment.NewLine + vTempStr + "')");
                                    //Response.Write("</" + "Script>");
                                    lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                }
                            }
                        }
                        break;

                    case ".xlsx": //EXCEL 2010 之後版本
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet vSheet_X = (XSSFSheet)wbExcel_X.GetSheet("分支路線展開");
                        if (vSheet_X.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_X.FirstRowNum + 1; vRCount <= vSheet_X.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    XSSFRow vRowTemp_X = (XSSFRow)vSheet_X.GetRow(vRCount);
                                    vTicketLineNo = vRowTemp_X.GetCell(2).ToString().Trim();
                                    vERPLineNo = (vRowTemp_X.GetCell(4).ToString().Trim() != "-") ? Int32.Parse(vRowTemp_X.GetCell(4).ToString().Trim()).ToString("D5") : vERPLineNo;
                                    vLineGovNo = vRowTemp_X.GetCell(0).ToString().Trim();

                                    vTempStr = "select count(LinesNo) RCount from Lines where LinesNo = '" + vERPLineNo + "' ";
                                    vRC = Int32.Parse(PF.GetValue(vConnStr, vTempStr, "RCount"));
                                    if (vRC == 0)
                                    {
                                        vErrorStr += (Environment.NewLine + "系統內找不到ERP路線代號 [" + vERPLineNo + "]，請重新確認！");
                                    }
                                    else if (vRC == 1)
                                    {
                                        vTempStr = "select top 1 DepNo from Department where [Name] = '" + vRowTemp_X.GetCell(10).ToString().Trim() + "' ";
                                        vDepNo = PF.GetValue(vConnStr, vTempStr, "DepNo");
                                        vTempStr = "select count(TicketLinesNo) RCount from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vRC = Int32.Parse(PF.GetValue(vConnStr, vTempStr, "RCount"));
                                        if (vRC == 0)
                                        {
                                            vIndexNoHead = DateTime.Today.ToString("yyyyMM") + "ERPLines";
                                            vTempStr = "select Max(IndexNo) MaxIndex from LinesNoChart where IndexNo like '" + vIndexNoHead + "%' ";
                                            vMaxIndexNoStr = PF.GetValue(vConnStr, vTempStr, "MaxIndex").Replace(vIndexNoHead, "");
                                            vMaxIndexNo = (vMaxIndexNoStr != "") ? (Int32.Parse(vMaxIndexNoStr) + 1).ToString("D4") : "0001";
                                            vIndexNo = vIndexNoHead.Trim() + vMaxIndexNo;
                                            vTempStr = "INSERT INTO LinesNoChart " + Environment.NewLine +
                                                       "       (IndexNo, DepNo, ERPLinesNo, GOVLinesNo, TicketLinesNo, BuMan, BuDate) " + Environment.NewLine +
                                                       "VALUES ('" + vIndexNo + "', '" + vDepNo + "', '" + vERPLineNo + "', '" + vLineGovNo + "', '" + vTicketLineNo + "', '" + vLoginID + "', GetDate())";
                                            PF.ExecSQL(vConnStr, vTempStr);
                                        }
                                        else if (vRC == 1)
                                        {
                                            vTempStr = "select IndexNo from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                            vIndexNo = PF.GetValue(vConnStr, vTempStr, "IndexNo");
                                            vTempStr = "UPDATE LinesNoChart " + Environment.NewLine +
                                                       "   SET DepNo = '" + vDepNo + "', ERPLinesNo = '" + vERPLineNo + "', GOVLinesNo = '" + vLineGovNo + "', " + Environment.NewLine +
                                                       "       TicketLinesNo = '" + vTicketLineNo + "', ModifyMan = '" + vLoginID + "', ModifyDate = GetDate() " + Environment.NewLine +
                                                       " WHERE (IndexNo = '" + vIndexNo + "' )";
                                            PF.ExecSQL(vConnStr, vTempStr);
                                        }
                                        else
                                        {

                                        }
                                        vTempStr = "update Lines " + Environment.NewLine +
                                                   "   set TicketLineNo = '" + Int32.Parse(vTicketLineNo).ToString("D6") + "', LinesGOVNo = '" + vLineGovNo + "', ApprovedDepNo = '" + vDepNo + "' " + Environment.NewLine +
                                                   " where LinesNo = '" + vERPLineNo + "' ";
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                    else
                                    {

                                    }
                                }
                                catch (Exception eMessage)
                                {
                                    //Response.Write("<Script language='Javascript'>");
                                    //Response.Write("alert('" + eMessage.Message + Environment.NewLine + vTempStr + "')");
                                    //Response.Write("</" + "Script>");
                                    lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                }
                            }
                        }
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 更新不休假獎金資料
        /// </summary>
        private void Import_UnusedBounds()
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            //string vExtName = "";
            string vTempStr = "";

            string vCalYM = "";
            string vEmpNo = "";
            int vTotalPay = 0;
            string vTotalPayStr = "";
            int vSPDays = 0;
            string vSPDaysStr = "";
            int vUsedDay = 0;
            string vUsedDayStr = "";
            int vUnusedDay = 0;
            string vUnusedDayStr = "";
            int vUnusedBounds = 0;
            string vUnusedBoundsStr = "";


            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;

                    case ".xls"://EXCEL 2003 之前版本
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet vSheet_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        if (vSheet_H.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_H.FirstRowNum + 1; vRCount <= vSheet_H.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    HSSFRow vRowTemp_H = (HSSFRow)vSheet_H.GetRow(vRCount);
                                    vCalYM = vRowTemp_H.GetCell(0).ToString().Trim();
                                    vEmpNo = vRowTemp_H.GetCell(1).ToString().Trim();
                                    vTotalPay = (vRowTemp_H.GetCell(2).ToString().Trim() != "") ? (int)vRowTemp_H.GetCell(2).NumericCellValue : 0;
                                    vTotalPayStr = vTotalPay.ToString();
                                    vSPDays = (vRowTemp_H.GetCell(4).ToString().Trim() != "") ? (int)vRowTemp_H.GetCell(4).NumericCellValue : 0;
                                    vSPDaysStr = vSPDays.ToString();
                                    vUsedDay = (vRowTemp_H.GetCell(5).ToString().Trim() != "") ? (int)vRowTemp_H.GetCell(5).NumericCellValue : 0;
                                    vUsedDayStr = vUsedDay.ToString();
                                    vUnusedDay = (vRowTemp_H.GetCell(6).ToString().Trim() != "") ? (int)vRowTemp_H.GetCell(6).NumericCellValue : 0;
                                    vUnusedDayStr = vUnusedDay.ToString();
                                    vUnusedBounds = (vRowTemp_H.GetCell(7).ToString().Trim() != "") ? (int)vRowTemp_H.GetCell(7).NumericCellValue : 0;
                                    vUnusedBoundsStr = vUnusedBounds.ToString();
                                    vTempStr = "update YearHoliday set Holibounds = " + vTotalPayStr + ", Spldays = " + vSPDaysStr + ", UsedDays = " + vUsedDayStr + Environment.NewLine +
                                               "       Predays = " + vUnusedDayStr + ", UnusedBounds = " + vUnusedBoundsStr + Environment.NewLine +
                                               " where EmpNo = '" + vEmpNo + "' " + Environment.NewLine +
                                               "   and Years = " + vCalYM;
                                    PF.ExecSQL(vConnStr, vTempStr);
                                }
                                catch (Exception eMessage)
                                {
                                    //Response.Write("<Script language='Javascript'>");
                                    //Response.Write("alert('" + eMessage.Message + Environment.NewLine + vTempStr + "')");
                                    //Response.Write("</" + "Script>");
                                    lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                }
                            }
                        }
                        break;

                    case ".xlsx"://EXCEL 2010 之後版本
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet vSheet_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        if (vSheet_X.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_X.FirstRowNum + 1; vRCount <= vSheet_X.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    XSSFRow vRowTemp_X = (XSSFRow)vSheet_X.GetRow(vRCount);
                                    vCalYM = vRowTemp_X.GetCell(0).ToString().Trim();
                                    vEmpNo = vRowTemp_X.GetCell(1).ToString().Trim();
                                    vTotalPay = (vRowTemp_X.GetCell(2).ToString().Trim() != "") ? (int)vRowTemp_X.GetCell(2).NumericCellValue : 0;
                                    vTotalPayStr = vTotalPay.ToString();
                                    vSPDays = (vRowTemp_X.GetCell(4).ToString().Trim() != "") ? (int)vRowTemp_X.GetCell(4).NumericCellValue : 0;
                                    vSPDaysStr = vSPDays.ToString();
                                    vUsedDay = (vRowTemp_X.GetCell(5).ToString().Trim() != "") ? (int)vRowTemp_X.GetCell(5).NumericCellValue : 0;
                                    vUsedDayStr = vUsedDay.ToString();
                                    vUnusedDay = (vRowTemp_X.GetCell(6).ToString().Trim() != "") ? (int)vRowTemp_X.GetCell(6).NumericCellValue : 0;
                                    vUnusedDayStr = vUnusedDay.ToString();
                                    vUnusedBounds = (vRowTemp_X.GetCell(7).ToString().Trim() != "") ? (int)vRowTemp_X.GetCell(7).NumericCellValue : 0;
                                    vUnusedBoundsStr = vUnusedBounds.ToString();
                                    vTempStr = "update YearHoliday set Holibounds = " + vTotalPayStr + ", Spldays = " + vSPDaysStr + ", UsedDays = " + vUsedDayStr + Environment.NewLine +
                                               "       Predays = " + vUnusedDayStr + ", UnusedBounds = " + vUnusedBoundsStr + Environment.NewLine +
                                               " where EmpNo = '" + vEmpNo + "' " + Environment.NewLine +
                                               "   and Years = " + vCalYM;
                                    PF.ExecSQL(vConnStr, vTempStr);
                                }
                                catch (Exception eMessage)
                                {
                                    //Response.Write("<Script language='Javascript'>");
                                    //Response.Write("alert('" + eMessage.Message + Environment.NewLine + vTempStr + "')");
                                    //Response.Write("</" + "Script>");
                                    lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                }
                            }
                        }
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 匯入MOU津貼資料
        /// </summary>
        private void Import_MOUPayData()
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            int vTempINT = 0;
            string vTempStr = "";
            string vTempNo = "";
            string vRCountStr = "";

            string vCalYM = "";
            string vEmpNo = "";
            string vPayNo = "";
            string vPayBNo = "";
            double vExpense = 0.0;
            string vPriNo = "";
            string vPayDate = "";
            string vRealPayDate = "";
            string vPayDur = "2";
            DateTime vAssumeday;
            double vDurNum = 0;
            int vTempINT_1 = 0;
            int vTempINT_2 = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;

                    case ".xls"://EXCEL 2003 之前版本
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet vSheet_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        if (vSheet_H.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_H.FirstRowNum + 1; vRCount <= vSheet_H.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    HSSFRow vRowTemp_H = (HSSFRow)vSheet_H.GetRow(vRCount);
                                    vEmpNo = vRowTemp_H.GetCell(0).StringCellValue.Trim();
                                    vPayNo = vRowTemp_H.GetCell(2).StringCellValue.Trim();
                                    vPayBNo = vRowTemp_H.GetCell(3).StringCellValue.Trim();
                                    vExpense = vRowTemp_H.GetCell(4).NumericCellValue;
                                    vTempStr = "select Assumeday from Employee where EmpNo = '" + vEmpNo + "' ";
                                    vAssumeday = DateTime.Parse(PF.GetValue(vConnStr, vTempStr, "Assumeday"));
                                    vTempINT_1 = PF.GetMonthDays(vAssumeday);
                                    vTempINT_2 = vAssumeday.Day;
                                    vDurNum = vTempINT_1 - vTempINT_2 + 1;
                                    vPayDate = PF.GetMonthFirstDay(vAssumeday, "B");
                                    vRealPayDate = PF.GetMonthFirstDay(vAssumeday.AddMonths(1), "B");
                                    vTempStr = "select PriNo from MSHZ where PayDate = '" + vPayDate + "' and EmpNo = '" + vEmpNo + "' and PayDur = '2' and PayBNo = '" + vPayBNo + "' ";
                                    vTempNo = PF.GetValue(vConnStr, vTempStr, "PriNo");
                                    if (vTempNo != "")
                                    {
                                        vPriNo = vTempNo.Trim();
                                        vTempStr = "update MSHZ set Expense = Expense + " + vExpense + " where PriNo = '" + vPriNo + "' ";
                                    }
                                    else
                                    {
                                        vCalYM = vAssumeday.ToString("yyyyMM");
                                        vPriNo = PF.GetDataIndex(vConnStr, "MSHZ", "PriNo", "B", false, vAssumeday, 6, "", 4);
                                        vTempStr = "insert into MSHZ(PriNo,PayDate,PinDate,PayNo,PayDur,EmpNo,Company,DepNo,Title,Type,Expense,PayBNo)" + Environment.NewLine +
                                                   "          values('" + vPriNo + "', '" + vPayDate + "', '" + vPayDate + "', '" + vPayNo + "', '" + vPayDur + "', '" + vEmpNo + "', 'A000', '06', '271', '00', " + vExpense.ToString() + ", '" + vPayBNo + "')";
                                    }
                                    PF.ExecSQL(vConnStr, vTempStr);
                                    vTempStr = "select count(EmpNo) RCount from PayRec where PayDate = '" + vPayDate + "' and PayDur = '2' and EmpNo = '" + vEmpNo + "' ";
                                    vRCountStr = PF.GetValue(vConnStr, vTempStr, "RCount");
                                    if ((Int32.TryParse(vRCountStr, out vTempINT)) && (vTempINT != 0))
                                    {
                                        vTempStr = "update PayRec set CashNum14 = CashNum14 + " + vExpense + ", Tracash1 = Tracash1 + " + vExpense + ", RealCash = RealCash + " + vExpense + ", GIVCash = GIVCash + " + vExpense + Environment.NewLine +
                                                   " where PayDate = '" + vPayDate + "' and PayDur = '2' and EmpNo = '" + vEmpNo + "' ";
                                    }
                                    else
                                    {
                                        vTempStr = "INSERT INTO PAYREC ( " + Environment.NewLine +
                                                   "       paydate, paydur, empno, name, company, depno, servno, groupno1, title, type, unit, paytype, bankno1, bankno2, account1, account2, supplynum, " + Environment.NewLine +
                                                   "       hikind, hiamt, laiamt, hinum, gikind, patnum, durnum, attnum, holnum, trunum, esctype1, esctype2, esctype3, esctype4, esctype5, esctype6, esctype7, " + Environment.NewLine +
                                                   "       esctype8, esctype9, esctype10, esctype11, esctype12, frontover, postover, superover, holover, dechour, exthour, patchour, latetime, latemin, forgetime, " + Environment.NewLine +
                                                   "       chorenum, remark, nowpay1, nowpay2, nowpay3, nowpay4, nowpay5, nowpay6, nowpay7, nowpay8, cashno01, cashno02, cashno03, cashno04, cashno05, cashno06, " + Environment.NewLine +
                                                   "       cashno07, cashno08, cashno09, cashno10, cashno11, cashno12, cashno13, cashno14, cashno15, cashno16, cashno17, cashno18, cashno19, cashno20, cashno21, " + Environment.NewLine +
                                                   "       cashno22, cashno23, cashno24, cashno25, cashno26, cashno27, cashno28, cashno29, cashno30, cashnum01, cashnum02, cashnum03, cashnum04, cashnum05, cashnum06, " + Environment.NewLine +
                                                   "       cashnum07, cashnum08, cashnum09, cashnum10, cashnum11, cashnum12, cashnum13, cashnum14, cashnum15, cashnum16, cashnum17, cashnum18, cashnum19, cashnum20, " + Environment.NewLine +
                                                   "       cashnum21, cashnum22, cashnum23, cashnum24, cashnum25, cashnum26, cashnum27, cashnum28, cashnum29, cashnum30, cashno41, cashno42, cashno43, cashno44, cashno45, " + Environment.NewLine +
                                                   "       cashno46, cashno47, cashno48, cashno49, cashno50, cashno51, cashno52, cashno53, cashno54, cashno55, cashno56, cashno57, cashno58, cashno59, cashno60, cashno61, " + Environment.NewLine +
                                                   "       cashno62, cashno63, cashno64, cashno65, cashno66, cashno67, cashno68, cashno69, cashno70, cashnum41, cashnum42, cashnum43, cashnum44, cashnum45, cashnum46, cashnum47, " + Environment.NewLine +
                                                   "       cashnum48, cashnum49, cashnum50, cashnum51, cashnum52, cashnum53, cashnum54, cashnum55, cashnum56, cashnum57, cashnum58, cashnum59, cashnum60, cashnum61, cashnum62, " + Environment.NewLine +
                                                   "       cashnum63, cashnum64, cashnum65, cashnum66, cashnum67, cashnum68, cashnum69, cashnum70, cashno71, cashno72, cashno73, cashno74, cashno75, cashno76, cashno77, cashno78, " + Environment.NewLine +
                                                   "       cashno79, cashnum71, cashnum72, cashnum73, cashnum74, cashnum75, cashnum76, cashnum77, cashnum78, cashnum79, fixtax, vartax, incotax, lifee, hifee, taxpay, notaxpay, " + Environment.NewLine +
                                                   "       bonuspay, givcash, nogivcash, realcash, cashpay, tracash1, tracash2, nholnum, repast, overfit, detain, gifee, Payrolldate, ACKIND, REALPAYDATE, FRONTOVER2, POSTOVER2, POSTOVER22)" + Environment.NewLine +
                                                   "select '" + vPayDate + "', '2', EmpNo, [Name], Company, '06', isnull(ServNo, ''), isnull(GroupNo1, ''), Title, Type, Unit, PayType, " + Environment.NewLine +
                                                   "       isnull(BankNo1, ''), isnull(BankNo2, ''), isnull(Account1, ''), isnull(Account2, ''), SupplyNum, HiKind, HiAmt, LaiAmt, HiNum, GiKind, 0, " + Environment.NewLine +
                                                   "       " + vDurNum.ToString().Trim() + ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, 0, 0, 0, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                   "       '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', " + Environment.NewLine +
                                                   "       0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " + vExpense + ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                   "       '41', '42', '43', '44', '45', '46', '47', '48', '49', '50', '51', '52', '53', '54', '55', '56', '57', '58', '59', '60', '61', '62', '63', '64', '65', '66', '67', '68', '69', '70', " + Environment.NewLine +
                                                   "       0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                   "       '71', '72', '73', '74', '75', '76', '77', '78', '79', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " + vExpense + ", " + Environment.NewLine +
                                                   "       0, " + vExpense + ", 0, " + vExpense + ", 0, " + vExpense + ", 0, 0, NULL, NULL, NULL, NULL, " + vPayDate + ", '1', " + vRealPayDate + ", 0, 0, 0 " + Environment.NewLine +
                                                   "  from Employee " + Environment.NewLine +
                                                   " where EmpNo = '" + vEmpNo + "' ";
                                    }
                                    PF.ExecSQL(vConnStr, vTempStr);
                                }
                                catch (Exception eMessage)
                                {
                                    //Response.Write("<Script language='Javascript'>");
                                    //Response.Write("alert('" + eMessage.Message + Environment.NewLine + vTempStr + "')");
                                    //Response.Write("</" + "Script>");
                                    lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                }
                            }
                        }
                        break;

                    case ".xlsx"://EXCEL 2010 之後版本
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet vSheet_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        if (vSheet_X.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_X.FirstRowNum + 1; vRCount <= vSheet_X.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    XSSFRow vRowTemp_X = (XSSFRow)vSheet_X.GetRow(vRCount);
                                    vEmpNo = vRowTemp_X.GetCell(0).StringCellValue.Trim();
                                    vPayNo = vRowTemp_X.GetCell(2).StringCellValue.Trim();
                                    vPayBNo = vRowTemp_X.GetCell(3).StringCellValue.Trim();
                                    vExpense = vRowTemp_X.GetCell(4).NumericCellValue;
                                    vTempStr = "select Assumeday from Employee where EmpNo = '" + vEmpNo + "' ";
                                    vAssumeday = DateTime.Parse(PF.GetValue(vConnStr, vTempStr, "Assumeday"));
                                    vTempINT_1 = PF.GetMonthDays(vAssumeday);
                                    vTempINT_2 = vAssumeday.Day;
                                    vDurNum = vTempINT_1 - vTempINT_2 + 1;
                                    vPayDate = PF.GetMonthFirstDay(vAssumeday, "B");
                                    vRealPayDate = PF.GetMonthFirstDay(vAssumeday.AddMonths(1), "B");
                                    vTempStr = "select PriNo from MSHZ where PayDate = '" + vPayDate + "' and EmpNo = '" + vEmpNo + "' and PayDur = '2' and PayBNo = '" + vPayBNo + "' ";
                                    vTempNo = PF.GetValue(vConnStr, vTempStr, "PriNo");
                                    if (vTempNo != "")
                                    {
                                        vPriNo = vTempNo.Trim();
                                        vTempStr = "update MSHZ set Expense = Expense + " + vExpense + " where PriNo = '" + vPriNo + "' ";
                                    }
                                    else
                                    {
                                        vCalYM = vAssumeday.ToString("yyyyMM");
                                        vPriNo = PF.GetDataIndex(vConnStr, "MSHZ", "PriNo", "B", false, vAssumeday, 6, "", 4);
                                        vTempStr = "insert into MSHZ(PriNo,PayDate,PinDate,PayNo,PayDur,EmpNo,Company,DepNo,Title,Type,Expense,PayBNo)" + Environment.NewLine +
                                                   "          values('" + vPriNo + "', '" + vPayDate + "', '" + vPayDate + "', '" + vPayNo + "', '" + vPayDur + "', '" + vEmpNo + "', 'A000', '06', '271', '00', " + vExpense.ToString() + ", '" + vPayBNo + "')";
                                    }
                                    PF.ExecSQL(vConnStr, vTempStr);
                                    vTempStr = "select count(EmpNo) RCount from PayRec where PayDate = '" + vPayDate + "' and PayDur = '2' and EmpNo = '" + vEmpNo + "' ";
                                    vRCountStr = PF.GetValue(vConnStr, vTempStr, "RCount");
                                    if ((Int32.TryParse(vRCountStr, out vTempINT)) && (vTempINT != 0))
                                    {
                                        vTempStr = "update PayRec set CashNum14 = CashNum14 + " + vExpense + ", Tracash1 = Tracash1 + " + vExpense + ", RealCash = RealCash + " + vExpense + ", GIVCash = GIVCash + " + vExpense + Environment.NewLine +
                                                   " where PayDate = '" + vPayDate + "' and PayDur = '2' and EmpNo = '" + vEmpNo + "' ";
                                    }
                                    else
                                    {
                                        vTempStr = "INSERT INTO PAYREC ( " + Environment.NewLine +
                                                   "       paydate, paydur, empno, name, company, depno, servno, groupno1, title, type, unit, paytype, bankno1, bankno2, account1, account2, supplynum, " + Environment.NewLine +
                                                   "       hikind, hiamt, laiamt, hinum, gikind, patnum, durnum, attnum, holnum, trunum, esctype1, esctype2, esctype3, esctype4, esctype5, esctype6, esctype7, " + Environment.NewLine +
                                                   "       esctype8, esctype9, esctype10, esctype11, esctype12, frontover, postover, superover, holover, dechour, exthour, patchour, latetime, latemin, forgetime, " + Environment.NewLine +
                                                   "       chorenum, remark, nowpay1, nowpay2, nowpay3, nowpay4, nowpay5, nowpay6, nowpay7, nowpay8, cashno01, cashno02, cashno03, cashno04, cashno05, cashno06, " + Environment.NewLine +
                                                   "       cashno07, cashno08, cashno09, cashno10, cashno11, cashno12, cashno13, cashno14, cashno15, cashno16, cashno17, cashno18, cashno19, cashno20, cashno21, " + Environment.NewLine +
                                                   "       cashno22, cashno23, cashno24, cashno25, cashno26, cashno27, cashno28, cashno29, cashno30, cashnum01, cashnum02, cashnum03, cashnum04, cashnum05, cashnum06, " + Environment.NewLine +
                                                   "       cashnum07, cashnum08, cashnum09, cashnum10, cashnum11, cashnum12, cashnum13, cashnum14, cashnum15, cashnum16, cashnum17, cashnum18, cashnum19, cashnum20, " + Environment.NewLine +
                                                   "       cashnum21, cashnum22, cashnum23, cashnum24, cashnum25, cashnum26, cashnum27, cashnum28, cashnum29, cashnum30, cashno41, cashno42, cashno43, cashno44, cashno45, " + Environment.NewLine +
                                                   "       cashno46, cashno47, cashno48, cashno49, cashno50, cashno51, cashno52, cashno53, cashno54, cashno55, cashno56, cashno57, cashno58, cashno59, cashno60, cashno61, " + Environment.NewLine +
                                                   "       cashno62, cashno63, cashno64, cashno65, cashno66, cashno67, cashno68, cashno69, cashno70, cashnum41, cashnum42, cashnum43, cashnum44, cashnum45, cashnum46, cashnum47, " + Environment.NewLine +
                                                   "       cashnum48, cashnum49, cashnum50, cashnum51, cashnum52, cashnum53, cashnum54, cashnum55, cashnum56, cashnum57, cashnum58, cashnum59, cashnum60, cashnum61, cashnum62, " + Environment.NewLine +
                                                   "       cashnum63, cashnum64, cashnum65, cashnum66, cashnum67, cashnum68, cashnum69, cashnum70, cashno71, cashno72, cashno73, cashno74, cashno75, cashno76, cashno77, cashno78, " + Environment.NewLine +
                                                   "       cashno79, cashnum71, cashnum72, cashnum73, cashnum74, cashnum75, cashnum76, cashnum77, cashnum78, cashnum79, fixtax, vartax, incotax, lifee, hifee, taxpay, notaxpay, " + Environment.NewLine +
                                                   "       bonuspay, givcash, nogivcash, realcash, cashpay, tracash1, tracash2, nholnum, repast, overfit, detain, gifee, Payrolldate, ACKIND, REALPAYDATE, FRONTOVER2, POSTOVER2, POSTOVER22)" + Environment.NewLine +
                                                   "select '" + vPayDate + "', '2', EmpNo, [Name], Company, '06', isnull(ServNo, ''), isnull(GroupNo1, ''), Title, Type, Unit, PayType, " + Environment.NewLine +
                                                   "       isnull(BankNo1, ''), isnull(BankNo2, ''), isnull(Account1, ''), isnull(Account2, ''), SupplyNum, HiKind, HiAmt, LaiAmt, HiNum, GiKind, 0, " + Environment.NewLine +
                                                   "       " + vDurNum.ToString().Trim() + ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, 0, 0, 0, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                   "       '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', " + Environment.NewLine +
                                                   "       0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " + vExpense + ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                   "       '41', '42', '43', '44', '45', '46', '47', '48', '49', '50', '51', '52', '53', '54', '55', '56', '57', '58', '59', '60', '61', '62', '63', '64', '65', '66', '67', '68', '69', '70', " + Environment.NewLine +
                                                   "       0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                   "       '71', '72', '73', '74', '75', '76', '77', '78', '79', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " + vExpense + ", " + Environment.NewLine +
                                                   "       0, " + vExpense + ", 0, " + vExpense + ", 0, " + vExpense + ", 0, 0, NULL, NULL, NULL, NULL, " + vPayDate + ", '1', " + vRealPayDate + ", 0, 0, 0 " + Environment.NewLine +
                                                   "  from Employee " + Environment.NewLine +
                                                   " where EmpNo = '" + vEmpNo + "' ";
                                    }
                                    PF.ExecSQL(vConnStr, vTempStr);
                                }
                                catch (Exception eMessage)
                                {
                                    //Response.Write("<Script language='Javascript'>");
                                    //Response.Write("alert('" + eMessage.Message + Environment.NewLine + vTempStr + "')");
                                    //Response.Write("</" + "Script>");
                                    lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                }
                            }
                        }
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        private void Modify_MOUPayData()
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            int vTempINT = 0;
            string vTempStr = "";
            string vTempNo = "";
            string vRCountStr = "";

            //string vCalYM = "";
            string vEmpNo = "";
            string vPayNo = "";
            string vPayBNo = "";
            double vExpense = 0.0;
            string vExpenseStr = "";
            string vPriNo = "";
            string vPayDate = "";
            string vRealPayDate = "";
            //string vPayDur = "2";
            DateTime vAssumeday;
            double vDurNum = 0;
            int vTempINT_1 = 0;
            int vTempINT_2 = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;

                    case ".xls"://EXCEL 2003 之前版本
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet vSheet_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        if (vSheet_H.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_H.FirstRowNum + 1; vRCount <= vSheet_H.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    HSSFRow vRowTemp_H = (HSSFRow)vSheet_H.GetRow(vRCount);
                                    vEmpNo = vRowTemp_H.GetCell(0).StringCellValue.Trim();
                                    vPayNo = vRowTemp_H.GetCell(2).StringCellValue.Trim();
                                    vPayBNo = vRowTemp_H.GetCell(3).StringCellValue.Trim();
                                    vExpense = vRowTemp_H.GetCell(4).NumericCellValue;
                                    vExpenseStr = vExpense.ToString();
                                    vTempStr = "select Assumeday from Employee where EmpNo = '" + vEmpNo + "' ";
                                    vAssumeday = DateTime.Parse(PF.GetValue(vConnStr, vTempStr, "Assumeday"));
                                    vTempINT_1 = PF.GetMonthDays(vAssumeday);
                                    vTempINT_2 = vAssumeday.Day;
                                    vDurNum = vTempINT_1 - vTempINT_2 + 1;
                                    vPayDate = PF.GetMonthFirstDay(vAssumeday, "B");
                                    vRealPayDate = PF.GetMonthFirstDay(vAssumeday.AddMonths(1), "B");
                                    vTempStr = "select PriNo from MSHZ where PayDate = '" + vPayDate + "' and EmpNo = '" + vEmpNo + "' and PayDur = '2' and PayBNo = '" + vPayBNo + "' and Expense >= " + vExpenseStr;
                                    vTempNo = PF.GetValue(vConnStr, vTempStr, "PriNo");
                                    if (vTempNo != "")
                                    {
                                        vPriNo = vTempNo.Trim();
                                        vTempStr = "update MSHZ set Expense = Expense - " + vExpenseStr + " where PriNo = '" + vPriNo + "' ";
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                    vTempStr = "select count(EmpNo) RCount from PayRec where PayDate = '" + vPayDate + "' and PayDur = '2' and EmpNo = '" + vEmpNo + "' and CashNum14 >= " + vExpenseStr;
                                    vRCountStr = PF.GetValue(vConnStr, vTempStr, "RCount");
                                    if ((Int32.TryParse(vRCountStr, out vTempINT)) && (vTempINT != 0))
                                    {
                                        vTempStr = "update PayRec set CashNum14 = CashNum14 - " + vExpenseStr + ", Tracash1 = Tracash1 - " + vExpenseStr + ", RealCash = RealCash - " + vExpenseStr + ", GIVCash = GIVCash - " + vExpense + Environment.NewLine +
                                                   " where PayDate = '" + vPayDate + "' and PayDur = '2' and EmpNo = '" + vEmpNo + "' ";
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                }
                                catch (Exception eMessage)
                                {
                                    lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                }
                            }
                            vTempStr = "delete MSHZ where Expense = 0";
                            PF.ExecSQL(vConnStr, vTempStr);
                            vTempStr = "delete Payrec where RealCash = 0 and GIVCash = 0 and PayDur = '2'";
                            PF.ExecSQL(vConnStr, vTempStr);
                        }
                        break;

                    case ".xlsx"://EXCEL 2010 之後版本
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet vSheet_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        if (vSheet_X.LastRowNum > 0)
                        {
                            for (int vRCount = vSheet_X.FirstRowNum + 1; vRCount <= vSheet_X.LastRowNum; vRCount++)
                            {
                                try
                                {
                                    XSSFRow vRowTemp_X = (XSSFRow)vSheet_X.GetRow(vRCount);
                                    vEmpNo = vRowTemp_X.GetCell(0).StringCellValue.Trim();
                                    vPayNo = vRowTemp_X.GetCell(2).StringCellValue.Trim();
                                    vPayBNo = vRowTemp_X.GetCell(3).StringCellValue.Trim();
                                    vExpense = vRowTemp_X.GetCell(4).NumericCellValue;
                                    vExpenseStr = vExpense.ToString();
                                    vTempStr = "select Assumeday from Employee where EmpNo = '" + vEmpNo + "' ";
                                    vAssumeday = DateTime.Parse(PF.GetValue(vConnStr, vTempStr, "Assumeday"));
                                    vTempINT_1 = PF.GetMonthDays(vAssumeday);
                                    vTempINT_2 = vAssumeday.Day;
                                    vDurNum = vTempINT_1 - vTempINT_2 + 1;
                                    vPayDate = PF.GetMonthFirstDay(vAssumeday, "B");
                                    vRealPayDate = PF.GetMonthFirstDay(vAssumeday.AddMonths(1), "B");
                                    vTempStr = "select PriNo from MSHZ where PayDate = '" + vPayDate + "' and EmpNo = '" + vEmpNo + "' and PayDur = '2' and PayBNo = '" + vPayBNo + "' and Expense >= " + vExpenseStr;
                                    vTempNo = PF.GetValue(vConnStr, vTempStr, "PriNo");
                                    if (vTempNo != "")
                                    {
                                        vPriNo = vTempNo.Trim();
                                        vTempStr = "update MSHZ set Expense = Expense - " + vExpenseStr + " where PriNo = '" + vPriNo + "' ";
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                    vTempStr = "select count(EmpNo) RCount from PayRec where PayDate = '" + vPayDate + "' and PayDur = '2' and EmpNo = '" + vEmpNo + "' and CashNum14 >= " + vExpenseStr;
                                    vRCountStr = PF.GetValue(vConnStr, vTempStr, "RCount");
                                    if ((Int32.TryParse(vRCountStr, out vTempINT)) && (vTempINT != 0))
                                    {
                                        vTempStr = "update PayRec set CashNum14 = CashNum14 - " + vExpenseStr + ", Tracash1 = Tracash1 - " + vExpenseStr + ", RealCash = RealCash - " + vExpenseStr + ", GIVCash = GIVCash - " + vExpense + Environment.NewLine +
                                                   " where PayDate = '" + vPayDate + "' and PayDur = '2' and EmpNo = '" + vEmpNo + "' ";
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                }
                                catch (Exception eMessage)
                                {
                                    lbError.Text = eMessage.Message + Environment.NewLine + vTempStr;
                                }
                            }
                            vTempStr = "delete MSHZ where Expense = 0";
                            PF.ExecSQL(vConnStr, vTempStr);
                            vTempStr = "delete Payrec where RealCash = 0 and GIVCash = 0 and PayDur = '2'";
                            PF.ExecSQL(vConnStr, vTempStr);
                        }
                        break;    
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先指定要匯入的檔案。')");
                Response.Write("</" + "Script>");
            }
        }

        protected void ddlTargetTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (ddlTargetTable.SelectedValue)
            {
                case "T005":
                    lbCalYM.Enabled = true;
                    eCalYear.Enabled = true;
                    lbCalYear.Enabled = true;
                    eCalMonth.Enabled = true;
                    lbCalMonth.Enabled = true;
                    break;
                default:
                    lbCalYM.Enabled = false;
                    eCalYear.Enabled = false;
                    lbCalYear.Enabled = false;
                    eCalMonth.Enabled = false;
                    lbCalMonth.Enabled = false;
                    break;
            }
        }
    }
}