using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
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
    public partial class TicketCheckP : System.Web.UI.Page
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

                    //事件日期
                    string vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_S_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        plSearch.Visible = true;
                        plPrint.Visible = false;
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

        /// <summary>
        /// 取得 SQL 查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            DateTime vCaseDate_S_Search = (eCaseDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_S_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_S_Temp = vCaseDate_S_Search.Year.ToString() + "/" + vCaseDate_S_Search.ToString("MM/dd");
            DateTime vCaseDate_E_Serach = (eCaseDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_E_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_E_Temp = vCaseDate_E_Serach.Year.ToString() + "/" + vCaseDate_E_Serach.ToString("MM/dd");
            string vWStr_CaseDate = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate between '" + vCaseDate_S_Temp + "' and '" + vCaseDate_E_Temp + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? "   and t.CaseDate = '" + vCaseDate_S_Temp + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate = '" + vCaseDate_E_Temp + "' " + Environment.NewLine : "";
            string vWStr_Inspector_Temp = (eInspector_Search.Text.Trim() != "") ? "   and t.Inspector = '" + eInspector_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver_Temp = (eDriver_Search.Text.Trim() != "") ? "   and t.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CarID_Temp = (eCarID_Search.Text.Trim() != "") ? "   and t.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_LinesNo = (eLinesNo_Search.Text.Trim() != "") ? "   and t.LinesNo like '" + eLinesNo_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vReturnStr = "SELECT CaseNo, LinesNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = t.DepNo)) AS DepName, " + Environment.NewLine +
                                "       Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, " + Environment.NewLine +
                                "       ArrivePosition, ArriveTime, Car_ID, CheckNote, CaseNote, Remark, IsPassed, RP_ReportNo, " + Environment.NewLine +
                                "       (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = t.Inspector)) AS InspectorName, HasAssetNo, HasSecurityLabel, " + Environment.NewLine +
                                "       StationLight, case when StationLight = 'X' then '異常' else '正常' end StationLight_C, " + Environment.NewLine +
                                "       StationToilet, " + Environment.NewLine +
                                "       case when isnull(StationToilet, '') = '' then '正常' " + Environment.NewLine +
                                "            else (select ClassTxt from DBDICB where FKey = '查票工作報告   TicketCheckP    StationToilet' and ClassNo = t.StationToilet) " + Environment.NewLine +
                                "       end StationToilet_C " + Environment.NewLine +
                                "  FROM TicketCheckReport AS t " + Environment.NewLine +
                                " WHERE (1 = 1)" + Environment.NewLine +
                                vWStr_CarID_Temp + vWStr_CaseDate + vWStr_Driver_Temp + vWStr_Inspector_Temp + vWStr_LinesNo +
                                " order by CaseNo DESC ";
            return vReturnStr;
        }

        /// <summary>
        /// 取得列印用的查詢字串
        /// </summary>
        /// <param name="fCaseNo"></param>
        /// <returns></returns>
        private string GetSelectStr_Print(string fCaseNo)
        {
            string vReturnStr = "";
            if (fCaseNo != "")
            {
                vReturnStr = "SELECT CaseNo, LinesNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = t.DepNo)) AS DepName, " + Environment.NewLine +
                             "       Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, " + Environment.NewLine +
                             "       ArrivePosition, ArriveTime, Car_ID, CheckNote, CaseNote, Remark, " + Environment.NewLine +
                             "       BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = t.BuMan)) AS BuManName, " + Environment.NewLine +
                             "       ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = t.ModifyMan)) AS ModifyManName, " + Environment.NewLine +
                             "       IsPassed, RP_ReportNo, Inspector, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = t.Inspector)) AS InspectorName, " + Environment.NewLine +
                             "       HasAssetNo, HasSecurityLabel, " + Environment.NewLine +
                             "       StationLight, case when StationLight = 'X' then '異常' when StationLight = 'N' then '正常' else '' end StationLight_C, " + Environment.NewLine +
                             "       StationToilet, " + Environment.NewLine +
                             "       case when isnull(StationToilet, '') = '' then '正常' " + Environment.NewLine +
                             "            else (select ClassTxt from DBDICB where FKey = '查票工作報告    TicketCheckP    StationToilet' and ClassNo = t.StationToilet) " + Environment.NewLine +
                             "       end StationToilet_C " + Environment.NewLine +
                             "  FROM TicketCheckReport AS t " + Environment.NewLine +
                             " WHERE (CaseNo = '" + fCaseNo + "')";
            }
            else
            {
                DateTime vCaseDate_S_Search = (eCaseDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_S_Search.Text.Trim()) : DateTime.Today;
                string vCaseDate_S_Temp = vCaseDate_S_Search.Year.ToString() + "/" + vCaseDate_S_Search.ToString("MM/dd");
                DateTime vCaseDate_E_Serach = (eCaseDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_E_Search.Text.Trim()) : DateTime.Today;
                string vCaseDate_E_Temp = vCaseDate_E_Serach.Year.ToString() + "/" + vCaseDate_E_Serach.ToString("MM/dd");
                string vWStr_CaseDate = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate between '" + vCaseDate_S_Temp + "' and '" + vCaseDate_E_Temp + "' " + Environment.NewLine :
                                        ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? "   and t.CaseDate = '" + vCaseDate_S_Temp + "' " + Environment.NewLine :
                                        ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate = '" + vCaseDate_E_Temp + "' " + Environment.NewLine : "";
                string vWStr_Inspector_Temp = (eInspector_Search.Text.Trim() != "") ? "   and t.Inspector = '" + eInspector_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_Driver_Temp = (eDriver_Search.Text.Trim() != "") ? "   and t.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_CarID_Temp = (eCarID_Search.Text.Trim() != "") ? "   and t.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_LinesNo = (eLinesNo_Search.Text.Trim() != "") ? "   and t.LinesNo like '" + eLinesNo_Search.Text.Trim() + "%' " + Environment.NewLine : "";
                vReturnStr = "SELECT CaseNo, LinesNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = t.DepNo)) AS DepName, " + Environment.NewLine +
                             "       Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, " + Environment.NewLine +
                             "       ArrivePosition, ArriveTime, Car_ID, CheckNote, CaseNote, Remark, " + Environment.NewLine +
                             "       BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = t.BuMan)) AS BuManName, " + Environment.NewLine +
                             "       ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = t.ModifyMan)) AS ModifyManName, " + Environment.NewLine +
                             "       IsPassed, RP_ReportNo, Inspector, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = t.Inspector)) AS InspectorName, HasAssetNo, HasSecurityLabel, " + Environment.NewLine +
                             "       StationLight, case when StationLight = 'X' then '異常' when StationLight = 'N' then '正常' else '' end StationLight_C, " + Environment.NewLine +
                             "       StationToilet, " + Environment.NewLine +
                             "       case when isnull(StationToilet, '') = '' then '正常' " + Environment.NewLine +
                             "            else (select ClassTxt from DBDICB where FKey = '查票工作報告    TicketCheckP    StationToilet' and ClassNo = t.StationToilet) " + Environment.NewLine +
                             "       end StationToilet_C " + Environment.NewLine +
                             "  FROM TicketCheckReport AS t " + Environment.NewLine +
                             " WHERE (1 = 1)" + Environment.NewLine +
                             vWStr_CarID_Temp + vWStr_CaseDate + vWStr_Driver_Temp + vWStr_Inspector_Temp + vWStr_LinesNo +
                             " order by CaseNo DESC ";
            }
            return vReturnStr;
        }

        /// <summary>
        /// 開啟資料庫
        /// </summary>
        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            sdsTicketCheckReportList.SelectCommand = vSelStr;
            gridTicketCheckReportList.DataSourceID = "";
            gridTicketCheckReportList.DataSource = sdsTicketCheckReportList;
            gridTicketCheckReportList.DataBind();
        }

        protected void eInspector_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vInspector_Temp = eInspector_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vInspector_Temp + "' ";
            string vInspectorName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vInspectorName_Temp == "")
            {
                vInspectorName_Temp = vInspector_Temp;
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vInspectorName_Temp + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vInspector_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eInspector_Search.Text = vInspector_Temp.Trim();
            eInspectorName_Search.Text = vInspectorName_Temp.Trim();
        }

        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDriver_Temp = eDriver_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp == "")
            {
                vDriverName_Temp = vDriver_Temp;
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Search.Text = vDriver_Temp.Trim();
            eDriverName_Search.Text = vDriverName_Temp.Trim();
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
        /// 匯入 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbImportExcel_Click(object sender, EventArgs e)
        {
            string vFileName = fuExcel.FileName;
            if (vFileName != "")
            {
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vUploadFileName = vUploadPath + fuExcel.FileName;
                fuExcel.SaveAs(vUploadFileName);
                string vExtName = Path.GetExtension(vFileName);
                //以下是 EXCEL 裡要用到的欄位
                string vCaseNo = "";
                string vLinesNo = "";
                string vCar_ID = "";
                string vDeparturePosition = "";
                string vDepartureTime = "";
                string vArrivePosition = "";
                string vArriveTime = "";
                string vDriverName = "";
                string vDriver = "";
                string vDepNo = "";
                string vHasAssetNo = "";
                string vHasSecurityLabel = "";
                string vCheckNote = "";
                string vCaseNote = "";
                string vCaseYM = "";
                string vSQLStr_Temp = "";
                string vMaxCaseNo = "";
                string vRemark = "";
                string vInspector = "";
                string vStationToilet = "";
                string vStationLight = "";
                int vIndex = 0;
                DateTime vCaseDate = DateTime.Today;
                string vCaseDate_Str = "";
                DateTime vBuDate = DateTime.Today;
                string vBuDate_Str = (vBuDate.Year - 1911).ToString() + "/" + vBuDate.ToString("MM/dd");
                //========================================================================
                if (vExtName == ".xls")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                    HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                    for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                    {
                        HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                        if (vRowTemp_H != null)
                        {
                            vLinesNo = vRowTemp_H.Cells[0].StringCellValue.Trim();
                            if (vLinesNo.IndexOf("中華民國") > -1)
                            {
                                vLinesNo = vLinesNo.Replace("中華民國", "");
                                vLinesNo = vLinesNo.Replace("年", "/");
                                vLinesNo = vLinesNo.Replace("月", "/");
                                vLinesNo = vLinesNo.Replace("日", "");
                                vCaseDate = DateTime.Parse(vLinesNo);
                                vCaseDate_Str = (vCaseDate.Year - 1911).ToString() + "/" + vCaseDate.ToString("MM/dd");
                            }
                            else if (vLinesNo.Trim() == "日報")
                            {
                                vRowTemp_H.Cells[1].SetCellType(CellType.String);
                                vInspector = vRowTemp_H.Cells[1].StringCellValue.Trim();
                            }
                            else if ((vLinesNo.Trim() != "路線") && (vLinesNo.Trim() != "查票工作報告"))
                            {
                                if (vLinesNo.Trim() != "")
                                {
                                    vCaseYM = vCaseDate.Year.ToString("D4") + vCaseDate.Month.ToString("D2");
                                    vSQLStr_Temp = "select max(CaseNo) MaxCaseNo from TicketCheckReport where CaseNo like '" + vCaseYM + "%' ";
                                    vMaxCaseNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxCaseNo");
                                    vIndex = (vMaxCaseNo != "") ? Int32.Parse(vMaxCaseNo.Replace(vCaseYM, "").Trim()) + 1 : 1;
                                    vCaseNo = vCaseYM + vIndex.ToString("D4");
                                    vCar_ID = vRowTemp_H.Cells[1].StringCellValue.Trim();
                                    vDeparturePosition = vRowTemp_H.Cells[2].StringCellValue.Trim();
                                    vDepartureTime = vRowTemp_H.Cells[3].StringCellValue.Trim();
                                    vArrivePosition = vRowTemp_H.Cells[4].StringCellValue.Trim();
                                    vArriveTime = vRowTemp_H.Cells[5].StringCellValue.Trim();
                                    vDriverName = vRowTemp_H.Cells[6].StringCellValue.Trim();
                                    vDriver = vRowTemp_H.Cells[7].StringCellValue.Trim();
                                    vDepNo = PF.GetValue(vConnStr, "select DepNo from Employee where EmpNo = '" + vDriver + "' ", "DepNo");
                                    vHasAssetNo = (vRowTemp_H.Cells[8].StringCellValue.Trim() == "有") ? "V" :
                                                  (vRowTemp_H.Cells[8].StringCellValue.Trim() == "無") ? "X" : "";
                                    vHasSecurityLabel = (vRowTemp_H.Cells[9].StringCellValue.Trim() == "有") ? "V" :
                                                        (vRowTemp_H.Cells[9].StringCellValue.Trim() == "無") ? "X" : "";
                                    vRemark = ((vHasAssetNo == "")) ? vRowTemp_H.Cells[8].StringCellValue.Trim() : "";
                                    vRemark = ((vHasSecurityLabel == "")) ? vRemark + Environment.NewLine + vRowTemp_H.Cells[9].StringCellValue.Trim() : vRemark;
                                    vStationToilet = (vRowTemp_H.Cells.Count >= 11) ? vRowTemp_H.Cells[10].StringCellValue.Trim() : "";
                                    vStationToilet = PF.GetValue(vConnStr, "select ClassNo from DBDICB where FKey = '查票工作報告    TicketCheckP    StationToilet' and ClassTxt = '" + vStationToilet.Trim() + "' ", "ClassNo").Trim();
                                    vStationLight = (vRowTemp_H.Cells.Count >= 12) ? vRowTemp_H.Cells[11].StringCellValue.Trim() : "";
                                    vStationLight = (vStationLight == "正常") ? "N" : (vStationLight == "異常") ? "X" : vStationLight;
                                    vCheckNote = (vRowTemp_H.Cells.Count >= 13) ? vRowTemp_H.Cells[12].StringCellValue.Trim() : "";
                                    vCaseNote = (vRowTemp_H.Cells.Count >= 14) ? vRowTemp_H.Cells[13].StringCellValue.Trim() : "";
                                    using (SqlDataSource dsTemp = new SqlDataSource())
                                    {
                                        dsTemp.ConnectionString = (vConnStr == "") ? PF.GetConnectionStr(Request.ApplicationPath) : vConnStr;
                                        dsTemp.InsertCommand = "INSERT INTO TicketCheckReport " + Environment.NewLine +
                                                               "       (CaseNo, LinesNo, DepNo, Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, ArrivePosition, " + Environment.NewLine +
                                                               "        ArriveTime, Car_ID, CheckNote, CaseNote, Remark, BuDate, BuMan, Inspector, HasAssetNo, HasSecurityLabel, " + Environment.NewLine +
                                                               "        StationToilet, StationLight) " + Environment.NewLine +
                                                               "VALUES (@CaseNo, @LinesNo, @DepNo, @Driver, @DriverName, @CaseDate, @DeparturePosition, @DepartureTime, " + Environment.NewLine +
                                                               "        @ArrivePosition, @ArriveTime, @Car_ID, @CheckNote, @CaseNote, @Remark, @BuDate, @BuMan, @Inspector, " + Environment.NewLine +
                                                               "        @HasAssetNo, @HasSecurityLabel, @StationToilet, @StationLight)";
                                        dsTemp.InsertParameters.Clear();
                                        dsTemp.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                        dsTemp.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                        dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                        dsTemp.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                        dsTemp.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                        dsTemp.InsertParameters.Add(new Parameter("CaseDate", DbType.DateTime, vCaseDate_Str));
                                        dsTemp.InsertParameters.Add(new Parameter("DeparturePosition", DbType.String, vDeparturePosition));
                                        dsTemp.InsertParameters.Add(new Parameter("DepartureTime", DbType.String, vDepartureTime));
                                        dsTemp.InsertParameters.Add(new Parameter("ArrivePosition", DbType.String, vArrivePosition));
                                        dsTemp.InsertParameters.Add(new Parameter("ArriveTime", DbType.String, vArriveTime));
                                        dsTemp.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                        dsTemp.InsertParameters.Add(new Parameter("CheckNote", DbType.String, vCheckNote));
                                        dsTemp.InsertParameters.Add(new Parameter("CaseNote", DbType.String, vCaseNote));
                                        dsTemp.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                        dsTemp.InsertParameters.Add(new Parameter("BuDate", DbType.DateTime, vBuDate_Str));
                                        dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                        dsTemp.InsertParameters.Add(new Parameter("Inspector", DbType.String, vInspector));
                                        dsTemp.InsertParameters.Add(new Parameter("HasAssetNo", DbType.String, vHasAssetNo));
                                        dsTemp.InsertParameters.Add(new Parameter("HasSecurityLabel", DbType.String, vHasSecurityLabel));
                                        dsTemp.InsertParameters.Add(new Parameter("StationToilet", DbType.String, vStationToilet));
                                        dsTemp.InsertParameters.Add(new Parameter("StationLight", DbType.String, vStationLight));
                                        dsTemp.Insert();
                                    }
                                }
                            }
                        }
                    }
                }
                else if (vExtName == ".xlsx")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                    XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                    for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                    {
                        XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                        if (vRowTemp_X != null)
                        {
                            vLinesNo = vRowTemp_X.Cells[0].StringCellValue.Trim();
                            if (vLinesNo.IndexOf("中華民國") > -1)
                            {
                                vLinesNo = vLinesNo.Replace("中華民國", "");
                                vLinesNo = vLinesNo.Replace("年", "/");
                                vLinesNo = vLinesNo.Replace("月", "/");
                                vLinesNo = vLinesNo.Replace("日", "");
                                vCaseDate = DateTime.Parse(vLinesNo);
                                vCaseDate_Str = (vCaseDate.Year - 1911).ToString() + "/" + vCaseDate.ToString("MM/dd");
                            }
                            else if (vLinesNo.Trim() == "日報")
                            {
                                vRowTemp_X.Cells[1].SetCellType(CellType.String);
                                vInspector = vRowTemp_X.Cells[1].StringCellValue.Trim();
                            }
                            else if ((vLinesNo.Trim() != "路線") && (vLinesNo.Trim() != "查票工作報告"))
                            {
                                if (vLinesNo.Trim() != "")
                                {
                                    vCaseYM = vCaseDate.Year.ToString("D4") + vCaseDate.Month.ToString("D2");
                                    vSQLStr_Temp = "select max(CaseNo) MaxCaseNo from TicketCheckReport where CaseNo like '" + vCaseYM + "%' ";
                                    vMaxCaseNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxCaseNo");
                                    vIndex = (vMaxCaseNo != "") ? Int32.Parse(vMaxCaseNo.Replace(vCaseYM, "").Trim()) + 1 : 1;
                                    vCaseNo = vCaseYM + vIndex.ToString("D4");
                                    vCar_ID = vRowTemp_X.Cells[1].StringCellValue.Trim();
                                    vDeparturePosition = vRowTemp_X.Cells[2].StringCellValue.Trim();
                                    vDepartureTime = vRowTemp_X.Cells[3].StringCellValue.Trim();
                                    vArrivePosition = vRowTemp_X.Cells[4].StringCellValue.Trim();
                                    vArriveTime = vRowTemp_X.Cells[5].StringCellValue.Trim();
                                    vDriverName = vRowTemp_X.Cells[6].StringCellValue.Trim();
                                    vDriver = vRowTemp_X.Cells[7].StringCellValue.Trim();
                                    vDepNo = PF.GetValue(vConnStr, "select DepNo from Employee where EmpNo = '" + vDriver + "' ", "DepNo");
                                    vHasAssetNo = (vRowTemp_X.Cells[8].StringCellValue.Trim() == "有") ? "V" :
                                                  (vRowTemp_X.Cells[8].StringCellValue.Trim() == "無") ? "X" : "";
                                    vHasSecurityLabel = (vRowTemp_X.Cells[9].StringCellValue.Trim() == "有") ? "V" :
                                                        (vRowTemp_X.Cells[9].StringCellValue.Trim() == "無") ? "X" : "";
                                    vRemark = ((vHasAssetNo == "")) ? vRowTemp_X.Cells[8].StringCellValue.Trim() : "";
                                    vRemark = ((vHasSecurityLabel == "")) ? vRemark + Environment.NewLine + vRowTemp_X.Cells[9].StringCellValue.Trim() : vRemark;
                                    vStationToilet = (vRowTemp_X.Cells.Count >= 11) ? vRowTemp_X.Cells[10].StringCellValue.Trim() : "";
                                    vStationToilet = PF.GetValue(vConnStr, "select ClassNo from DBDICB where FKey = '查票工作報告    TicketCheckP    StationToilet' and ClassTxt = '" + vStationToilet.Trim() + "' ", "ClassNo").Trim();
                                    vStationLight = (vRowTemp_X.Cells.Count >= 12) ? vRowTemp_X.Cells[11].StringCellValue.Trim() : "";
                                    vStationLight = (vStationLight == "正常") ? "N" : (vStationLight == "異常") ? "X" : vStationLight;
                                    vCheckNote = (vRowTemp_X.Cells.Count >= 13) ? vRowTemp_X.Cells[12].StringCellValue.Trim() : "";
                                    vCaseNote = (vRowTemp_X.Cells.Count >= 14) ? vRowTemp_X.Cells[13].StringCellValue.Trim() : "";
                                    using (SqlDataSource dsTemp = new SqlDataSource())
                                    {
                                        dsTemp.ConnectionString = (vConnStr == "") ? PF.GetConnectionStr(Request.ApplicationPath) : vConnStr;
                                        dsTemp.InsertCommand = "INSERT INTO TicketCheckReport " + Environment.NewLine +
                                                               "       (CaseNo, LinesNo, DepNo, Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, ArrivePosition, " + Environment.NewLine +
                                                               "        ArriveTime, Car_ID, CheckNote, CaseNote, Remark, BuDate, BuMan, Inspector, HasAssetNo, HasSecurityLabel, " + Environment.NewLine +
                                                               "        StationToilet, StationLight) " + Environment.NewLine +
                                                               "VALUES (@CaseNo, @LinesNo, @DepNo, @Driver, @DriverName, @CaseDate, @DeparturePosition, @DepartureTime, " + Environment.NewLine +
                                                               "        @ArrivePosition, @ArriveTime, @Car_ID, @CheckNote, @CaseNote, @Remark, @BuDate, @BuMan, @Inspector, " + Environment.NewLine +
                                                               "        @HasAssetNo, @HasSecurityLabel, @StationToilet, @StationLight)";
                                        dsTemp.InsertParameters.Clear();
                                        dsTemp.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo));
                                        dsTemp.InsertParameters.Add(new Parameter("LinesNo", DbType.String, vLinesNo));
                                        dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                        dsTemp.InsertParameters.Add(new Parameter("Driver", DbType.String, vDriver));
                                        dsTemp.InsertParameters.Add(new Parameter("DriverName", DbType.String, vDriverName));
                                        dsTemp.InsertParameters.Add(new Parameter("CaseDate", DbType.DateTime, vCaseDate_Str));
                                        dsTemp.InsertParameters.Add(new Parameter("DeparturePosition", DbType.String, vDeparturePosition));
                                        dsTemp.InsertParameters.Add(new Parameter("DepartureTime", DbType.String, vDepartureTime));
                                        dsTemp.InsertParameters.Add(new Parameter("ArrivePosition", DbType.String, vArrivePosition));
                                        dsTemp.InsertParameters.Add(new Parameter("ArriveTime", DbType.String, vArriveTime));
                                        dsTemp.InsertParameters.Add(new Parameter("Car_ID", DbType.String, vCar_ID));
                                        dsTemp.InsertParameters.Add(new Parameter("CheckNote", DbType.String, vCheckNote));
                                        dsTemp.InsertParameters.Add(new Parameter("CaseNote", DbType.String, vCaseNote));
                                        dsTemp.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark));
                                        dsTemp.InsertParameters.Add(new Parameter("BuDate", DbType.DateTime, vBuDate_Str));
                                        dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                        dsTemp.InsertParameters.Add(new Parameter("Inspector", DbType.String, vInspector));
                                        dsTemp.InsertParameters.Add(new Parameter("HasAssetNo", DbType.String, vHasAssetNo));
                                        dsTemp.InsertParameters.Add(new Parameter("HasSecurityLabel", DbType.String, vHasSecurityLabel));
                                        dsTemp.InsertParameters.Add(new Parameter("StationToilet", DbType.String, vStationToilet));
                                        dsTemp.InsertParameters.Add(new Parameter("StationLight", DbType.String, vStationLight));
                                        dsTemp.Insert();
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

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 預覽工作報告報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vSelStr = GetSelectStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintPoint = new SqlDataAdapter(vSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintPoint = new DataTable("TicketCheckReportP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("TicketCheckReportP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\TicketCheckReportP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "查票工作報告"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;

                    string vCaseDateStr = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "從 " + eCaseDate_S_Search.Text.Trim() + " 起至 " + eCaseDate_E_Search.Text.Trim() + " 止" :
                                          ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? eCaseDate_S_Search.Text.Trim() :
                                          ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? eCaseDate_E_Search.Text.Trim() : "不分日期";
                    string vInspectorStr = (eInspector_Search.Text.Trim() != "") ? eInspector_Search.Text.Trim() : "全部";
                    string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "全部";
                    string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                    string vLinesNoStr = (eLinesNo_Search.Text.Trim() != "") ? eLinesNo_Search.Text.Trim() : "全部";
                    string vRecordNote = "預覽報表_查票工作報告" + Environment.NewLine +
                                         "TicketCheckP.aspx" + Environment.NewLine +
                                         "查票日期：" + vCaseDateStr + Environment.NewLine +
                                         "查票人：" + vInspectorStr + Environment.NewLine +
                                         "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                         "駕駛員：" + vDriverStr + Environment.NewLine +
                                         "路線：" + vLinesNoStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        /// <summary>
        /// GRID 換頁
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gridTicketCheckReportList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridTicketCheckReportList.PageIndex = e.NewPageIndex;
            gridTicketCheckReportList.DataBind();
        }

        /// <summary>
        /// 明細畫面資料完成繫結
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void fvTicketCheckReportDetail_DataBound(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCaseDateURL_Temp = "";
            string vCaseDateScript_Temp = "";
            string vSQLStr_Temp = "";

            switch (fvTicketCheckReportDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Button bbEdit_Temp = (Button)fvTicketCheckReportDetail.FindControl("bbEdit_List");
                    Button bbDel_Temp = (Button)fvTicketCheckReportDetail.FindControl("bbDel_List");
                    Button bbCreateRPReport_Temp = (Button)fvTicketCheckReportDetail.FindControl("bbCreateRPReport_List");
                    Button bbPrintRPReport_Temp = (Button)fvTicketCheckReportDetail.FindControl("bbPrintRpReport_List");
                    Label eIsPassed_Temp = (Label)fvTicketCheckReportDetail.FindControl("eIsPassed_List");
                    CheckBox cbIsPassed_TempR = (CheckBox)fvTicketCheckReportDetail.FindControl("cbIsPassed_List");
                    Label eHasAssetNo_List = (Label)fvTicketCheckReportDetail.FindControl("eHasAssetNo_List");
                    CheckBox cbHasAssetNo_List = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasAssetNo_List");
                    Label eHasSecurityLabel_List = (Label)fvTicketCheckReportDetail.FindControl("eHasSecurityLabel_List");
                    CheckBox cbHasSecurityLabel_List = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasSecurityLabel_List");

                    if ((bbEdit_Temp != null) && (cbIsPassed_TempR != null))
                    {
                        cbIsPassed_TempR.Checked = (eIsPassed_Temp.Text.Trim() == "True");
                        bbEdit_Temp.Visible = !cbIsPassed_TempR.Checked;
                    }
                    if ((bbDel_Temp != null) && (cbIsPassed_TempR != null))
                    {
                        cbIsPassed_TempR.Checked = (eIsPassed_Temp.Text.Trim() == "True");
                        bbDel_Temp.Visible = !cbIsPassed_TempR.Checked;
                    }
                    if ((bbCreateRPReport_Temp != null) && (cbIsPassed_TempR != null))
                    {
                        cbIsPassed_TempR.Checked = (eIsPassed_Temp.Text.Trim() == "True");
                        bbCreateRPReport_Temp.Visible = !cbIsPassed_TempR.Checked;
                        bbPrintRPReport_Temp.Visible = cbIsPassed_TempR.Checked;
                    }
                    if (cbHasAssetNo_List != null)
                    {
                        cbHasAssetNo_List.Checked = (eHasAssetNo_List.Text.Trim() == "V");
                        cbHasSecurityLabel_List.Checked = (eHasSecurityLabel_List.Text.Trim() == "V");
                    }
                    break;
                case FormViewMode.Edit:
                    TextBox eCaseDate_TempE = (TextBox)fvTicketCheckReportDetail.FindControl("eCaseDate_Edit");
                    Label eModifyDate_Temp = (Label)fvTicketCheckReportDetail.FindControl("eModifyDate_Edit");
                    Label eModifyMan_Temp = (Label)fvTicketCheckReportDetail.FindControl("eModifyMan_Edit");
                    Label eModifyManName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eModifyManName_Edit");
                    Label eIsPassed_TempE = (Label)fvTicketCheckReportDetail.FindControl("eIsPassed_Edit");
                    CheckBox cbIsPassed_TempE = (CheckBox)fvTicketCheckReportDetail.FindControl("cbIsPassed_Edit");
                    Label eHasAssetNo_Edit = (Label)fvTicketCheckReportDetail.FindControl("eHasAssetNo_Edit");
                    CheckBox cbHasAssetNo_Edit = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasAssetNo_Edit");
                    Label eHasSecurityLabel_Edit = (Label)fvTicketCheckReportDetail.FindControl("eHasSecurityLabel_Edit");
                    CheckBox cbHasSecurityLabel_Edit = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasSecurityLabel_Edit");
                    DropDownList ddlStationLight_Edit = (DropDownList)fvTicketCheckReportDetail.FindControl("ddlStationLight_Edit");
                    TextBox eStationLight_Edit = (TextBox)fvTicketCheckReportDetail.FindControl("eStationLight_Edit");
                    DropDownList ddlStationToilet_Edit = (DropDownList)fvTicketCheckReportDetail.FindControl("ddlStationToilet_Edit");
                    TextBox eStationToilet_Edit = (TextBox)fvTicketCheckReportDetail.FindControl("eStationToilet_Edit");

                    if (cbIsPassed_TempE != null)
                    {
                        cbIsPassed_TempE.Checked = (eIsPassed_TempE.Text.Trim() == "True");
                        cbHasAssetNo_Edit.Checked = (eHasAssetNo_Edit.Text.Trim() == "V");
                        cbHasSecurityLabel_Edit.Checked = (eHasSecurityLabel_Edit.Text.Trim() == "V");
                    }

                    if ((ddlStationLight_Edit != null) && (eStationLight_Edit != null))
                    {
                        ddlStationLight_Edit.SelectedIndex = ddlStationLight_Edit.Items.IndexOf(ddlStationLight_Edit.Items.FindByValue(eStationLight_Edit.Text.Trim()));
                    }

                    if ((ddlStationToilet_Edit != null) && (eStationToilet_Edit != null))
                    {
                        ddlStationToilet_Edit.SelectedIndex = ddlStationToilet_Edit.Items.IndexOf(ddlStationToilet_Edit.Items.FindByValue(eStationToilet_Edit.Text.Trim()));
                    }

                    eModifyDate_Temp.Text = PF.GetChinsesDate(DateTime.Today);
                    eModifyMan_Temp.Text = vLoginID.Trim();
                    vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vLoginID.Trim() + "' ";
                    eModifyManName_Temp.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");

                    vCaseDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_TempE.ClientID;
                    vCaseDateScript_Temp = "window.open('" + vCaseDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_TempE.Attributes["onClick"] = vCaseDateScript_Temp;
                    break;
                case FormViewMode.Insert:
                    TextBox eCaseDate_TempI = (TextBox)fvTicketCheckReportDetail.FindControl("eCaseDate_INS");
                    Label eBuDate_Temp = (Label)fvTicketCheckReportDetail.FindControl("eBuDate_INS");
                    Label eBuMan_Temp = (Label)fvTicketCheckReportDetail.FindControl("eBuMan_INS");
                    Label eBuManName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eBuManName_INS");
                    Label eIsPassed_TempI = (Label)fvTicketCheckReportDetail.FindControl("eIsPassed_INS");
                    CheckBox cbIsPassed_TempI = (CheckBox)fvTicketCheckReportDetail.FindControl("cbIsPassed_INS");
                    Label eHasAssetNo_INS = (Label)fvTicketCheckReportDetail.FindControl("eHasAssetNo_INS");
                    CheckBox cbHasAssetNo_INS = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasAssetNo_INS");
                    Label eHasSecurityLabel_INS = (Label)fvTicketCheckReportDetail.FindControl("eHasSecurityLabel_INS");
                    CheckBox cbHasSecurityLabel_INS = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasSecurityLabel_INS");
                    DropDownList ddlStationLight_INS = (DropDownList)fvTicketCheckReportDetail.FindControl("ddlStationLight_INS");
                    TextBox eStationLight_INS = (TextBox)fvTicketCheckReportDetail.FindControl("eStationLight_INS");

                    if (cbIsPassed_TempI != null)
                    {
                        cbIsPassed_TempI.Checked = (eIsPassed_TempI.Text.Trim() == "True");
                        cbHasAssetNo_INS.Checked = (eHasAssetNo_INS.Text.Trim() == "V");
                        cbHasSecurityLabel_INS.Checked = (eHasSecurityLabel_INS.Text.Trim() == "V");
                        ddlStationLight_INS.SelectedIndex = 1;
                        eStationLight_INS.Text = ddlStationLight_INS.SelectedValue.Trim();
                    }

                    eBuDate_Temp.Text = PF.GetChinsesDate(DateTime.Today);
                    eBuMan_Temp.Text = vLoginID.Trim();
                    vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vLoginID.Trim() + "' ";
                    eBuManName_Temp.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "Name").Trim();

                    vCaseDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_TempI.ClientID;
                    vCaseDateScript_Temp = "window.open('" + vCaseDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_TempI.Attributes["onClick"] = vCaseDateScript_Temp;
                    break;
            }
        }

        protected void eDriver_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eDriver_Edit");
            Label eDriverName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDriverName_Edit");
            string vDriver_Temp = eDriver_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp.Trim() + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp == "")
            {
                vDriverName_Temp = vDriver_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Temp.Text = vDriver_Temp.Trim();
            eDriverName_Temp.Text = vDriverName_Temp.Trim();
            if (vDriver_Temp.Trim() != "")
            {
                string vDepData_Temp = PF.GetValue(vConnStr, "select (DepNo + ',' + [Name]) DepData from Department where DepNo = (select DepNo from Employee where EmpNo = '" + vDriver_Temp.Trim() + "') ", "DepData");
                string[] vTempData = vDepData_Temp.Split(',');
                TextBox eDepNo_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eDepNo_Edit");
                Label eDepName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDepName_Edit");
                eDepNo_Temp.Text = vTempData[0].Trim();
                eDepName_Temp.Text = vTempData[1].Trim();
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eDepNo_Edit");
            Label eDepName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDepName_Edit");
            string vDepNo_Temp = eDepNo_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Temp.Text = vDepNo_Temp.Trim();
            eDepName_Temp.Text = vDepName_Temp.Trim();
        }

        protected void cbHasAssetNo_Edit_CheckedChanged(object sender, EventArgs e)
        {
            Label eHasAssetNo_Edit = (Label)fvTicketCheckReportDetail.FindControl("eHasAssetNo_Edit");
            CheckBox cbHasAssetNo_Edit = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasAssetNo_Edit");
            eHasAssetNo_Edit.Text = (cbHasAssetNo_Edit.Checked) ? "V" : "X";
        }

        protected void cbHasSecurityLabel_Edit_CheckedChanged(object sender, EventArgs e)
        {
            Label eHasSecurityLabel_Edit = (Label)fvTicketCheckReportDetail.FindControl("eHasSecurityLabel_Edit");
            CheckBox cbHasSecurityLabel_Edit = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasSecurityLabel_Edit");
            eHasSecurityLabel_Edit.Text = (cbHasSecurityLabel_Edit.Checked) ? "V" : "X";
        }

        protected void ddlStationToilet_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlStationToilet_Temp = (DropDownList)fvTicketCheckReportDetail.FindControl("ddlStationToilet_Edit");
            TextBox eStationToilet_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eStstionLight_INS");
            if ((ddlStationToilet_Temp != null) && (eStationToilet_Temp != null))
            {
                eStationToilet_Temp.Text = ddlStationToilet_Temp.SelectedValue.Trim();
            }
        }

        protected void ddlStationLight_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlStationLight_Temp = (DropDownList)fvTicketCheckReportDetail.FindControl("ddlStationLight_Edit");
            TextBox eStationLight_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eStstionLight_Edit");
            if ((ddlStationLight_Temp != null) && (eStationLight_Temp != null))
            {
                eStationLight_Temp.Text = ddlStationLight_Temp.SelectedValue.Trim();
            }
        }

        protected void eInspector_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eInspector_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eInspector_Edit");
            Label eInspectorName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eInspectorName_Edit");
            string vInspector_Temp = eInspector_Temp.Text.Trim();
            string vSQLStr_Temp = "";
            string vInspectorName_Temp = "";
            if (vInspector_Temp != "")
            {
                vSQLStr_Temp = "select [Name] from EMployee where EmpNo = '" + vInspector_Temp + "' ";
                vInspectorName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vInspectorName_Temp == "")
                {
                    vInspectorName_Temp = vInspector_Temp.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vInspectorName_Temp + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                    vInspector_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eInspector_Temp.Text = vInspector_Temp.Trim();
                eInspectorName_Temp.Text = vInspectorName_Temp.Trim();
            }
        }

        protected void eDriver_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eDriver_INS");
            Label eDriverName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDriverName_INS");
            string vDriver_Temp = eDriver_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp.Trim() + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp == "")
            {
                vDriverName_Temp = vDriver_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Temp.Text = vDriver_Temp.Trim();
            eDriverName_Temp.Text = vDriverName_Temp.Trim();
            if (vDriver_Temp.Trim() != "")
            {
                string vDepData_Temp = PF.GetValue(vConnStr, "select (DepNo + ',' + [Name]) DepData from Department where DepNo = (select DepNo from Employee where EmpNo = '" + vDriver_Temp.Trim() + "') ", "DepData");
                string[] vTempData = vDepData_Temp.Split(',');
                TextBox eDepNo_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eDepNo_INS");
                Label eDepName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDepName_INS");
                eDepNo_Temp.Text = vTempData[0].Trim();
                eDepName_Temp.Text = vTempData[1].Trim();
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eDepNo_INS");
            Label eDepName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDepName_INS");
            string vDepNo_Temp = eDepNo_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Temp.Text = vDepNo_Temp.Trim();
            eDepName_Temp.Text = vDepName_Temp.Trim();
        }

        protected void cbHasAssetNo_INS_CheckedChanged(object sender, EventArgs e)
        {
            Label eHasAssetNo_INS = (Label)fvTicketCheckReportDetail.FindControl("eHasAssetNo_INS");
            CheckBox cbHasAssetNo_INS = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasAssetNo_INS");
            eHasAssetNo_INS.Text = (cbHasAssetNo_INS.Checked) ? "V" : "X";
        }

        protected void cbHasSecurityLabel_INS_CheckedChanged(object sender, EventArgs e)
        {
            Label eHasSecurityLabel_INS = (Label)fvTicketCheckReportDetail.FindControl("eHasSecurityLabel_INS");
            CheckBox cbHasSecurityLabel_INS = (CheckBox)fvTicketCheckReportDetail.FindControl("cbHasSecurityLabel_INS");
            eHasSecurityLabel_INS.Text = (cbHasSecurityLabel_INS.Checked) ? "V" : "X";
        }

        protected void ddlStationToilet_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlStationToilet_Temp = (DropDownList)fvTicketCheckReportDetail.FindControl("ddlStationToilet_Edit");
            TextBox eStationToilet_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eStstionLight_INS");
            if ((ddlStationToilet_Temp != null) && (eStationToilet_Temp != null))
            {
                eStationToilet_Temp.Text = ddlStationToilet_Temp.SelectedValue.Trim();
            }
        }

        protected void ddlStationLight_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlStationLight_Temp = (DropDownList)fvTicketCheckReportDetail.FindControl("ddlStationLight_INS");
            TextBox eStationLight_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eStstionLight_INS");
            if ((ddlStationLight_Temp != null) && (eStationLight_Temp != null))
            {
                eStationLight_Temp.Text = ddlStationLight_Temp.SelectedValue.Trim();
            }
        }

        protected void eInspector_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eInspector_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eInspector_INS");
            Label eInspectorName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eInspectorName_INS");
            string vInspector_Temp = eInspector_Temp.Text.Trim();
            string vSQLStr_Temp = "";
            string vInspectorName_Temp = "";
            if (vInspector_Temp != "")
            {
                vSQLStr_Temp = "select [Name] from EMployee where EmpNo = '" + vInspector_Temp + "' ";
                vInspectorName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vInspectorName_Temp == "")
                {
                    vInspectorName_Temp = vInspector_Temp.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vInspectorName_Temp + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                    vInspector_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eInspector_Temp.Text = vInspector_Temp.Trim();
                eInspectorName_Temp.Text = vInspectorName_Temp.Trim();
            }
        }

        /// <summary>
        /// 刪除工作報告
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDel_List_Click(object sender, EventArgs e)
        {
            CheckBox cbIsPassed_Temp = (CheckBox)fvTicketCheckReportDetail.FindControl("cbIsPassed_List");
            if (cbIsPassed_Temp != null)
            {
                if (!cbIsPassed_Temp.Checked)
                {
                    Label eCaseNo_Temp = (Label)fvTicketCheckReportDetail.FindControl("eCaseNo_List");
                    string vCaseNo_Temp = eCaseNo_Temp.Text.Trim();
                    if (vCaseNo_Temp != "")
                    {
                        //先產生一筆異動記錄
                        string vSQLStr_Temp = "select isnull(MAX(Items), '0') MaxItems from TicketCheckReport_History where CaseNo = '" + vCaseNo_Temp + "' ";
                        int vMaxItems = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems")) + 1;
                        string vNewCaseNoItems = vCaseNo_Temp + vMaxItems.ToString("D4");
                        vSQLStr_Temp = "insert into TicketCheckReport_History (CaseNoItem, Items, ModifyMode, CaseNo, LinesNo, DepNo, Driver, DriverName, " + Environment.NewLine +
                                       "                                       CaseDate, DeparturePosition, DepartureTime, ArrivePosition, ArriveTime, " + Environment.NewLine +
                                       "                                       Car_ID, CheckNote, CaseNote, Remark, Inspector, BuDate, BuMan, " + Environment.NewLine +
                                       "                                       ModifyDate, ModifyMan, IsPassed, RP_ReportNo, HasAssetNo, HasSecurityLabel, " + Environment.NewLine +
                                       "                                       StationToilet, StationLight)" + Environment.NewLine +
                                       "select '" + vNewCaseNoItems.Trim() + "', '" + vMaxItems.ToString("D4") + "', 'DEL', CaseNo, LinesNo, DepNo, " + Environment.NewLine +
                                       "       Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, ArrivePosition, ArriveTime, Car_ID, " + Environment.NewLine +
                                       "       CheckNote, CaseNote, Remark, Inspector, BuDate, BuMan, ModifyDate, ModifyMan, IsPassed, RP_ReportNo, " + Environment.NewLine +
                                       "       HasAssetNo, HasSecurityLabel, StationToilet, StationLight " + Environment.NewLine +
                                       "  from TicketCheckReport " + Environment.NewLine +
                                       " where CaseNo = '" + vCaseNo_Temp + "' ";
                        try
                        {
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            //實際進行資料刪除
                            sdsTicketCheckReportDetail.DeleteParameters.Clear();
                            sdsTicketCheckReportDetail.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                            sdsTicketCheckReportDetail.Delete();

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
                        Response.Write("alert('請先選擇要刪除的查票單！')");
                        Response.Write("</" + "Script>");
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('已開勸導單，不可刪除！')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// 開立勸導單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCreateRPReport_List_Click(object sender, EventArgs e)
        {
            Label eCaseNo_Temp = (Label)fvTicketCheckReportDetail.FindControl("eCaseNo_List");
            string vCaseNo_Temp = eCaseNo_Temp.Text.Trim();
            string vExecStr = "select RP_ReportNo from TicketCheckReport where CaseNo = '" + vCaseNo_Temp + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if ((vCaseNo_Temp != "") && (PF.GetValue(vConnStr, vExecStr, "RP_ReportNo") == ""))
            {
                Label eCaseDate_Temp = (Label)fvTicketCheckReportDetail.FindControl("eCaseDate_List");
                DateTime vCaseDate_Temp = (eCaseDate_Temp.Text.Trim() != "") ? DateTime.Parse(eCaseDate_Temp.Text.Trim()) : DateTime.Today;
                Label eDriver_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDriver_List");
                Label eDeparturePosition_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDeparturePosition_List");
                Label eDepartureTime_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDepartureTime_List");
                string vCaseYear = (vCaseDate_Temp.Year - 1911).ToString() + "A";
                string vSQLStr_Temp = "select MAX(CaseNo) MAXCaseNo from AdviceReport where CaseNo like '" + vCaseYear + "%' ";
                string vMaxCaseNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MAXCaseNo");
                int vCaseIndex = (vMaxCaseNo != "") ? Int32.Parse(vMaxCaseNo.Replace(vCaseYear, "")) + 1 : 1;
                string vAdviceCaseNo = vCaseYear + vCaseIndex.ToString("D4");
                string vEmpNo_Temp = eDriver_Temp.Text.Trim();
                vSQLStr_Temp = "select ([Name] + ',' + DepNo + ',' + Title) ReturnData from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
                string vReturnDataStr = PF.GetValue(vConnStr, vSQLStr_Temp, "ReturnData");
                string[] vReturnData = vReturnDataStr.Split(',');
                string vEmpName = vReturnData[0].ToString();
                string vDepNo = vReturnData[1].ToString();
                string vTitle = vReturnData[2].ToString();
                string vCaseTime = eDepartureTime_Temp.Text.Trim();
                string vPosition = eDeparturePosition_Temp.Text.Trim();
                vExecStr = "insert into AdviceReport(CaseNo, CaseDate, EmpNo, EmpName, DepNo, Title, CaseTime, Position, BuMan, BuDate, SourceType) " + Environment.NewLine +
                           " values('" + vAdviceCaseNo + "', '" + vCaseDate_Temp.Year.ToString() + "/" + vCaseDate_Temp.ToString("MM/dd") + "', '" +
                                         vEmpNo_Temp + "', '" + vEmpName + "', '" + vDepNo + "', '" + vTitle + "', '" + vCaseTime + "', '" +
                                         vPosition + "', '" + vLoginID + "', '" + DateTime.Today.Year.ToString() + "/" + DateTime.Today.ToString("MM/dd") + "', 'T')";
                try
                {
                    PF.ExecSQL(vConnStr, vExecStr);
                    vExecStr = "update TicketCheckreport set IsPassed = 1, RP_ReportNo = '" + vAdviceCaseNo + "' where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vExecStr);
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
                Response.Write("alert('請先選擇查票報告項目！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 列印勸導單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintRPReport_List_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eRP_ReportNo_Temp = (Label)fvTicketCheckReportDetail.FindControl("eRP_ReportNo_List");
            if (eRP_ReportNo_Temp.Text.Trim() != "")
            {
                string vRPReportNo = eRP_ReportNo_Temp.Text.Trim();
                string vSQLStr_Temp = "SELECT CaseNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, " + Environment.NewLine +
                                      "       EmpNo, EmpName, Title, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = a.Title)) AS Title_C, " + Environment.NewLine +
                                      "       CaseDate, CaseTime, Position, CaseNote, Remark, DenialReason, BuDate, " + Environment.NewLine +
                                      "       BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, " + Environment.NewLine +
                                      "       ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName " + Environment.NewLine +
                                      "  FROM AdviceReport AS a " + Environment.NewLine +
                                      " WHERE CaseNo = '" + vRPReportNo + "' ";
                using (SqlConnection connPrintAdviceReport = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daPrintAdviceReport = new SqlDataAdapter(vSQLStr_Temp, connPrintAdviceReport);
                    connPrintAdviceReport.Open();
                    DataTable dtPrintAdviceReport = new DataTable();
                    daPrintAdviceReport.Fill(dtPrintAdviceReport);
                    if (dtPrintAdviceReport.Rows.Count > 0)
                    {
                        ReportDataSource rdsPrint = new ReportDataSource("AdviceReportP", dtPrintAdviceReport);

                        rvPrint.LocalReport.DataSources.Clear();
                        rvPrint.LocalReport.ReportPath = @"Report\AdviceReportP.rdlc";
                        rvPrint.LocalReport.DataSources.Add(rdsPrint);
                        rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "桃園客運員工疏失（違規）改善勸導單"));
                        rvPrint.LocalReport.Refresh();
                        plPrint.Visible = true;
                        plShowData.Visible = false;
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到正確的勸導單，請重新確認！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 列印違規通知單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintViolationReport_List_Click(object sender, EventArgs e)
        {
            Label eCaseNo_Temp = (Label)fvTicketCheckReportDetail.FindControl("eCaseNo_List");
            string vCaseNo_Temp = eCaseNo_Temp.Text.Trim();
            if (vCaseNo_Temp.Trim() != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                using (SqlConnection connViolationReport = new SqlConnection(vConnStr))
                {
                    string vSQLSelStr = GetSelectStr_Print(vCaseNo_Temp);
                    vSQLSelStr = vSQLSelStr + Environment.NewLine +
                                 " union all " + Environment.NewLine +
                                 "select 'ZZ001' as CaseNo, cast('' as varchar) as LinesNo, cast('' as varchar) as DepNo, cast('' as varchar) as DepName, " + Environment.NewLine +
                                 "       cast('' as varchar) as Driver, cast('' as varchar) as DriverName, cast(null as datetime) as CaseDate, " + Environment.NewLine +
                                 "       cast('' as varchar) as DeparturePosition, cast('' as varchar) as DeaprtureTime, " + Environment.NewLine +
                                 "       cast('' as varchar) as ArrivePosition, cast('' as varchar) as ArriveTime, cast('' as varchar) as Car_ID, " + Environment.NewLine +
                                 "       cast('' as varchar) as CheckNote, cast('' as varchar) as CaseNote, cast('' as varchar) as Remark, " + Environment.NewLine +
                                 "       cast(null as datetime) as BuDate, cast('' as varchar) as BuMan, cast('' as varchar) as BuManName, " + Environment.NewLine +
                                 "       cast(null as datetime) as ModifyDate, cast('' as varchar) as ModifyMan, cast('' as varchar) as ModifyManName, " + Environment.NewLine +
                                 "       cast('' as varchar) as IsPassed, cast('' as varchar) as RP_ReportNo, " + Environment.NewLine +
                                 "       cast('' as varchar) as Inspector, cast('' as varchar) as InspectorName, " + Environment.NewLine +
                                 "       cast('' as varchar) as HasAssetNo, cast('' as varchar) as HasSecurityLabel, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationLight, cast('' as varchar) as StationLight_C, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationToilet, cast('' as varchar) as StationToilet_C " + Environment.NewLine +
                                 " union all " + Environment.NewLine +
                                 "select 'ZZ002' as CaseNo, cast('' as varchar) as LinesNo, cast('' as varchar) as DepNo, cast('' as varchar) as DepName, " + Environment.NewLine +
                                 "       cast('' as varchar) as Driver, cast('' as varchar) as DriverName, cast(null as datetime) as CaseDate, " + Environment.NewLine +
                                 "       cast('' as varchar) as DeparturePosition, cast('' as varchar) as DeaprtureTime, " + Environment.NewLine +
                                 "       cast('' as varchar) as ArrivePosition, cast('' as varchar) as ArriveTime, cast('' as varchar) as Car_ID, " + Environment.NewLine +
                                 "       cast('' as varchar) as CheckNote, cast('' as varchar) as CaseNote, cast('' as varchar) as Remark, " + Environment.NewLine +
                                 "       cast(null as datetime) as BuDate, cast('' as varchar) as BuMan, cast('' as varchar) as BuManName, " + Environment.NewLine +
                                 "       cast(null as datetime) as ModifyDate, cast('' as varchar) as ModifyMan, cast('' as varchar) as ModifyManName, " + Environment.NewLine +
                                 "       cast('' as varchar) as IsPassed, cast('' as varchar) as RP_ReportNo, " + Environment.NewLine +
                                 "       cast('' as varchar) as Inspector, cast('' as varchar) as InspectorName, " + Environment.NewLine +
                                 "       cast('' as varchar) as HasAssetNo, cast('' as varchar) as HasSecurityLabel, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationLight, cast('' as varchar) as StationLight_C, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationToilet, cast('' as varchar) as StationToilet_C " + Environment.NewLine +
                                 " union all " + Environment.NewLine +
                                 "select 'ZZ003' as CaseNo, cast('' as varchar) as LinesNo, cast('' as varchar) as DepNo, cast('' as varchar) as DepName, " + Environment.NewLine +
                                 "       cast('' as varchar) as Driver, cast('' as varchar) as DriverName, cast(null as datetime) as CaseDate, " + Environment.NewLine +
                                 "       cast('' as varchar) as DeparturePosition, cast('' as varchar) as DeaprtureTime, " + Environment.NewLine +
                                 "       cast('' as varchar) as ArrivePosition, cast('' as varchar) as ArriveTime, cast('' as varchar) as Car_ID, " + Environment.NewLine +
                                 "       cast('' as varchar) as CheckNote, cast('' as varchar) as CaseNote, cast('' as varchar) as Remark, " + Environment.NewLine +
                                 "       cast(null as datetime) as BuDate, cast('' as varchar) as BuMan, cast('' as varchar) as BuManName, " + Environment.NewLine +
                                 "       cast(null as datetime) as ModifyDate, cast('' as varchar) as ModifyMan, cast('' as varchar) as ModifyManName, " + Environment.NewLine +
                                 "       cast('' as varchar) as IsPassed, cast('' as varchar) as RP_ReportNo, " + Environment.NewLine +
                                 "       cast('' as varchar) as Inspector, cast('' as varchar) as InspectorName, " + Environment.NewLine +
                                 "       cast('' as varchar) as HasAssetNo, cast('' as varchar) as HasSecurityLabel, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationLight, cast('' as varchar) as StationLight_C, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationToilet, cast('' as varchar) as StationToilet_C " + Environment.NewLine +
                                 " union all " + Environment.NewLine +
                                 "select 'ZZ004' as CaseNo, cast('' as varchar) as LinesNo, cast('' as varchar) as DepNo, cast('' as varchar) as DepName, " + Environment.NewLine +
                                 "       cast('' as varchar) as Driver, cast('' as varchar) as DriverName, cast(null as datetime) as CaseDate, " + Environment.NewLine +
                                 "       cast('' as varchar) as DeparturePosition, cast('' as varchar) as DeaprtureTime, " + Environment.NewLine +
                                 "       cast('' as varchar) as ArrivePosition, cast('' as varchar) as ArriveTime, cast('' as varchar) as Car_ID, " + Environment.NewLine +
                                 "       cast('' as varchar) as CheckNote, cast('' as varchar) as CaseNote, cast('' as varchar) as Remark, " + Environment.NewLine +
                                 "       cast(null as datetime) as BuDate, cast('' as varchar) as BuMan, cast('' as varchar) as BuManName, " + Environment.NewLine +
                                 "       cast(null as datetime) as ModifyDate, cast('' as varchar) as ModifyMan, cast('' as varchar) as ModifyManName, " + Environment.NewLine +
                                 "       cast('' as varchar) as IsPassed, cast('' as varchar) as RP_ReportNo, " + Environment.NewLine +
                                 "       cast('' as varchar) as Inspector, cast('' as varchar) as InspectorName, " + Environment.NewLine +
                                 "       cast('' as varchar) as HasAssetNo, cast('' as varchar) as HasSecurityLabel, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationLight, cast('' as varchar) as StationLight_C, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationToilet, cast('' as varchar) as StationToilet_C " + Environment.NewLine +
                                 " union all " + Environment.NewLine +
                                 "select 'ZZ005' as CaseNo, cast('' as varchar) as LinesNo, cast('' as varchar) as DepNo, cast('' as varchar) as DepName, " + Environment.NewLine +
                                 "       cast('' as varchar) as Driver, cast('' as varchar) as DriverName, cast(null as datetime) as CaseDate, " + Environment.NewLine +
                                 "       cast('' as varchar) as DeparturePosition, cast('' as varchar) as DeaprtureTime, " + Environment.NewLine +
                                 "       cast('' as varchar) as ArrivePosition, cast('' as varchar) as ArriveTime, cast('' as varchar) as Car_ID, " + Environment.NewLine +
                                 "       cast('' as varchar) as CheckNote, cast('' as varchar) as CaseNote, cast('' as varchar) as Remark, " + Environment.NewLine +
                                 "       cast(null as datetime) as BuDate, cast('' as varchar) as BuMan, cast('' as varchar) as BuManName, " + Environment.NewLine +
                                 "       cast(null as datetime) as ModifyDate, cast('' as varchar) as ModifyMan, cast('' as varchar) as ModifyManName, " + Environment.NewLine +
                                 "       cast('' as varchar) as IsPassed, cast('' as varchar) as RP_ReportNo, " + Environment.NewLine +
                                 "       cast('' as varchar) as Inspector, cast('' as varchar) as InspectorName, " + Environment.NewLine +
                                 "       cast('' as varchar) as HasAssetNo, cast('' as varchar) as HasSecurityLabel, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationLight, cast('' as varchar) as StationLight_C, " + Environment.NewLine +
                                 "       cast('' as varchar) as StationToilet, cast('' as varchar) as StationToilet_C ";
                    Label eDepName_Temp = (Label)fvTicketCheckReportDetail.FindControl("eDepName_List");
                    string vDepName = eDepName_Temp.Text.Trim();
                    SqlDataAdapter daViolationReport = new SqlDataAdapter(vSQLSelStr, connViolationReport);
                    connViolationReport.Open();
                    DataTable dtViolationReport = new DataTable();
                    daViolationReport.Fill(dtViolationReport);

                    if (dtViolationReport.Rows.Count > 0)
                    {
                        //查詢有資料才列印
                        ReportDataSource rdsPrint = new ReportDataSource("TicketCheckReportP", dtViolationReport);

                        rvPrint.LocalReport.DataSources.Clear();
                        rvPrint.LocalReport.ReportPath = @"Report\ViolationReportP.rdlc";
                        rvPrint.LocalReport.DataSources.Add(rdsPrint);
                        rvPrint.LocalReport.SetParameters(new ReportParameter("DepName", vDepName));
                        rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "行車人員違規事項通知單"));
                        rvPrint.LocalReport.Refresh();
                        plPrint.Visible = true;
                        plShowData.Visible = false;
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('查無符合條件的違規通知可列印')");
                        Response.Write("</" + "Script>");
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到正確的勸導單，請重新確認！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 結束報表預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
        }

        /// <summary>
        /// 新增資料存檔前
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void sdsTicketCheckReportDetail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eCaseDate_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eCaseDate_INS");
            DateTime vCaseDate_Temp = (eCaseDate_Temp.Text.Trim() != "") ? DateTime.Parse(eCaseDate_Temp.Text.Trim()) : DateTime.Today;
            Label eBuDate_Temp = (Label)fvTicketCheckReportDetail.FindControl("eBuDate_INS");
            DateTime vBuDate_Temp = (eBuDate_Temp.Text.Trim() != "") ? DateTime.Parse(eBuDate_Temp.Text.Trim()) : DateTime.Today;
            string vCaseYM = vCaseDate_Temp.Year.ToString("D4") + vCaseDate_Temp.Month.ToString("D2");
            string vSQLStr_Temp = "select max(CaseNo) MaxCaseNo from TicketCheckReport where CaseNo like '" + vCaseYM + "%' ";
            string vMaxCaseNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxCaseNo");
            int vIndex = (vMaxCaseNo != "") ? Int32.Parse(vMaxCaseNo.Replace(vCaseYM, "").Trim()) + 1 : 1;
            string vNewCaseNo = vCaseYM + vIndex.ToString("D4");

            e.Command.Parameters["@CaseNo"].Value = vNewCaseNo;
            e.Command.Parameters["@CaseDate"].Value = vCaseDate_Temp;
            e.Command.Parameters["@BuDate"].Value = vBuDate_Temp;
        }

        protected void sdsTicketCheckReportDetail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridTicketCheckReportList.DataBind();
            }
        }

        protected void sdsTicketCheckReportDetail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridTicketCheckReportList.DataBind();
            }
        }

        protected void sdsTicketCheckReportDetail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridTicketCheckReportList.DataBind();
            }
        }

        /// <summary>
        /// 修改內容存檔前
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void sdsTicketCheckReportDetail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eCaseDate_Temp = (TextBox)fvTicketCheckReportDetail.FindControl("eCaseDate_Edit");
            Label eModifyDate_Temp = (Label)fvTicketCheckReportDetail.FindControl("eModifyDate_Edit");
            DateTime vCaseDate_Temp = (eCaseDate_Temp.Text.Trim() != "") ? DateTime.Parse(eCaseDate_Temp.Text.Trim()) : DateTime.Today;
            DateTime vModifyDate_Temp = (eModifyDate_Temp.Text.Trim() != "") ? DateTime.Parse(eModifyDate_Temp.Text.Trim()) : DateTime.Today;
            Label eCaseNo_Temp = (Label)fvTicketCheckReportDetail.FindControl("eCaseNo_Edit");
            string vCaseNo_Temp = eCaseNo_Temp.Text.Trim();
            if (vCaseNo_Temp != "")
            {
                //新增一筆異動記錄
                string vSQLStr_Temp = "select isnull(MAX(Items), '0') MaxItems from TicketCheckReport_History where CaseNo = '" + vCaseNo_Temp + "' ";
                int vMaxItems = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems")) + 1;
                string vNewCaseNoItems = vCaseNo_Temp + vMaxItems.ToString("D4");
                vSQLStr_Temp = "insert into TicketCheckReport_History (CaseNoItem, Items, ModifyMode, CaseNo, LinesNo, DepNo, Driver, DriverName, " + Environment.NewLine +
                               "                                       CaseDate, DeparturePosition, DepartureTime, ArrivePosition, ArriveTime, " + Environment.NewLine +
                               "                                       Car_ID, CheckNote, CaseNote, Remark, Inspector, BuDate, BuMan, " + Environment.NewLine +
                               "                                       ModifyDate, ModifyMan, IsPassed, RP_ReportNo, HasAssetNo, HasSecurityLabel, " + Environment.NewLine +
                               "                                       StationToilet, StationLight)" + Environment.NewLine +
                               "select '" + vNewCaseNoItems.Trim() + "', '" + vMaxItems.ToString("D4") + "', 'EDIT', CaseNo, LinesNo, DepNo, " + Environment.NewLine +
                               "       Driver, DriverName, CaseDate, DeparturePosition, DepartureTime, ArrivePosition, ArriveTime, Car_ID, " + Environment.NewLine +
                               "       CheckNote, CaseNote, Remark, Inspector, BuDate, BuMan, ModifyDate, ModifyMan, IsPassed, RP_ReportNo, " + Environment.NewLine +
                               "       HasAssetNo, HasSecurityLabel, StationToilet, StationLight " + Environment.NewLine +
                               "  from TicketCheckReport " + Environment.NewLine +
                               " where CaseNo = '" + vCaseNo_Temp + "' ";
                PF.ExecSQL(vConnStr, vSQLStr_Temp);
            }

            e.Command.Parameters["@CaseDate"].Value = vCaseDate_Temp;
            e.Command.Parameters["@ModifyDate"].Value = vModifyDate_Temp;
        }

        /// <summary>
        /// 資料匯出成 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExportExcel_Click(object sender, EventArgs e)
        {
            string vFileName = "查票工作報告";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                string vSelStr = GetSelectStr_Print("");
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
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
                    fontData_Red.Color = IndexedColors.Red.Index; //字體顏色
                    csData_Red.SetFont(fontData_Red);

                    HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                    HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                    csData_Int.DataFormat = format.GetFormat("###,##0");

                    string vHeaderText = "";
                    int vLinesNo = 0;

                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "CASENO") ? "序號" :
                                      (drExcel.GetName(i).ToUpper() == "LINESNO") ? "路線" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "單位代號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "單位" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員代號" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "CASEDATE") ? "查票日期" :
                                      (drExcel.GetName(i).ToUpper() == "DEPARTUREPOSITION") ? "出發地點" :
                                      (drExcel.GetName(i).ToUpper() == "DEPARTURETIME") ? "出發時間" :
                                      (drExcel.GetName(i).ToUpper() == "ARRIVEPOSITION") ? "到達地點" :
                                      (drExcel.GetName(i).ToUpper() == "ARRIVETIME") ? "到達時間" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "車號" :
                                      (drExcel.GetName(i).ToUpper() == "CHECKNOTE") ? "查票經過" :
                                      (drExcel.GetName(i).ToUpper() == "CASENOTE") ? "處理經過" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK") ? "備註" :
                                      (drExcel.GetName(i).ToUpper() == "BUDATE") ? "建檔日期" :
                                      (drExcel.GetName(i).ToUpper() == "BUMAN") ? "建檔人代號" :
                                      (drExcel.GetName(i).ToUpper() == "BUMANNAME") ? "建檔人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYDATE") ? "異動日期" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYMAN") ? "異動人代號" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYMANNAME") ? "異動人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "INSPECTOR") ? "稽查人員代號" :
                                      (drExcel.GetName(i).ToUpper() == "INSPECTORNAME") ? "稽查人員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "HASASSETNO") ? "財產編號" :
                                      (drExcel.GetName(i).ToUpper() == "HASSECURITYLABEL") ? "防偽標籤" :
                                      (drExcel.GetName(i).ToUpper() == "STATIONTOILET") ? "場站廁所" :
                                      (drExcel.GetName(i).ToUpper() == "STATIONLIGHT") ? "場站燈光" :
                                      "";
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
                            if (((drExcel.GetName(i).ToUpper() == "BUDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "CASEDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "MODIFYDATE")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                DateTime vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 1911).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
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
                        var msTarget = new NPOIMemoryStream();
                        msTarget.AllowClose = false;
                        wbExcel.Write(msTarget);
                        msTarget.Flush();
                        msTarget.Seek(0, SeekOrigin.Begin);
                        msTarget.AllowClose = true;

                        if (msTarget.Length > 0)
                        {
                            string vCaseDateStr = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "從 " + eCaseDate_S_Search.Text.Trim() + " 起至 " + eCaseDate_E_Search.Text.Trim() + " 止" :
                                                  ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? eCaseDate_S_Search.Text.Trim() :
                                                  ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? eCaseDate_E_Search.Text.Trim() : "不分日期";
                            string vInspectorStr = (eInspector_Search.Text.Trim() != "") ? eInspector_Search.Text.Trim() : "全部";
                            string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "全部";
                            string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vLinesNoStr = (eLinesNo_Search.Text.Trim() != "") ? eLinesNo_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_查票工作報告" + Environment.NewLine +
                                                 "TicketCheckP.aspx" + Environment.NewLine +
                                                 "查票日期：" + vCaseDateStr + Environment.NewLine +
                                                 "查票人：" + vInspectorStr + Environment.NewLine +
                                                 "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                                 "駕駛員：" + vDriverStr + Environment.NewLine +
                                                 "路線：" + vLinesNoStr;
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