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
    public partial class AnecdoteCase : System.Web.UI.Page
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

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                System.Threading.Thread.CurrentThread.CurrentCulture = Cal;

                //從 SESSION 取回使用者帳號資料
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
                    //有登入成功取得使用者帳號
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
                    }
                    //定義畫面上事件日期欄位各自的 onclick 事件
                    //因為畫面日期的定義是民國年，所以對應的日期選擇視窗是 InputDate_ChineseYears.aspx
                    string vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Start_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Start_Search.Attributes["onClick"] = vCaseDateScript;
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_End_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_End_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        fuExcel.Visible = (vLoginDepNo == "09");
                        bbImportData.Visible = (vLoginDepNo == "09");
                        vStationNo = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vLoginDepNo + "' and InSHReport = 'V'", "DepNo");

                        eDepNo_Start_Search.Text = vStationNo.Trim();
                        eDepNo_End_Search.Text = "";
                        eDepNo_Start_Search.Enabled = (vStationNo == "");
                        eDepNo_End_Search.Enabled = (vStationNo == "");
                    }
                    else
                    {
                        //進行資料繫結
                        CaseDataBind();
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
        /// 取回查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            string vHasInsuStr = (chkHasInsurance.Checked) ? " and isnull(HasInsurance, 0) = 1" + Environment.NewLine : " and isnull(HasInsurance, 0) = 0" + Environment.NewLine;
            string vCaseCloseStr = (chkCaseClose.Checked) ? " and isnull(CaseClose, 0) = 1" + Environment.NewLine : " and isnull(CaseClose, 0) = 0" + Environment.NewLine;
            string vDepNoWhereStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? " and DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                    "";
            string vCaseDateWhereStr = ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() != "")) ?
                                       " and BuildDate between '" + PF.TransDateString(eCaseDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eCaseDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() == "")) ?
                                       " and BuildDate = '" + PF.TransDateString(eCaseDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eCaseDate_Start_Search.Text.Trim() == "") && (eCaseDate_End_Search.Text.Trim() != "")) ?
                                       " and BuildDate = '" + PF.TransDateString(eCaseDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       "";
            string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? " and Car_ID like '%" + eCarID_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vDriverStr = (eDriver_Search.Text.Trim() != "") ? " and Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "SELECT CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, " + Environment.NewLine +
                                "       (SELECT NAME FROM Employee WHERE (EMPNO = AnecdoteCase.BuildMan)) AS BuildManName, " + Environment.NewLine +
                                "       Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, " + Environment.NewLine +
                                "       IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose " + Environment.NewLine +
                                "  FROM AnecdoteCase " + Environment.NewLine +
                                " where isnull(CaseNo, '') <> '' " + Environment.NewLine +
                                vHasInsuStr +
                                vCaseCloseStr +
                                vDepNoWhereStr +
                                vCaseDateWhereStr +
                                vCarIDStr +
                                vDriverStr +
                                " order by CaseNo DESC";
            return vResultStr;
        }

        /// <summary>
        /// 取回EXCEL用的查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr_Excel()
        {
            string vHasInsuStr = (chkHasInsurance.Checked) ? " and HasInsurance = 1" + Environment.NewLine : " and isnull(HasInsurance, 0) = 0" + Environment.NewLine;
            string vCaseCloseStr = (chkCaseClose.Checked) ? " and CaseClose = 1" + Environment.NewLine : " and isnull(CaseClose, 0) = 0" + Environment.NewLine;
            string vDepNoWhereStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? " and DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                    "";
            string vCaseDateWhereStr = ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() != "")) ?
                                       " and BuildDate between '" + PF.TransDateString(eCaseDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eCaseDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() == "")) ?
                                       " and BuildDate = '" + PF.TransDateString(eCaseDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eCaseDate_Start_Search.Text.Trim() == "") && (eCaseDate_End_Search.Text.Trim() != "")) ?
                                       " and BuildDate = '" + PF.TransDateString(eCaseDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       "";
            string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? " and Car_ID like '%" + eCarID_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vDriverStr = (eDriver_Search.Text.Trim() != "") ? " and Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "SELECT CaseNo, " + Environment.NewLine +
                                "       case when HasInsurance = 1 then '是' else '否' end HasInsurance, " + Environment.NewLine +
                                "       DepNo, DepName, BuildDate, BuildMan, " + Environment.NewLine +
                                "       (SELECT NAME FROM Employee WHERE (EMPNO = AnecdoteCase.BuildMan)) AS BuildManName, " + Environment.NewLine +
                                "       Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, " + Environment.NewLine +
                                "       case when IsNoDeduction = 1 then '是' else '否' end IsNoDeduction, " + Environment.NewLine +
                                "       DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose " + Environment.NewLine +
                                "  FROM AnecdoteCase " + Environment.NewLine +
                                " where isnull(CaseNo, '') <> '' " + Environment.NewLine +
                                vHasInsuStr +
                                vCaseCloseStr +
                                vDepNoWhereStr +
                                vCaseDateWhereStr +
                                vCarIDStr +
                                vDriverStr +
                                " order by CaseNo DESC";
            return vResultStr;
        }

        /// <summary>
        /// 資料繫結
        /// </summary>
        private void CaseDataBind()
        {
            if (vLoginID != "")
            {
                sdsAnecdoteCaseA_List.SelectCommand = GetSelectStr();
                gridAnecdoteCaseA_List.DataBind();
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            CaseDataBind();
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 從 EXCEL 匯入資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbImportData_Click(object sender, EventArgs e)
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\"; //取回主機上的應用程式實體資料夾路徑
            string vUploadFileName = "";
            string vExtName = "";
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
            int vColIndex = 0;

            if (fuExcel.FileName != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
                }
                vExtName = Path.GetExtension(fuExcel.FileName); //取得 EXCEL 檔的副檔名，用來判斷 EXCEL 版本
                vUploadFileName = vUploadPath + fuExcel.FileName; //取得EXCEL 檔的完整路徑
                fuExcel.SaveAs(vUploadFileName);
                if (vExtName == ".xls") //舊版 EXCEL (97~2003)
                {
                    HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                    HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                    for (int i = sheetExcel_H.FirstRowNum + 2; i <= sheetExcel_H.LastRowNum; i++) //標題列有二行，所以起始列加 2
                    {
                        HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                        if (vRowTemp_H.GetCell(1).CellType != CellType.Blank)
                        {
                            // 2023.10.19 避免發生 EXCEL 欄位空白導致欄位對應錯亂
                            vDepNo_Temp = "";
                            vDepName_Temp = "";
                            vBuildDate = "";
                            vCaseNo = "";
                            vBuildMan = vLoginID;
                            vCar_ID = "";
                            vDriverName = "";
                            vDriver = "";
                            vInsuMan = vLoginID;
                            vAncedotalResRatio = "";
                            vRemark = "";
                            vDeductionDate = "";
                            vIsNoDeduction = "";
                            for (int iColumnIndex = 0; iColumnIndex < vRowTemp_H.Cells.Count; iColumnIndex++)
                            {
                                vColIndex = vRowTemp_H.Cells[iColumnIndex].ColumnIndex;
                                switch (vColIndex)
                                {
                                    case 1:
                                        vDepNo_Temp = vRowTemp_H.GetCell(vColIndex).ToString().Trim();
                                        vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                                        vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
                                        break;
                                    case 3:
                                        vBuildDate = (DateTime.ParseExact(vRowTemp_H.GetCell(vColIndex).ToString().Trim(), "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd");
                                        vCaseNo = (DateTime.Today.Year - 1911).ToString("D3") + DateTime.Today.Month.ToString("D2") + "A";
                                        vSQLStr = "select MAX(CaseNo) CaseNo from AnecdoteCase where CaseNo like '" + vCaseNo + "%' ";
                                        vTempIndex = (PF.GetValue(vConnStr, vSQLStr, "CaseNo")).Replace(vCaseNo, "").Trim();
                                        vTempIndex = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D4") : "0001";
                                        vCaseNo = vCaseNo + vTempIndex;
                                        break;
                                    case 4:
                                        vCar_ID = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(4).ToString().Trim() : "";
                                        break;
                                    case 5:
                                        vDriverName = vRowTemp_H.GetCell(vColIndex).ToString().Trim();
                                        vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' ";
                                        vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                        break;
                                    case 8:
                                        vInsuMan = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(8).ToString().Trim() : "";
                                        break;
                                    case 18:
                                        vAncedotalResRatio = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 19:
                                        if (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank)
                                        {
                                            vTempStr = vRowTemp_H.GetCell(vColIndex).ToString().Trim();
                                            vIsNoDeduction = (vTempStr == "免扣") ? "true" : "false";
                                            vTempStr = vTempStr.Replace("免扣", "");
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
                                        break;
                                    case 20:
                                        if (vRemark != "")
                                        {
                                            vRemark = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRemark + Environment.NewLine + vRowTemp_H.GetCell(vColIndex).ToString().Trim() : vRemark;
                                        }
                                        else
                                        {
                                            vRemark = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(20).ToString().Trim() : "";
                                        }
                                        break;
                                }
                            }
                            /*
                            vDepNo_Temp = vRowTemp_H.GetCell(1).ToString().Trim();
                            vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
                            vBuildDate = (DateTime.ParseExact(vRowTemp_H.GetCell(3).ToString().Trim(), "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd");
                            //2023.08.18 CaseNo 改用建單當下的年月取單號，而不是用發生日期取單號
                            //vCaseNo = (DateTime.Parse(vBuildDate).Year - 1911).ToString("D3") + DateTime.Parse(vBuildDate).Month.ToString("D2") + "A";
                            vCaseNo = (DateTime.Today.Year - 1911).ToString("D3") + DateTime.Today.Month.ToString("D2") + "A";
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
                            } //*/

                            using (SqlDataSource dsTemp = new SqlDataSource())
                            {
                                dsTemp.ConnectionString = (vConnStr != "") ? vConnStr : PF.GetConnectionStr(Request.ApplicationPath);
                                dsTemp.InsertCommand = "insert into AnecdoteCase " + Environment.NewLine +
                                                       "       (CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, DriverName, " + Environment.NewLine +
                                                       "        InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark)" + Environment.NewLine +
                                                       "values (@CaseNo, @HasInsurance, @DepNo, @DepName, @BuildDate, @BuildMan, @Car_ID, @Driver, @DriverName, " + Environment.NewLine +
                                                       "        @InsuMan, @AnecdotalResRatio, @IsNoDeduction, @DeductionDate, @Remark)";
                                dsTemp.InsertParameters.Clear(); //先清空參數
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
                                dsTemp.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark != "") ? vRemark : String.Empty));
                                dsTemp.Insert();
                            }
                        }

                        if (vRowTemp_H.GetCell(6).CellType != CellType.Blank)
                        {
                            vSQLStr = "select MAX(Items) MaxItem from AnecdoteCaseB where CaseNo = '" + vCaseNo + "' ";
                            vTempIndex = PF.GetValue(vConnStr, vSQLStr, "MaxItem");
                            vItems = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D4") : "0001";
                            vCaseNoItems = vCaseNo + vItems;
                            // 2023.10.19 避免發生 EXCEL 欄位空白導致欄位對應錯亂
                            vRelationship = "";
                            vRelCar_ID = "";
                            vRemarkB = "";
                            vEstimatedAmount = "0";
                            vThirdInsurance = "0";
                            vCompInsurance = "0";
                            vDriverSharing = "0";
                            vCompanySharing = "0";
                            vCarDamageAMT = "0";
                            vPersonDamageAMT = "0";
                            vRelationComp = "0";
                            vPassengerInsu = "0";
                            vReconciliationDate = "";
                            for (int iColumnIndex = 0; iColumnIndex < vRowTemp_H.Cells.Count; iColumnIndex++)
                            {
                                vColIndex = vRowTemp_H.Cells[iColumnIndex].ColumnIndex;
                                switch (vColIndex)
                                {
                                    case 6:
                                        vRelationship = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim() : "";
                                        break;
                                    case 7:
                                        vRelCar_ID = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim() : "";
                                        break;
                                    case 9:
                                        vEstimatedAmount = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
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
                                        break;
                                    case 10:
                                        vThirdInsurance = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 11:
                                        vCompInsurance = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 12:
                                        vDriverSharing = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 13:
                                        vCompanySharing = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 14:
                                        vCarDamageAMT = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 15:
                                        vPersonDamageAMT = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 16:
                                        vRelationComp = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 17:
                                        vTempStr = (vRowTemp_H.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_H.GetCell(vColIndex).ToString().Trim() : "";
                                        vTempStr = vTempStr.Replace("和解", "");
                                        vTempStr = vTempStr.Replace("自行", "");
                                        vPassengerInsu = "0"; //乘客險先用 0 轉入
                                        vReconciliationDate = (vTempStr != "") ? (DateTime.ParseExact(vTempStr, "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd") : vTempStr;
                                        break;
                                }
                            }
                            /*
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
                            //*/ 

                            using (SqlDataSource dsTempB = new SqlDataSource())
                            {
                                dsTempB.ConnectionString = (vConnStr != "") ? vConnStr : PF.GetConnectionStr(Request.ApplicationPath);
                                dsTempB.InsertCommand = "INSERT INTO AnecdoteCaseB " + Environment.NewLine +
                                                        "      (CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, " + Environment.NewLine +
                                                        "       DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, " + Environment.NewLine +
                                                        "       PassengerInsu, Remark) " + Environment.NewLine +
                                                        "values(@CaseNo, @Items, @CaseNoItems, @Relationship, @RelCar_ID, @EstimatedAmount, @ThirdInsurance, @CompInsurance, " + Environment.NewLine +
                                                        "       @DriverSharing, @CompanySharing, @CarDamageAMT, @PersonDamageAMT, @RelationComp, @ReconciliationDate, " + Environment.NewLine +
                                                        "       @PassengerInsu, @Remark)";
                                dsTempB.InsertParameters.Clear(); //先清空參數
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
                                dsTempB.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemarkB != "") ? vRemarkB : String.Empty));
                                dsTempB.InsertParameters.Add(new Parameter("PassengerInsu", DbType.Int32, vPassengerInsu));
                                dsTempB.Insert();
                            }
                        }
                    }
                }
                else if (vExtName == ".xlsx") //2010 之後的 EXCEL
                {
                    XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                    XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                    for (int i = sheetExcel_X.FirstRowNum + 2; i <= sheetExcel_X.LastRowNum; i++) //標題列有二行，所以起始列 +2
                    {
                        XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                        if (vRowTemp_X.GetCell(1).CellType != CellType.Blank)
                        {
                            // 2023.10.19 避免發生 EXCEL 欄位空白導致欄位對應錯亂
                            vDepNo_Temp = "";
                            vDepName_Temp = "";
                            vBuildDate = "";
                            vCaseNo = "";
                            vBuildMan = vLoginID;
                            vCar_ID = "";
                            vDriverName = "";
                            vDriver = "";
                            vInsuMan = vLoginID;
                            vAncedotalResRatio = "";
                            vRemark = "";
                            vDeductionDate = "";
                            vIsNoDeduction = "";
                            for (int iColumnIndex = 0; iColumnIndex < vRowTemp_X.Cells.Count; iColumnIndex++)
                            {
                                vColIndex = vRowTemp_X.Cells[iColumnIndex].ColumnIndex;
                                switch (vColIndex)
                                {
                                    case 1:
                                        vDepNo_Temp = vRowTemp_X.GetCell(vColIndex).ToString().Trim();
                                        vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                                        vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
                                        break;
                                    case 3:
                                        vBuildDate = (DateTime.ParseExact(vRowTemp_X.GetCell(vColIndex).ToString().Trim(), "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd");
                                        vCaseNo = (DateTime.Today.Year - 1911).ToString("D3") + DateTime.Today.Month.ToString("D2") + "A";
                                        vSQLStr = "select MAX(CaseNo) CaseNo from AnecdoteCase where CaseNo like '" + vCaseNo + "%' ";
                                        vTempIndex = (PF.GetValue(vConnStr, vSQLStr, "CaseNo")).Replace(vCaseNo, "").Trim();
                                        vTempIndex = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D4") : "0001";
                                        vCaseNo = vCaseNo + vTempIndex;
                                        break;
                                    case 4:
                                        vCar_ID = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(4).ToString().Trim() : "";
                                        break;
                                    case 5:
                                        vDriverName = vRowTemp_X.GetCell(vColIndex).ToString().Trim();
                                        vSQLStr = "select EmpNo from Employee where [Name] = '" + vDriverName + "' ";
                                        vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                                        break;
                                    case 8:
                                        vInsuMan = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(8).ToString().Trim() : "";
                                        break;
                                    case 18:
                                        vAncedotalResRatio = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 19:
                                        if (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank)
                                        {
                                            vTempStr = vRowTemp_X.GetCell(vColIndex).ToString().Trim();
                                            vIsNoDeduction = (vTempStr == "免扣") ? "true" : "false";
                                            vTempStr = vTempStr.Replace("免扣", "");
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
                                        break;
                                    case 20:
                                        if (vRemark != "")
                                        {
                                            vRemark = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRemark + Environment.NewLine + vRowTemp_X.GetCell(vColIndex).ToString().Trim() : vRemark;
                                        }
                                        else
                                        {
                                            vRemark = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(20).ToString().Trim() : "";
                                        }
                                        break;
                                }
                            }
                            /*
                            vDepNo_Temp = vRowTemp_X.GetCell(1).ToString().Trim();
                            vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
                            vBuildDate = (DateTime.ParseExact(vRowTemp_X.GetCell(3).ToString().Trim(), "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd");
                            //2023.08.18 CaseNo 改用建單當下的年月取單號，而不是用發生日期取單號
                            //vCaseNo = (DateTime.Parse(vBuildDate).Year - 1911).ToString("D3") + DateTime.Parse(vBuildDate).Month.ToString("D2") + "A";
                            vCaseNo = (DateTime.Today.Year - 1911).ToString("D3") + DateTime.Today.Month.ToString("D2") + "A";
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
                            } //*/

                            using (SqlDataSource dsTemp = new SqlDataSource())
                            {
                                dsTemp.ConnectionString = (vConnStr != "") ? vConnStr : PF.GetConnectionStr(Request.ApplicationPath);
                                dsTemp.InsertCommand = "insert into AnecdoteCase " + Environment.NewLine +
                                                       "       (CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, DriverName, " + Environment.NewLine +
                                                       "        InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark)" + Environment.NewLine +
                                                       "values (@CaseNo, @HasInsurance, @DepNo, @DepName, @BuildDate, @BuildMan, @Car_ID, @Driver, @DriverName, " + Environment.NewLine +
                                                       "        @InsuMan, @AnecdotalResRatio, @IsNoDeduction, @DeductionDate, @Remark)";
                                dsTemp.InsertParameters.Clear(); //先清空參數
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
                                dsTemp.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark != "") ? vRemark : String.Empty));
                                dsTemp.Insert();
                            }
                        }

                        if (vRowTemp_X.GetCell(6).CellType != CellType.Blank)
                        {
                            vSQLStr = "select MAX(Items) MaxItem from AnecdoteCaseB where CaseNo = '" + vCaseNo + "' ";
                            vTempIndex = PF.GetValue(vConnStr, vSQLStr, "MaxItem");
                            vItems = (vTempIndex != "") ? (int.Parse(vTempIndex) + 1).ToString("D4") : "0001";
                            vCaseNoItems = vCaseNo + vItems;
                            // 2023.10.19 避免發生 EXCEL 欄位空白導致欄位對應錯亂
                            vRelationship = "";
                            vRelCar_ID = "";
                            vRemarkB = "";
                            vEstimatedAmount = "0";
                            vThirdInsurance = "0";
                            vCompInsurance = "0";
                            vDriverSharing = "0";
                            vCompanySharing = "0";
                            vCarDamageAMT = "0";
                            vPersonDamageAMT = "0";
                            vRelationComp = "0";
                            vPassengerInsu = "0";
                            vReconciliationDate = "";
                            for (int iColumnIndex = 0; iColumnIndex < vRowTemp_X.Cells.Count; iColumnIndex++)
                            {
                                vColIndex = vRowTemp_X.Cells[iColumnIndex].ColumnIndex;
                                switch (vColIndex)
                                {
                                    case 6:
                                        vRelationship = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim() : "";
                                        break;
                                    case 7:
                                        vRelCar_ID = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim() : "";
                                        break;
                                    case 9:
                                        vEstimatedAmount = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
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
                                        break;
                                    case 10:
                                        vThirdInsurance = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 11:
                                        vCompInsurance = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 12:
                                        vDriverSharing = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 13:
                                        vCompanySharing = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 14:
                                        vCarDamageAMT = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 15:
                                        vPersonDamageAMT = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 16:
                                        vRelationComp = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim().Replace(",", "") : "0";
                                        break;
                                    case 17:
                                        vTempStr = (vRowTemp_X.GetCell(vColIndex).CellType != CellType.Blank) ? vRowTemp_X.GetCell(vColIndex).ToString().Trim() : "";
                                        vTempStr = vTempStr.Replace("和解", "");
                                        vTempStr = vTempStr.Replace("自行", "");
                                        vPassengerInsu = "0"; //乘客險先用 0 轉入
                                        vReconciliationDate = (vTempStr != "") ? (DateTime.ParseExact(vTempStr, "yyy.MM.dd", CultureInfo.InvariantCulture).AddYears(1911)).ToString("yyyy/MM/dd") : vTempStr;
                                        break;
                                }
                            }
                            /*
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
                            //*/

                            using (SqlDataSource dsTempB = new SqlDataSource())
                            {
                                dsTempB.ConnectionString = (vConnStr != "") ? vConnStr : PF.GetConnectionStr(Request.ApplicationPath);
                                dsTempB.InsertCommand = "INSERT INTO AnecdoteCaseB " + Environment.NewLine +
                                                        "      (CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, " + Environment.NewLine +
                                                        "       DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, " + Environment.NewLine +
                                                        "       PassengerInsu, Remark) " + Environment.NewLine +
                                                        "values(@CaseNo, @Items, @CaseNoItems, @Relationship, @RelCar_ID, @EstimatedAmount, @ThirdInsurance, @CompInsurance, " + Environment.NewLine +
                                                        "       @DriverSharing, @CompanySharing, @CarDamageAMT, @PersonDamageAMT, @RelationComp, @ReconciliationDate, " + Environment.NewLine +
                                                        "       @PassengerInsu, @Remark)";
                                dsTempB.InsertParameters.Clear(); //先清空參數
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
                                dsTempB.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemarkB != "") ? vRemarkB : String.Empty));
                                dsTempB.InsertParameters.Add(new Parameter("PassengerInsu", DbType.Int32, vPassengerInsu));
                                dsTempB.Insert();
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 主檔資料更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            Label eCaseNo_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_Edit");
            if (eCaseNo_Edit != null)
            {
                string vCaseNo_Temp = eCaseNo_Edit.Text.Trim(); //先取回單號

                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
                }
                try
                {
                    //複製一份到異動檔
                    DateTime vModifyDate = DateTime.Today;
                    string vHistoryNo = vModifyDate.Year.ToString("D4") + vModifyDate.Month.ToString("D2") + "MODA";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%'";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndex_Str = (vMaxNo.Trim() != "") ? vMaxNo.Replace(vHistoryNo, "").Trim() : "0";
                    int vNewIndex = Int32.Parse(vIndex_Str) + 1;
                    string vNewHistoryNo = vHistoryNo + vNewIndex.ToString("D4");
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseHistory] " + Environment.NewLine +
                                   "            ([HistoryNo],[CaseNo],[HasInsurance],[DepNo],[DepName],[BuildDate],[BuildMan],[Car_ID],[Driver], " + Environment.NewLine +
                                   "             [DriverName],[InsuMan],[AnecdotalResRatio],[IsNoDeduction],[DeductionDate],[Remark]," + Environment.NewLine +
                                   "             [ModifyType],[ModifyDate],[ModifyMan],[CaseOccurrence],[ERPCouseNo],[CaseClose]) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "',[CaseNo],[HasInsurance],[DepNo],[DepName],[BuildDate],[BuildMan],[Car_ID],[Driver], " + Environment.NewLine +
                                   "       [DriverName],[InsuMan],[AnecdotalResRatio],[IsNoDeduction],[DeductionDate],[Remark]," + Environment.NewLine +
                                   "       'EDIT',GetDate(),'" + vLoginID + "',[CaseOccurrence],[ERPCouseNo],[CaseClose] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCase] " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);

                    //開始真正的進行更新
                    Label eHasInsuranceTitle_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsuranceTitle_Edit");
                    TextBox eDepNo_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_Edit");
                    Label eDepName_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("edepName_Edit");
                    TextBox eCar_ID_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCarID_Edit");
                    TextBox eDriver_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_Edit");
                    Label eDriverName_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_Edit");
                    TextBox eInsuMan_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eInsuMan_Edit");
                    TextBox eAnecdotalResRatio_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAnecdotalResRatio_Edit");
                    Label eIsNoDeductionTitle_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeductionTitle_Edit");
                    TextBox eDeductionDate_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDeductionDate_Edit");
                    TextBox eRemark_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eRemark_Edit");
                    TextBox eBuildDate_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildDate_Edit");
                    TextBox eCaseOccurrence_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseOccurrence_Edit");
                    Label eCaseCloseTitle_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseCloseTitle_Edit");

                    string vHasInsurance_Temp = (eHasInsuranceTitle_Edit.Text.Trim() != "") ? eHasInsuranceTitle_Edit.Text.Trim() : "False";
                    string vDepNo_Temp = eDepNo_Edit.Text.Trim();
                    string vDepName_Temp = eDepName_Edit.Text.Trim();
                    string vCarID_Temp = eCar_ID_Edit.Text.Trim().ToUpper();
                    string vDriver_Temp = eDriver_Edit.Text.Trim();
                    string vDriverName_Temp = eDriverName_Edit.Text.Trim();
                    string vInsuMan_Temp = eInsuMan_Edit.Text.Trim();
                    string vAnecdotalResRatio_Temp = eAnecdotalResRatio_Edit.Text.Trim();
                    string vIsNoDeduction_Temp = (eIsNoDeductionTitle_Edit.Text.Trim() != "") ? eIsNoDeductionTitle_Edit.Text.Trim() : "False";
                    string vDeductionDate_Temp = (eDeductionDate_Edit.Text.Trim() != "") ? DateTime.Parse(eDeductionDate_Edit.Text.Trim()).ToShortDateString() : "";
                    string vRemark_Temp = eRemark_Edit.Text.Trim();
                    string vBuildDate_Temp = (eBuildDate_Edit.Text.Trim() != "") ? DateTime.Parse(eBuildDate_Edit.Text.Trim()).ToShortDateString() : "";
                    string vCaseOccurrence_Temp = eCaseOccurrence_Edit.Text.Trim();
                    string vCaseClose_Temp = (eCaseCloseTitle_Edit.Text.Trim() != "") ? eCaseCloseTitle_Edit.Text.Trim() : "False";

                    sdsAnecdoteCaseA_Data.UpdateCommand = "UPDATE AnecdoteCase " + Environment.NewLine +
                                                          "   SET HasInsurance = @HasInsurance, DepNo = @DepNo, DepName = @DepName, Car_ID = @Car_ID, Driver = @Driver, " + Environment.NewLine +
                                                          "       DriverName = @DriverName, InsuMan = @InsuMan, AnecdotalResRatio = @AnecdotalResRatio, IsNoDeduction = @IsNoDeduction, " + Environment.NewLine +
                                                          "       DeductionDate = @DeductionDate, Remark = @Remark, BuildDate = @BuildDate, CaseOccurrence = @CaseOccurrence, CaseClose = @CaseClose " + Environment.NewLine +
                                                          " WHERE (CaseNo = @CaseNo)";
                    sdsAnecdoteCaseA_Data.UpdateParameters.Clear();
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("HasInsurance", DbType.Boolean, vHasInsurance_Temp));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_Temp));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("DepName", DbType.String, vDepName_Temp));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Car_ID", DbType.String, vCarID_Temp));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Driver", DbType.String, vDriver_Temp));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("DriverName", DbType.String, vDriverName_Temp));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("InsuMan", DbType.String, vInsuMan_Temp));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("AnecdotalResRatio", DbType.Double, (vAnecdotalResRatio_Temp != "") ? vAnecdotalResRatio_Temp : string.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("IsNoDeduction", DbType.Boolean, vIsNoDeduction_Temp));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("DeductionDate", DbType.Date, (vDeductionDate_Temp != "") ? vDeductionDate_Temp : string.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Remark", DbType.String, (vRemark_Temp != "") ? vRemark_Temp : string.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("BuildDate", DbType.Date, (vBuildDate_Temp != "") ? vBuildDate_Temp : string.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("CaseOccurrence", DbType.String, (vCaseOccurrence_Temp != "") ? vCaseOccurrence_Temp : string.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("CaseClose", DbType.Boolean, vCaseClose_Temp));

                    sdsAnecdoteCaseA_Data.Update();
                    fvAnecdoteCaseA_Data.ChangeMode(FormViewMode.ReadOnly);
                    gridAnecdoteCaseA_List.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_Edit");
            if (eDepNo_Edit != null)
            {
                Label eDepName_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("eDepName_Edit");
                string vDepNo_Temp = eDepNo_Edit.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
                string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName_Temp == "")
                {
                    vDepName_Temp = vDepNo_Temp.Trim();
                    vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                    vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_Edit.Text = vDepNo_Temp.Trim();
                eDepName_Edit.Text = vDepName_Temp.Trim();
            }
        }

        protected void eDriver_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_Edit");
            if (eDriver_Edit != null)
            {
                Label eDriverName_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_Edit");
                string vDriver_Temp = eDriver_Edit.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp.Trim() + "' ";
                string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDriverName_Temp == "")
                {
                    vDriverName_Temp = vDriver_Temp.Trim();
                    vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' ";
                    vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eDriver_Edit.Text = vDriver_Temp.Trim();
                eDriverName_Edit.Text = vDriverName_Temp.Trim();
            }
        }

        /// <summary>
        /// 變更 "出險" 選項
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eHasInsurance_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox eHasInsurance_Edit = (CheckBox)fvAnecdoteCaseA_Data.FindControl("eHasInsurance_Edit");
            if (eHasInsurance_Edit != null)
            {
                Label eHasInsuranceTitle_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsuranceTitle_Edit");
                eHasInsuranceTitle_Edit.Text = (eHasInsurance_Edit.Checked == true) ? "True" : "False";
            }
        }

        /// <summary>
        /// 變更 "不扣精勤" 選項
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eIsNoDeduction_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox eIsNoDeduction_Edit = (CheckBox)fvAnecdoteCaseA_Data.FindControl("eIsNoDeduction_Edit");
            if (eIsNoDeduction_Edit != null)
            {
                Label eIsNoDeductionTitle_Edit = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeductionTitle_Edit");
                eIsNoDeductionTitle_Edit.Text = (eIsNoDeduction_Edit.Checked == true) ? "True" : "False";
            }
        }

        /// <summary>
        /// 新增主檔資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eBuildDate_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildDate_INS");
            TextBox eDepNo_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_INS");
            try
            {
                if ((eBuildDate_INS != null) && ((eDepNo_INS != null) && (eDepNo_INS.Text.Trim() != "")))
                {
                    DateTime vBuildDate = DateTime.Parse(eBuildDate_INS.Text.Trim());
                    //2023.08.18 CaseNo 改用建單日期取單號，而不是用發生日期取單號
                    //string vCaseNo_FirstCode = (vBuildDate.Year - 1911).ToString() + vBuildDate.Month.ToString("D2") + "A";
                    string vCaseNo_FirstCode = (DateTime.Today.Year - 1911).ToString() + DateTime.Today.Month.ToString("D2") + "A";
                    string vSQLStr_Temp = "select max(CaseNo) MaxNo from AnecdoteCase where CaseNo like '" + vCaseNo_FirstCode + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndexStr = (vMaxNo.Trim() != "") ? vMaxNo.Replace(vCaseNo_FirstCode, "").Trim() : "0";
                    int vIndex = Int32.Parse(vIndexStr) + 1;
                    string vCaseNo_Temp = vCaseNo_FirstCode + vIndex.ToString("D4");
                    string vBuildDateStr_Temp = (eBuildDate_INS.Text != "") ? vBuildDate.ToShortDateString() : "";
                    Label eHasInsuranceTitle_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsuranceTitle_INS");
                    string vHasInsurance_Temp = (eHasInsuranceTitle_INS.Text.Trim() != "") ? eHasInsuranceTitle_INS.Text.Trim() : "False";
                    Label eDepName_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eDepName_INS");
                    string vDepNo_Temp = eDepNo_INS.Text.Trim();
                    string vDepName_Temp = eDepName_INS.Text.Trim();
                    TextBox eCarID_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCarID_INS");
                    string vCarID_Temp = eCarID_INS.Text.Trim().ToUpper();
                    TextBox eDriver_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_INS");
                    Label eDriverName_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_INS");
                    string vDriver_Temp = eDriver_INS.Text.Trim();
                    string vDriverName_Temp = eDriverName_INS.Text.Trim();
                    TextBox eInsuMan_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eInsuMan_INS");
                    string vInsuMan_Temp = eInsuMan_INS.Text.Trim();
                    TextBox eAnecdotalResRatio_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAnecdotalResRatio_INS");
                    string vAnecdotalResRatio_Temp = eAnecdotalResRatio_INS.Text.Trim();
                    Label eIsNoDeductionTitle_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeductionTitle_INS");
                    string vIsNoDeduction_Temp = (eIsNoDeductionTitle_INS.Text.Trim() != "") ? eIsNoDeductionTitle_INS.Text.Trim() : "False";
                    TextBox eDeductionDate_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDeductionDate_INS");
                    DateTime vDeductionDate = (eDeductionDate_INS.Text.Trim() != "") ? DateTime.Parse(eDeductionDate_INS.Text.Trim()) : DateTime.Today;
                    string vDeductionDateStr_Temp = (eDeductionDate_INS.Text.Trim() != "") ? vDeductionDate.ToShortDateString() : "";
                    TextBox eRemark_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eRemark_INS");
                    string vRemark_Temp = eRemark_INS.Text.Trim();
                    TextBox eCaseOccurrence_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseOccurrence_INS");
                    string vCaseOccurrence_Temp = eCaseOccurrence_INS.Text.Trim();
                    Label eCaseCloseTitle_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseCloseTitle_INS");
                    string vCaseClose_Temp = (eCaseCloseTitle_INS.Text.Trim() != "") ? eCaseCloseTitle_INS.Text.Trim() : "False";

                    sdsAnecdoteCaseA_Data.InsertCommand = "INSERT INTO AnecdoteCase " + Environment.NewLine +
                                                          "       (CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, DriverName, " + Environment.NewLine +
                                                          "        InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, CaseOccurrence, CaseClose) " + Environment.NewLine +
                                                          "VALUES (@CaseNo, @HasInsurance, @DepNo, @DepName, @BuildDate, @BuildMan, @Car_ID, @Driver, @DriverName, " + Environment.NewLine +
                                                          "        @InsuMan, @AnecdotalResRatio, @IsNoDeduction, @DeductionDate, @Remark, @CaseOccurrence, @CaseClose)";
                    sdsAnecdoteCaseA_Data.InsertParameters.Clear();
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("HasInsurance", DbType.Boolean, vHasInsurance_Temp));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("DepNo", DbType.String, (vDepNo_Temp != "") ? vDepNo_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("DepName", DbType.String, (vDepName_Temp != "") ? vDepName_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, (vBuildDateStr_Temp != "") ? vBuildDateStr_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vLoginID));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Car_ID", DbType.String, (vCarID_Temp != "") ? vCarID_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Driver", DbType.String, (vDriver_Temp != "") ? vDriver_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("DriverName", DbType.String, (vDriverName_Temp != "") ? vDriverName_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("InsuMan", DbType.String, (vInsuMan_Temp != "") ? vInsuMan_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("AnecdotalResRatio", DbType.Double, (vAnecdotalResRatio_Temp != "") ? vAnecdotalResRatio_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("IsNoDeduction", DbType.Boolean, (vIsNoDeduction_Temp != "") ? vIsNoDeduction_Temp : "False"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("DeductionDate", DbType.Date, (vDeductionDateStr_Temp != "") ? vDeductionDateStr_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark_Temp != "") ? vRemark_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("CaseOccurrence", DbType.String, (vCaseOccurrence_Temp != "") ? vCaseOccurrence_Temp : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("CaseClose", DbType.Boolean, (vCaseClose_Temp != "") ? vCaseClose_Temp : "False"));
                    sdsAnecdoteCaseA_Data.Insert();
                    fvAnecdoteCaseA_Data.ChangeMode(FormViewMode.ReadOnly);
                    gridAnecdoteCaseA_List.DataBind();
                    fvAnecdoteCaseA_Data.DataBind();
                }
                else if ((eDepNo_INS != null) && (eDepNo_INS.Text.Trim() == ""))
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('站別不可空白！')");
                    Response.Write("</" + "Script>");
                    eDepNo_INS.Focus();
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

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eDepName_INS");
            if (eDepNo_INS != null)
            {
                string vDepNo_Temp = eDepNo_INS.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
                string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName_Temp == "")
                {
                    vDepName_Temp = vDepNo_Temp.Trim();
                    vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                    vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_INS.Text = vDepNo_Temp.Trim();
                eDepName_INS.Text = vDepName_Temp.Trim();
            }
        }

        protected void eDriver_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_INS");
            Label eDriverName_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_INS");
            if (eDriver_INS != null)
            {
                string vDriver_Temp = eDriver_INS.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp.Trim() + "' ";
                string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDriverName_Temp == "")
                {
                    vDriverName_Temp = vDriver_Temp.Trim();
                    vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' ";
                    vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eDriver_INS.Text = vDriver_Temp.Trim();
                eDriverName_INS.Text = vDriverName_Temp.Trim();
            }
        }

        /// <summary>
        /// 變更 "出險" 選項
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eHasInsurance_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox eHasInsurance_INS = (CheckBox)fvAnecdoteCaseA_Data.FindControl("eHasInsurance_INS");
            Label eHasInsuranceTitle_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsuranceTitle_INS");
            eHasInsuranceTitle_INS.Text = (eHasInsurance_INS.Checked == true) ? "True" : "False";
        }

        /// <summary>
        /// 變更 "不扣精勤" 選項
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eIsNoDeduction_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox eIsNoDeduction_INS = (CheckBox)fvAnecdoteCaseA_Data.FindControl("eIsNoDeduction_INS");
            Label eIsNoDeductionTitle_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeductionTitle_INS");
            eIsNoDeductionTitle_INS.Text = (eIsNoDeduction_INS.Checked == true) ? "True" : "False";
        }

        /// <summary>
        /// 主檔刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_Edit_Click(object sender, EventArgs e)
        {
            Label eCaseNo_List = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_List");
            if (eCaseNo_List != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    //取回要刪除的肇事單號
                    string vCaseNo_Temp = eCaseNo_List.Text.Trim();
                    //因為是刪主檔，所以要連 B 檔和 C 檔都一起刪，可是 C 檔沒有異動檔所以只針對 B 檔做異動
                    DateTime vDelDate = DateTime.Today;
                    string vHistoryNo = vDelDate.Year.ToString("D4") + vDelDate.Month.ToString("D2") + "DELA";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    //string vIndex_Str = (vMaxNo.Trim() != "") ? vMaxNo.Replace(vHistoryNo, "").Trim() : "0";
                    string vIndex_Str = vMaxNo.Replace(vHistoryNo, "").Trim();
                    int vIndex = (vIndex_Str != "") ? Int32.Parse(vIndex_Str) + 1 : 1;
                    string vNewHistoryNo = vHistoryNo + vIndex.ToString("D4");
                    vSQLStr_Temp = "select max(ItemsH) MaxItem from AnecdoteCaseBHistory where HistoryNo = '" + vNewHistoryNo + "' ";
                    string vMaxItemsH = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    string vNewItemsH = (vMaxItemsH.Trim() != "") ? Int32.Parse(vMaxItemsH).ToString("D4") : "0001";
                    string vNewHistoryNoItems = vNewHistoryNo.Trim() + vNewItemsH.Trim();
                    //複製 B 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseBHistory] " + Environment.NewLine +
                                   "            ([HistoryNo],[ItemsH],[HistoryNoItems],[CaseNo],[Items],[CaseNoItems],[Relationship],[RelCar_ID], " + Environment.NewLine +
                                   "             [EstimatedAmount],[ThirdInsurance],[CompInsurance],[DriverSharing],[CompanySharing],[CarDamageAMT], " + Environment.NewLine +
                                   "             [PersonDamageAMT],[RelationComp],[ReconciliationDate],[Remark]," + Environment.NewLine +
                                   "             [ModifyType],[ModifyDate],[ModifyMan],[PassengerInsu])" + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "','" + vNewItemsH + "','" + vNewHistoryNoItems + "',[CaseNo],[Items],[CaseNoItems],[Relationship], " + Environment.NewLine +
                                   "       [RelCar_ID],[EstimatedAmount],[ThirdInsurance],[CompInsurance],[DriverSharing],[CompanySharing], " + Environment.NewLine +
                                   "       [CarDamageAMT],[PersonDamageAMT],[RelationComp],[ReconciliationDate],[Remark], " + Environment.NewLine +
                                   "       'DELA',GetDate(),'" + vLoginID + "',[PassengerInsu] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCaseB] " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //複製 A 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseHistory] " + Environment.NewLine +
                                   "            ([HistoryNo],[CaseNo],[HasInsurance],[DepNo],[DepName],[BuildDate],[BuildMan],[Car_ID],[Driver], " + Environment.NewLine +
                                   "             [DriverName],[InsuMan],[AnecdotalResRatio],[IsNoDeduction],[DeductionDate],[Remark]," + Environment.NewLine +
                                   "             [ModifyType],[ModifyDate],[ModifyMan],[CaseOccurrence],[ERPCouseNo],[CaseClose]) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "',[CaseNo],[HasInsurance],[DepNo],[DepName],[BuildDate],[BuildMan],[Car_ID],[Driver], " + Environment.NewLine +
                                   "       [DriverName],[InsuMan],[AnecdotalResRatio],[IsNoDeduction],[DeductionDate],[Remark]," + Environment.NewLine +
                                   "       'DELA',GetDate(),'" + vLoginID + "',[CaseOccurrence],[ERPCouseNo],[CaseClose] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCase] " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //開始刪除
                    using (SqlDataSource dsDelTemp = new SqlDataSource())
                    {
                        dsDelTemp.ConnectionString = vConnStr;
                        //刪除 C 檔
                        dsDelTemp.DeleteCommand = "delete AnecdoteCaseC where CaseNo = @CaseNo ";
                        dsDelTemp.DeleteParameters.Clear();
                        dsDelTemp.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                        dsDelTemp.Delete();
                        //刪除 B 檔
                        dsDelTemp.DeleteCommand = "delete AnecdoteCaseB where CaseNo = @CaseNo ";
                        dsDelTemp.DeleteParameters.Clear();
                        dsDelTemp.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                        dsDelTemp.Delete();
                        //刪除 A 檔
                        dsDelTemp.DeleteCommand = "delete AnecdoteCase where CaseNo = @CaseNo ";
                        dsDelTemp.DeleteParameters.Clear();
                        dsDelTemp.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                        dsDelTemp.Delete();
                    }
                    fvAnecdoteCaseA_Data.ChangeMode(FormViewMode.ReadOnly);
                    gridAnecdoteCaseA_List.DataBind();
                    fvAnecdoteCaseA_Data.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// FormView [fvAnecdoteCaseA_Data] 已經完成資料繫結
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void fvAnecdoteCaseA_Data_DataBound(object sender, EventArgs e)
        {
            bool vAllowDep = ((vLoginID.ToUpper() == "SUPERVISOR") || (vLoginDepNo == "09") || (vLoginDepNo == "06"));
            switch (fvAnecdoteCaseA_Data.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    //plSearch.Visible = true;
                    plDetailDataShow.Visible = ((gridAnecdoteCaseA_List.Rows.Count > 0) && (gridAnecdoteCaseA_List.SelectedIndex >= 0));
                    Button bbExportERP_Temp = (Button)fvAnecdoteCaseA_Data.FindControl("bbExportERP_List");
                    if (bbExportERP_Temp != null)
                    {
                        Label eERPCouseNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eERPCouseNo_List");
                        Button bbDelERP_Temp = (Button)fvAnecdoteCaseA_Data.FindControl("bbDelERP_List");
                        bbExportERP_Temp.Visible = ((vAllowDep) && (eERPCouseNo_Temp.Text.Trim() == ""));
                        bbDelERP_Temp.Visible = ((vAllowDep) && (eERPCouseNo_Temp.Text.Trim() != ""));
                    }
                    /* 以下是用來限制各車站人員登入之後的權限用的...現在都先開放不限制
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vStationNo_FV = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vDepNo + "' and InSHReport = 'V'", "DepNo");
                    Button bbInsert_List = (Button)fvAnecdoteCaseA_Data.FindControl("bbNew_List");
                    Button bbInsert_Empty = (Button)fvAnecdoteCaseA_Data.FindControl("bbNew_Empty");
                    Button bbEdit_List = (Button)fvAnecdoteCaseA_Data.FindControl("bbEdit_List");
                    Button bbDelete_List = (Button)fvAnecdoteCaseA_Data.FindControl("bbDel_List");

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
                    } //*/
                    break;
                case FormViewMode.Edit:
                    //plSearch.Visible = false;
                    plDetailDataShow.Visible = false;
                    TextBox eDeductionDate_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDeductionDate_Edit");
                    if (eDeductionDate_Edit != null)
                    {
                        string vDeductionDate_EditURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDeductionDate_Edit.ClientID;
                        string vDeductionDate_EditScript = "window.open('" + vDeductionDate_EditURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eDeductionDate_Edit.Attributes["onClick"] = vDeductionDate_EditScript;
                    }
                    TextBox eBuildDate_Edit = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildDate_Edit");
                    if (eBuildDate_Edit != null)
                    {
                        string vBuildDate_EditURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_Edit.ClientID;
                        string vBuildDate_EditScript = "window.open('" + vBuildDate_EditURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eBuildDate_Edit.Attributes["onClick"] = vBuildDate_EditScript;
                    }
                    break;
                case FormViewMode.Insert:
                    //plSearch.Visible = false;
                    plDetailDataShow.Visible = false;
                    TextBox eBuildDate_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildDate_INS");
                    if (eBuildDate_INS != null)
                    {
                        //建檔日期
                        eBuildDate_INS.Text = DateTime.Today.ToString("yyyy/MM/dd");
                        string vBuildDate_INSURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_INS.ClientID;
                        string vBuildDate_INSScript = "window.open('" + vBuildDate_INSURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eBuildDate_INS.Attributes["onClick"] = vBuildDate_INSScript;
                        //建檔人工號
                        Label eBuildMan_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eBuildMan_INS");
                        eBuildMan_INS.Text = vLoginID;
                        //建檔人姓名
                        Label eBuildManName_INS = (Label)fvAnecdoteCaseA_Data.FindControl("eBuildManName_INS");
                        eBuildManName_INS.Text = Session["LoginName"].ToString().Trim();
                        //精勤扣款日期
                        TextBox eDeductionDate_INS = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDeductionDate_INS");
                        string vDeductionDate_INSURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDeductionDate_INS.ClientID;
                        string vDeductionDate_INSScript = "window.open('" + vDeductionDate_INSURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eDeductionDate_INS.Attributes["onClick"] = vDeductionDate_INSScript;
                    }
                    break;
            }
        }

        /// <summary>
        /// 更新 B 檔資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKCaseB_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            try
            {
                Label eCaseNo_Edit = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseBCaseNo_Edit");
                if (eCaseNo_Edit != null)
                {
                    //雖然是異動 B 檔，還是要在 A 檔的異動記錄加入一筆資料
                    DateTime vModifyDate = DateTime.Today;
                    string vHistoryNo = vModifyDate.Year.ToString("D4") + vModifyDate.Month.ToString("D2") + "MODB";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndex_Str = (vMaxNo.Replace(vHistoryNo, "").Trim() != "") ? vMaxNo.Replace(vHistoryNo, "").Trim() : "0";
                    int vIndex = Int32.Parse(vIndex_Str) + 1;
                    string vNewHistoryNo = vHistoryNo + vIndex.ToString("D4");
                    string vCaseNo_Temp = eCaseNo_Edit.Text.Trim();
                    //寫入 A 檔異動記錄
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseHistory] " + Environment.NewLine +
                                   "            ([HistoryNo],[CaseNo],[HasInsurance],[DepNo],[DepName],[BuildDate],[BuildMan],[Car_ID],[Driver], " + Environment.NewLine +
                                   "             [DriverName],[InsuMan],[AnecdotalResRatio],[IsNoDeduction],[DeductionDate],[Remark]," + Environment.NewLine +
                                   "             [ModifyType],[ModifyDate],[ModifyMan],[CaseOccurrence],[ERPCouseNo],[CaseClose]) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "',[CaseNo],[HasInsurance],[DepNo],[DepName],[BuildDate],[BuildMan],[Car_ID],[Driver], " + Environment.NewLine +
                                   "       [DriverName],[InsuMan],[AnecdotalResRatio],[IsNoDeduction],[DeductionDate],[Remark]," + Environment.NewLine +
                                   "       'MODB',GetDate(),'" + vLoginID + "',[CaseOccurrence],[ERPCouseNo],[CaseClose] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCase] " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入 B 檔異動記錄
                    Label eCaseNoItems_Edit = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseBCaseNoItems_Edit");
                    string vCaseNoItems_Temp = eCaseNoItems_Edit.Text.Trim();
                    vSQLStr_Temp = "select max(Items) MaxItem from AnecdoteCaseBHistory where HistoryNo = '" + vNewHistoryNo + "' ";
                    vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    int vItems = (vMaxNo != "") ? Int32.Parse(vMaxNo.Trim()) + 1 : 1;
                    //複製 B 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseBHistory] " + Environment.NewLine +
                                   "            ([HistoryNo],[ItemsH],[HistoryNoItems],[CaseNo],[Items],[CaseNoItems],[Relationship],[RelCar_ID], " + Environment.NewLine +
                                   "             [EstimatedAmount],[ThirdInsurance],[CompInsurance],[DriverSharing],[CompanySharing],[CarDamageAMT], " + Environment.NewLine +
                                   "             [PersonDamageAMT],[RelationComp],[ReconciliationDate],[Remark]," + Environment.NewLine +
                                   "             [ModifyType],[ModifyDate],[ModifyMan],[PassengerInsu])" + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "','" + vItems.ToString("D4") + "','" + vNewHistoryNo + vItems.ToString("D4") + "',[CaseNo],[Items],[CaseNoItems], " + Environment.NewLine +
                                   "       [Relationship],[RelCar_ID],[EstimatedAmount],[ThirdInsurance],[CompInsurance],[DriverSharing],[CompanySharing], " + Environment.NewLine +
                                   "       [CarDamageAMT],[PersonDamageAMT],[RelationComp],[ReconciliationDate],[Remark], " + Environment.NewLine +
                                   "       'MODB',GetDate(),'" + vLoginID + "',[PassengerInsu] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCaseB] " + Environment.NewLine +
                                   " where CaseNoItems = '" + vCaseNoItems_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);

                    //實際開始更新資料
                    TextBox eRelationship_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBRelationship_Edit");
                    string vRelationship_Temp = eRelationship_Edit.Text.Trim();
                    TextBox eRelCarID_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBRelCarID_Edit");
                    string vRelCarID_Temp = eRelCarID_Edit.Text.Trim().ToUpper();
                    TextBox eEstimatedAmount_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBEstimatedAmount_Edit");
                    string vEstimatedAmount_Temp = eEstimatedAmount_Edit.Text.Trim();
                    TextBox eThirdInsurance_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBThirdInsurance_Edit");
                    string vThirdInsurance_Temp = eThirdInsurance_Edit.Text.Trim();
                    TextBox eCompInsurance_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBCompInsurance_Edit");
                    string vCompInsurance_Temp = eCompInsurance_Edit.Text.Trim();
                    TextBox eDriverSharing_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBDriverSharing_Edit");
                    string vDriverSharing_Temp = eDriverSharing_Edit.Text.Trim();
                    TextBox eCompanySharing_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBCompanySharing_Edit");
                    string vCompanySharing_Temp = eCompanySharing_Edit.Text.Trim();
                    TextBox eCarDamageAMT_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBCarDamageAMT_Edit");
                    string vCarDamageAMT_Temp = eCarDamageAMT_Edit.Text.Trim();
                    TextBox ePersonDamageAMT_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBPersonDamageAMT_Edit");
                    string vPersonDamageAMT_Temp = ePersonDamageAMT_Edit.Text.Trim();
                    TextBox eRelationComp_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBRelationComp_Edit");
                    string vRelationComp_Temp = eRelationComp_Edit.Text.Trim();
                    TextBox eReconciliationDate_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBReconciliationDate_Edit");
                    string vReconciliationDate_Temp = (eReconciliationDate_Edit.Text.Trim() != "") ? DateTime.Parse(eReconciliationDate_Edit.Text.Trim()).ToShortDateString() : "";
                    TextBox ePassengerInsu_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBPassengerInsu_Edit");
                    string vPassengerInsu_Temp = ePassengerInsu_Edit.Text.Trim();
                    TextBox eRemark_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBRemark_Edit");
                    string vRemark_Temp = eRemark_Edit.Text.Trim();

                    sdsAnecdoteCaseB_Data.UpdateCommand = "UPDATE AnecdoteCaseB " + Environment.NewLine +
                                                          "   SET Relationship = @Relationship, RelCar_ID = @RelCar_ID, EstimatedAmount = @EstimatedAmount, " + Environment.NewLine +
                                                          "       ThirdInsurance = @ThirdInsurance, CompInsurance = @CompInsurance, DriverSharing = @DriverSharing, " + Environment.NewLine +
                                                          "       CompanySharing = @CompanySharing, CarDamageAMT = @CarDamageAMT, PersonDamageAMT = @PersonDamageAMT, " + Environment.NewLine +
                                                          "       RelationComp = @RelationComp, ReconciliationDate = @ReconciliationDate, PassengerInsu = @PassengerInsu, " + Environment.NewLine +
                                                          "       Remark = @Remark WHERE (CaseNoItems = @CaseNoItems)";
                    sdsAnecdoteCaseB_Data.UpdateParameters.Clear();
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("Relationship", DbType.String, (vRelationship_Temp.Trim() != "") ? vRelationship_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelCar_ID", DbType.String, (vRelCarID_Temp.Trim() != "") ? vRelCarID_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("EstimatedAmount", DbType.Double, (vEstimatedAmount_Temp.Trim() != "") ? vEstimatedAmount_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("ThirdInsurance", DbType.Double, (vThirdInsurance_Temp.Trim() != "") ? vThirdInsurance_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("CompInsurance", DbType.Double, (vCompInsurance_Temp.Trim() != "") ? vCompInsurance_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("DriverSharing", DbType.Double, (vDriverSharing_Temp.Trim() != "") ? vDriverSharing_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("CompanySharing", DbType.Double, (vCompanySharing_Temp.Trim() != "") ? vCompanySharing_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("CarDamageAMT", DbType.Double, (vCarDamageAMT_Temp.Trim() != "") ? vCarDamageAMT_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("PersonDamageAMT", DbType.Double, (vPersonDamageAMT_Temp.Trim() != "") ? vPersonDamageAMT_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelationComp", DbType.Double, (vRelationComp_Temp.Trim() != "") ? vRelationComp_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("ReconciliationDate", DbType.Date, (vReconciliationDate_Temp.Trim() != "") ? vReconciliationDate_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("PassengerInsu", DbType.Double, (vPassengerInsu_Temp.Trim() != "") ? vPassengerInsu_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("Remark", DbType.String, (vRemark_Temp.Trim() != "") ? vRemark_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("CaseNoItems", DbType.String, vCaseNoItems_Temp.Trim()));

                    sdsAnecdoteCaseB_Data.Update();
                    fvAnecdoteCaseB_Data.ChangeMode(FormViewMode.ReadOnly);
                    gridAnecdoteCaseB_List.DataBind();
                    fvAnecdoteCaseB_Data.DataBind();
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

        /// <summary>
        /// 新增 B 檔資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKCaseB_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            try
            {
                Label eCaseBCaseNo_INS = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseBCaseNo_INS");
                if (eCaseBCaseNo_INS != null)
                {
                    string vCaseNo_Temp = eCaseBCaseNo_INS.Text.Trim();
                    string vSQLStr_Temp = "select max(Items) MAxItem from AnecdoteCaseB where CaseNo = '" + vCaseNo_Temp + "' ";
                    string vMaxItem = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    int vItem = (vMaxItem != "") ? Int32.Parse(vMaxItem) + 1 : 1;
                    string vNewItems = vItem.ToString("D4");
                    string vNewCaseNoItems = vCaseNo_Temp.Trim() + vItem.ToString("D4");
                    TextBox eRelationship_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBRelationship_INS");
                    string vRelationship_Temp = eRelationship_INS.Text.Trim();
                    TextBox eRelCarID_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBRelCarID_INS");
                    string vRelCarID_Temp = eRelCarID_INS.Text.Trim().ToUpper();
                    TextBox eEstimatedAmount_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBEstimatedAmount_INS");
                    string vEstimatedAmount_Temp = eEstimatedAmount_INS.Text.Trim();
                    TextBox eThirdInsurance_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBThirdInsurance_INS");
                    string vThirdInsurance_Temp = eThirdInsurance_INS.Text.Trim();
                    TextBox eCompInsurance_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBCompInsurance_INS");
                    string vCompInsurance_Temp = eCompInsurance_INS.Text.Trim();
                    TextBox eDriverSharing_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBDriverSharing_INS");
                    string vDriverSharing_Temp = eDriverSharing_INS.Text.Trim();
                    TextBox eCompanySharing_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBCompanySharing_INS");
                    string vCompanySharing_Temp = eCompanySharing_INS.Text.Trim();
                    TextBox eCarDamageAMT_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBCarDamageAMT_INS");
                    string vCarDamageAMT_Temp = eCarDamageAMT_INS.Text.Trim();
                    TextBox ePersonDamageAMT_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBPersonDamageAMT_INS");
                    string vPersonDamageAMT_Temp = ePersonDamageAMT_INS.Text.Trim();
                    TextBox eRelationComp_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBRelationComp_INS");
                    string vRelationComp_Temp = eRelationComp_INS.Text.Trim();
                    TextBox eReconciliationDate_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBReconciliationDate_INS");
                    string vReconciliationDate_Temp = (eReconciliationDate_INS.Text.Trim() != "") ? DateTime.Parse(eReconciliationDate_INS.Text.Trim()).ToShortDateString() : "";
                    TextBox ePassengerInsu_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBPassengerInsu_INS");
                    string vPassengerInsu_Temp = ePassengerInsu_INS.Text.Trim();
                    TextBox eRemark_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBRemark_INS");
                    string vRemark_Temp = eRemark_INS.Text.Trim();

                    sdsAnecdoteCaseB_Data.InsertCommand = "INSERT INTO AnecdoteCaseB " + Environment.NewLine +
                                                          "            (CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, " + Environment.NewLine +
                                                          "             DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, " + Environment.NewLine +
                                                          "             PassengerInsu, Remark) " + Environment.NewLine +
                                                          "     VALUES (@CaseNo, @Items, @CaseNoItems, @Relationship, @RelCar_ID, @EstimatedAmount, @ThirdInsurance, @CompInsurance, " + Environment.NewLine +
                                                          "             @DriverSharing, @CompanySharing, @CarDamageAMT, @PersonDamageAMT, @RelationComp, @ReconciliationDate, @PassengerInsu, @Remark)";

                    sdsAnecdoteCaseB_Data.InsertParameters.Clear();
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("Items", DbType.String, vNewItems.Trim()));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CaseNoItems", DbType.String, vNewCaseNoItems.Trim()));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("Relationship", DbType.String, (vRelationship_Temp.Trim() != "") ? vRelationship_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelCar_ID", DbType.String, (vRelCarID_Temp.Trim() != "") ? vRelCarID_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("EstimatedAmount", DbType.Double, (vEstimatedAmount_Temp.Trim() != "") ? vEstimatedAmount_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("ThirdInsurance", DbType.Double, (vThirdInsurance_Temp.Trim() != "") ? vThirdInsurance_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CompInsurance", DbType.Double, (vCompInsurance_Temp.Trim() != "") ? vCompInsurance_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("DriverSharing", DbType.Double, (vDriverSharing_Temp.Trim() != "") ? vDriverSharing_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CompanySharing", DbType.Double, (vCompanySharing_Temp.Trim() != "") ? vCompanySharing_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CarDamageAMT", DbType.Double, (vCarDamageAMT_Temp.Trim() != "") ? vCarDamageAMT_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("PersonDamageAMT", DbType.Double, (vPersonDamageAMT_Temp.Trim() != "") ? vPersonDamageAMT_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelationComp", DbType.Double, (vRelationComp_Temp.Trim() != "") ? vRelationComp_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("ReconciliationDate", DbType.Date, (vReconciliationDate_Temp.Trim() != "") ? vReconciliationDate_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("PassengerInsu", DbType.Double, (vPassengerInsu_Temp.Trim() != "") ? vPassengerInsu_Temp.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark_Temp.Trim() != "") ? vRemark_Temp.Trim() : String.Empty));

                    sdsAnecdoteCaseB_Data.Insert();

                    fvAnecdoteCaseB_Data.ChangeMode(FormViewMode.ReadOnly);
                    gridAnecdoteCaseB_List.DataBind();
                    fvAnecdoteCaseB_Data.DataBind();
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

        /// <summary>
        /// 刪除 B 檔資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelCaseB_List_Click(object sender, EventArgs e)
        {
            try
            {
                Label eCaseBCaseNo_List = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseBCaseNo_List");
                if (eCaseBCaseNo_List != null)
                {
                    //雖然是刪除 B 檔，還是要在 A 檔的異動記錄加入一筆資料
                    DateTime vModifyDate = DateTime.Today;
                    string vHistoryNo = vModifyDate.Year.ToString("D4") + vModifyDate.Month.ToString("D2") + "DELB";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndex_Str = (vMaxNo.Replace(vHistoryNo, "").Trim() != "") ? vMaxNo.Replace(vHistoryNo, "").Trim() : "0";
                    int vIndex = Int32.Parse(vIndex_Str) + 1;
                    string vNewHistoryNo = vHistoryNo + vIndex.ToString("D4");
                    string vCaseNo_Temp = eCaseBCaseNo_List.Text.Trim();
                    //寫入 A 檔異動記錄
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseHistory] " + Environment.NewLine +
                                   "            ([HistoryNo],[CaseNo],[HasInsurance],[DepNo],[DepName],[BuildDate],[BuildMan],[Car_ID],[Driver], " + Environment.NewLine +
                                   "             [DriverName],[InsuMan],[AnecdotalResRatio],[IsNoDeduction],[DeductionDate],[Remark]," + Environment.NewLine +
                                   "             [ModifyType],[ModifyDate],[ModifyMan],[CaseOccurrence],[ERPCouseNo],[CaseClose]) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "',[CaseNo],[HasInsurance],[DepNo],[DepName],[BuildDate],[BuildMan],[Car_ID],[Driver], " + Environment.NewLine +
                                   "       [DriverName],[InsuMan],[AnecdotalResRatio],[IsNoDeduction],[DeductionDate],[Remark]," + Environment.NewLine +
                                   "       'DELB',GetDate(),'" + vLoginID + "',[CaseOccurrence],[ERPCouseNo],[CaseClose] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCase] " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入 B 檔異動記錄
                    Label eCaseNoItems_List = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseBCaseNoItems_List");
                    string vCaseNoItems_Temp = eCaseNoItems_List.Text.Trim();
                    vSQLStr_Temp = "select max(Items) MaxItem from AnecdoteCaseBHistory where HistoryNo = '" + vNewHistoryNo + "' ";
                    vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    int vItems = (vMaxNo != "") ? Int32.Parse(vMaxNo.Trim()) + 1 : 1;
                    //複製 B 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseBHistory] " + Environment.NewLine +
                                   "            ([HistoryNo],[ItemsH],[HistoryNoItems],[CaseNo],[Items],[CaseNoItems],[Relationship],[RelCar_ID], " + Environment.NewLine +
                                   "             [EstimatedAmount],[ThirdInsurance],[CompInsurance],[DriverSharing],[CompanySharing],[CarDamageAMT], " + Environment.NewLine +
                                   "             [PersonDamageAMT],[RelationComp],[ReconciliationDate],[Remark]," + Environment.NewLine +
                                   "             [ModifyType],[ModifyDate],[ModifyMan],[PassengerInsu])" + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "','" + vItems.ToString("D4") + "','" + vNewHistoryNo + vItems.ToString("D4") + "',[CaseNo],[Items],[CaseNoItems], " + Environment.NewLine +
                                   "       [Relationship],[RelCar_ID],[EstimatedAmount],[ThirdInsurance],[CompInsurance],[DriverSharing],[CompanySharing], " + Environment.NewLine +
                                   "       [CarDamageAMT],[PersonDamageAMT],[RelationComp],[ReconciliationDate],[Remark], " + Environment.NewLine +
                                   "       'DELB',GetDate(),'" + vLoginID + "',[PassengerInsu] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCaseB] " + Environment.NewLine +
                                   " where CaseNoItems = '" + vCaseNoItems_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);

                    //開始刪除
                    using (SqlDataSource dsDelTemp = new SqlDataSource())
                    {
                        dsDelTemp.ConnectionString = vConnStr;
                        //刪除 B 檔
                        dsDelTemp.DeleteCommand = "delete AnecdoteCaseB where CaseNoItems = @CaseNoItems ";
                        dsDelTemp.DeleteParameters.Clear();
                        dsDelTemp.DeleteParameters.Add(new Parameter("CaseNoItems", DbType.String, vCaseNoItems_Temp.Trim()));
                        dsDelTemp.Delete();
                    }
                    fvAnecdoteCaseB_Data.ChangeMode(FormViewMode.ReadOnly);
                    gridAnecdoteCaseB_List.DataBind();
                    fvAnecdoteCaseB_Data.DataBind();
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

        protected void fvAnecdoteCaseB_Data_DataBound(object sender, EventArgs e)
        {
            string vTempDateURL = "";
            string vTempDateScript = "";

            switch (fvAnecdoteCaseB_Data.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    TextBox eCaseBReconciliationDate_Edit = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBReconciliationDate_Edit");
                    if (eCaseBReconciliationDate_Edit != null)
                    {
                        vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseBReconciliationDate_Edit.ClientID;
                        vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eCaseBReconciliationDate_Edit.Attributes["onClick"] = vTempDateScript;
                    }
                    break;
                case FormViewMode.Insert:
                    Label eCaseNo_List = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_List");
                    string vCaseNo = eCaseNo_List.Text.Trim();
                    Label eCaseBCaseNo = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseBCaseNo_INS");
                    eCaseBCaseNo.Text = vCaseNo;
                    TextBox eCaseBReconciliationDate_INS = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCaseBReconciliationDate_INS");
                    if (eCaseBReconciliationDate_INS != null)
                    {
                        vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseBReconciliationDate_INS.ClientID;
                        vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eCaseBReconciliationDate_INS.Attributes["onClick"] = vTempDateScript;
                    }
                    break;
            }
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
            string vSelStr = GetSelectStr_Excel();
            string vFileName = "肇事案件處理清冊";
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
                        vHeaderText = (drExcel.GetName(i) == "CaseNo") ? "序號" :
                                      (drExcel.GetName(i) == "HasInsurance") ? "是否出險" :
                                      (drExcel.GetName(i) == "DepNo") ? "" :
                                      (drExcel.GetName(i) == "DepName") ? "站別" :
                                      (drExcel.GetName(i) == "BuildDate") ? "發生日期" :
                                      (drExcel.GetName(i) == "BuildMan") ? "" :
                                      (drExcel.GetName(i) == "BuildManName") ? "建檔人" :
                                      (drExcel.GetName(i) == "Car_ID") ? "牌照號碼" :
                                      (drExcel.GetName(i) == "Driver") ? "" :
                                      (drExcel.GetName(i) == "DriverName") ? "駕駛員" :
                                      (drExcel.GetName(i) == "InsuMan") ? "保險經辦人" :
                                      (drExcel.GetName(i) == "AnecdotalResRatio") ? "肇責比率" :
                                      (drExcel.GetName(i) == "IsNoDeduction") ? "免扣精勤" :
                                      (drExcel.GetName(i) == "DeductionDate") ? "扣發日期" :
                                      (drExcel.GetName(i) == "Remark") ? "備註" :
                                      (drExcel.GetName(i) == "CaseOccurrence") ? "肇事經過" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i) == "DeductionDate") ||
                                 (drExcel.GetName(i) == "BuildDate")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if ((drExcel.GetName(i) == "AnecdotalResRatio") && (drExcel[i].ToString() != ""))
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
                            string vRecordInsuranceStr = (chkHasInsurance.Checked) ? "己出險" : "全部";
                            string vRecordDateStr = ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() != "")) ? eCaseDate_Start_Search.Text.Trim() + "~" + eCaseDate_End_Search.Text.Trim() :
                                                    ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() == "")) ? eCaseDate_Start_Search.Text.Trim() :
                                                    ((eCaseDate_Start_Search.Text.Trim() == "") && (eCaseDate_End_Search.Text.Trim() != "")) ? eCaseDate_End_Search.Text.Trim() : "不指定日期";
                            string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "不指定車號";
                            string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "所有駕駛員";
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "AnecdoteCase.aspx" + Environment.NewLine +
                                                 "出險情況：" + vRecordInsuranceStr + Environment.NewLine +
                                                 "發生日期：" + vRecordDateStr + Environment.NewLine +
                                                 "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                                 "駕駛員：" + vDriverStr;
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

        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            string vDriver = eDriver_Search.Text.Trim();
            string vDriverName = "";
            string vSQLStr_Temp = "select [Name] from  Employee where EmpNo = '" + vDriver + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vDriverName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriver;
                vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vDriverName + "' order by EmpNo DESC";
                vDriver = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Search.Text = vDriver.Trim();
            eDriverName_Search.Text = vDriverName.Trim();
        }

        protected void bbExportERP_Click(object sender, EventArgs e)
        {
            string vSQLStr_Temp = "";
            Label eCaseNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_List");
            if (eCaseNo_Temp != null)
            {
                Label eDepNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eDepNo_List");
                Label eBuildDate_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eBuildDate_List");
                Label eCarID_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eCarID_List");
                Label eBuildMan_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eBuildMan_List");
                Label eDriver_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eDriver_List");
                TextBox eRemark_Temp = (TextBox)fvAnecdoteCaseA_Data.FindControl("eRemark_List");
                TextBox eCaseOccurrence_Temp = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseOccurrence_List");
                Label eERPCouseNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eERPCouseNo_List");
                string vCaseNo = eCaseNo_Temp.Text.Trim();
                string vBuileDate_TempStr = eBuildDate_Temp.Text.Trim();
                string[] vaBuildDate_Temp = vBuileDate_TempStr.Split('/');
                int vYear_Temp = ((vaBuildDate_Temp != null) && (vaBuildDate_Temp[0].Trim() != "")) ? Int32.Parse(vaBuildDate_Temp[0].Trim()) : 0;
                int vMonth_Temp = ((vaBuildDate_Temp != null) && (vaBuildDate_Temp[1].Trim() != "")) ? Int32.Parse(vaBuildDate_Temp[1].Trim()) : 0;
                int vDay_Temp = ((vaBuildDate_Temp != null) && (vaBuildDate_Temp[2].Trim() != "")) ? Int32.Parse(vaBuildDate_Temp[2].Trim()) : 0;
                string vCouseDate_New = (vYear_Temp + vMonth_Temp + vDay_Temp == 0) ? "" :
                                        (vYear_Temp < 1911) ? (vYear_Temp + 1911).ToString("D4") + "/" + vMonth_Temp.ToString("D2") + "/" + vDay_Temp.ToString("D2") :
                                         vYear_Temp.ToString("D4") + "/" + vMonth_Temp.ToString("D2") + "/" + vDay_Temp.ToString("D2");
                string vCaseYM = (DateTime.Today.Year - 1911).ToString("D3") + DateTime.Today.Month.ToString("D2");
                vSQLStr_Temp = "select MAX(CouseNo) MaxNo from Car_CouseTrouble where CouseNo like '" + vCaseYM + "%' ";
                string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                int vIndex = (vMaxIndex != "") ? Int32.Parse(vMaxIndex.Replace(vCaseYM, "")) + 1 : 1;
                string vCouseNo_New = vCaseYM + vIndex.ToString("D4");
                string vDepNo_New = eDepNo_Temp.Text.Trim();
                string vCarID_New = eCarID_Temp.Text.Trim();
                vSQLStr_Temp = "select (isnull(Car_ID, '') + ',' + isnull(Car_Class, '') + ',' + " + Environment.NewLine +
                               "        isnull(convert(varchar(10),ProdDate, 111), '') + ',' + " + Environment.NewLine +
                               "        isnull(BCENo, '') + ',' + isnull(Exhaust, '')) CarData " + Environment.NewLine +
                               "  from Car_InfoA " + Environment.NewLine +
                               " where Car_ID = '" + vCarID_New + "' ";
                string vCarData = PF.GetValue(vConnStr, vSQLStr_Temp, "CarData");
                string[] vaCarData_New = vCarData.Split(',');
                string vCarClass_New = (vaCarData_New != null) ? vaCarData_New[1].ToString().Trim() : "";
                string vProdDate_New = (vaCarData_New != null) ? vaCarData_New[2].ToString().Trim() : "";
                string vBCENo_New = (vaCarData_New != null) ? vaCarData_New[3].ToString().Trim() : "";
                string vExhaust_New = (vaCarData_New != null) ? vaCarData_New[4].ToString().Trim() : "";
                string vEsempNo_New = eBuildMan_Temp.Text.Trim();
                string vDriver_New = eDriver_Temp.Text.Trim();
                string vMemo_New = eRemark_Temp.Text.Trim();
                string vCouseCondition_New = eCaseOccurrence_Temp.Text.Trim();
                vSQLStr_Temp = "insert into Car_CouseTrouble(CouseNo, DepNo, CouseDate, Car_ID, Car_Class, ProdDate, Displacement, EngineNo, EsempNo, Driver, Memo, Course_Condition) " + Environment.NewLine +
                               "values(@CouseNo, @DepNo, @CouseDate, @Car_ID, @Car_Class, @ProdDate, @Displacement, @EngineNo, @EsempNo, @Driver, @Memo, @Course_Condition)";
                try
                {
                    using (SqlConnection connInsert = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdInsert = new SqlCommand(vSQLStr_Temp, connInsert);
                        cmdInsert.Parameters.Clear();
                        cmdInsert.Parameters.Add(new SqlParameter("CouseNo", vCouseNo_New));
                        cmdInsert.Parameters.Add(new SqlParameter("DepNo", vDepNo_New));
                        cmdInsert.Parameters.Add(new SqlParameter("CouseDate", vCouseDate_New));
                        cmdInsert.Parameters.Add(new SqlParameter("Car_ID", vCarID_New));
                        cmdInsert.Parameters.Add(new SqlParameter("Car_Class", vCarClass_New));
                        cmdInsert.Parameters.Add(new SqlParameter("ProdDate", vProdDate_New));
                        cmdInsert.Parameters.Add(new SqlParameter("Displacement", vExhaust_New));
                        cmdInsert.Parameters.Add(new SqlParameter("EngineNo", vBCENo_New));
                        cmdInsert.Parameters.Add(new SqlParameter("EsempNo", vEsempNo_New));
                        cmdInsert.Parameters.Add(new SqlParameter("Driver", vDriver_New));
                        cmdInsert.Parameters.Add(new SqlParameter("Memo", vMemo_New));
                        cmdInsert.Parameters.Add(new SqlParameter("Course_Condition", vCouseCondition_New));
                        connInsert.Open();
                        cmdInsert.ExecuteNonQuery();
                    }
                    vSQLStr_Temp = "update AnecdoteCase set ERPCouseNo = @ERPCouseNo where CaseNo = @CaseNo";
                    using (SqlConnection connUpdate = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdUpdate = new SqlCommand(vSQLStr_Temp, connUpdate);
                        cmdUpdate.Parameters.Clear();
                        cmdUpdate.Parameters.Add(new SqlParameter("ERPCouseNo", vCouseNo_New));
                        cmdUpdate.Parameters.Add(new SqlParameter("CaseNo", vCaseNo));
                        connUpdate.Open();
                        cmdUpdate.ExecuteNonQuery();
                    }
                    gridAnecdoteCaseA_List.DataBind();
                    fvAnecdoteCaseA_Data.DataBind();
                    string vRecordNote = "與ERP同步" + Environment.NewLine +
                                         "AnecdoteCase.aspx" + Environment.NewLine +
                                         "線上系統單號：" + vCaseNo + Environment.NewLine +
                                         "ERP單號：" + vCouseNo_New;
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

        protected void eCaseClose_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbCaseClose_Temp = (CheckBox)fvAnecdoteCaseA_Data.FindControl("eCaseClose_Edit");
            Label eCaseCloseTitle_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseCloseTitle_Edit");
            if (cbCaseClose_Temp != null)
            {
                eCaseCloseTitle_Temp.Text = (cbCaseClose_Temp.Checked == true) ? "True" : "False";
            }
        }

        protected void eCaseClose_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbCaseClose_Temp = (CheckBox)fvAnecdoteCaseA_Data.FindControl("eCaseClose_INS");
            Label eCaseCloseTitle_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseCloseTitle_INS");
            if (cbCaseClose_Temp != null)
            {
                eCaseCloseTitle_Temp.Text = (cbCaseClose_Temp.Checked == true) ? "True" : "False";
            }
        }

        protected void bbDelERP_List_Click(object sender, EventArgs e)
        {
            Label eCaseNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_List");
            Label eERPCouseNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eERPCouseNo_List");
            if ((eERPCouseNo_Temp != null) && (eERPCouseNo_Temp.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSQLStr_Temp = "delete Car_CouseTrouble where CouseNo = @CouseNo";
                try
                {
                    using (SqlConnection connDelete = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdDelete = new SqlCommand(vSQLStr_Temp, connDelete);
                        cmdDelete.Parameters.Clear();
                        cmdDelete.Parameters.Add(new SqlParameter("CouseNo", eERPCouseNo_Temp.Text.Trim()));
                        connDelete.Open();
                        cmdDelete.ExecuteNonQuery();
                    }
                    vSQLStr_Temp = "update AnecdoteCase set ERPCouseNo = null where CaseNo = @CaseNo";
                    using (SqlConnection connUpdate=new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdUpdate = new SqlCommand(vSQLStr_Temp, connUpdate);
                        cmdUpdate.Parameters.Clear();
                        cmdUpdate.Parameters.Add(new SqlParameter("CaseNo", eCaseNo_Temp.Text.Trim()));
                        connUpdate.Open();
                        cmdUpdate.ExecuteNonQuery();
                    }
                    gridAnecdoteCaseA_List.DataBind();
                    fvAnecdoteCaseA_Data.DataBind();
                    string vRecordNote = "取消與ERP同步" + Environment.NewLine +
                                         "AnecdoteCase.aspx" + Environment.NewLine +
                                         "線上系統單號：" + eCaseNo_Temp.Text.Trim() + Environment.NewLine +
                                         "ERP單號：" + eERPCouseNo_Temp.Text.Trim();
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