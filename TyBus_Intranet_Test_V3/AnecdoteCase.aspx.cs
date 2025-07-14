using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class AnecdoteCase : Page
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
        private string vCaseDateURL;
        private string vCaseDateScript;

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
                    //查詢介面出險日期_起
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Start_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Start_Search.Attributes["onClick"] = vCaseDateScript;
                    //查詢介面出險日期_迄
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_End_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_End_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        //不是因為觸發 PostBack 導致的畫面刷新，而是第一次打開畫面
                        vStationNo = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vLoginDepNo + "' and InSHReport = 'V'", "DepNo");

                        eDepNo_Start_Search.Text = vStationNo.Trim();
                        eDepNo_End_Search.Text = "";
                        eDepNo_Start_Search.Enabled = (vStationNo == "");
                        eDepNo_End_Search.Enabled = (vStationNo == "");

                        plPrint.Visible = false;
                        plSearch.Visible = true;
                        plMainDataShow.Visible = true;
                        plDetailDataShow.Visible = true;
                        eErrorMSG_Main.Text = "";
                        eErrorMSG_Main.Visible = false;
                        eErrorMSG_A.Text = "";
                        eErrorMSG_A.Visible = false;
                        eErrorMSG_B.Text = "";
                        eErrorMSG_B.Visible = false;
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
                                //2024.05.08 新增欄位
                                //"       IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose " + Environment.NewLine +
                                "       IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose, " + Environment.NewLine +
                                "       IsExemption, PaidAmount, InsuAmount, Penalty, PenaltyRatio " + Environment.NewLine +
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
        /// 取回匯出 EXCEL 檔用的查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelStr_Excel()
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
            string vResultStr = "select a.CaseNo, case when a.HasInsurance = 1 then '是' else '否' end HasInsurance, " + Environment.NewLine +
                                "       a.DepNo, a.DepName, convert(varchar(10), a.BuildDate, 111) BuildDate, a.BuildMan, " + Environment.NewLine +
                                "       e.[Name] BuildManName, a.Car_ID, a.Driver, a.DriverName, a.InsuMan, a.AnecdotalResRatio, " + Environment.NewLine +
                                "       case when a.IsNoDeduction = 1 then '是' else '否' end IsNoDeduction, " + Environment.NewLine +
                                "       convert(varchar(10), a.DeductionDate, 111) DeductionDate, a.Remark, a.CaseOccurrence, " + Environment.NewLine +
                                "       a.ERPCouseNo, case when a.CaseClose = 1 then '是' else '否' end CaseClose, " + Environment.NewLine +
                                "       case when a.IsExemption = 'Y' then '是' else '否' end IsExemption, a.PaisAmount, a.Penalty, a.PenaltyRatio, " + Environment.NewLine +
                                "       a.InsuAmount, a.IDCardNo, convert(varchar(10), a.Birthday, 111) Birthday, convert(varchar(10), a.Assumeday, 111) Assumeday, " + Environment.NewLine +
                                "       a.TelephoneNo, a.Address, a.PersonDamage, a.CarDamage, convert(varchar(10), a.ReportDate, 111) ReportDate, " + Environment.NewLine +
                                "       convert(varchar(10), a.CaseDate, 111) CaseDate, a.CaseTime, a.OutReportNo, a.CasePosition, a.PoliceUnit, a.PoliceName, " + Environment.NewLine +
                                "       case when a.HasVideo = 'Y' then '有' else '無' end HasVideo, a.NoVideoReason, " + Environment.NewLine +
                                "       case when a.HasCaseData = 'V' then '有' else '無' end HasCaseData, a.ModifyMan, " + Environment.NewLine +
                                "       convert(varchar(10), a.ModifyDate, 111) ModifyDate, case when a.HasAccReport = 'Y' then '有' else '無' end HasAccReport " + Environment.NewLine +
                                "  from AnecdoteCase as a left join Employee as e on e.EmpNo = a.BuildMan " + Environment.NewLine +
                                " where isnull(a.CaseNo, '') <> '' " + Environment.NewLine +
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
        /// 繫結主檔資料
        /// </summary>
        protected void CaseDataBind()
        {
            if (vLoginID != "")
            {
                sdsAnecdoteCaseA_List.SelectCommand = GetSelectStr();
                gridAnecdoteCaseA_List.DataBind();
            }
        }

        /// <summary>
        /// 查詢畫面帶回駕駛員姓名
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            //先假設畫面上輸入的是工號，用工號抓姓名
            string vDriver = eDriver_Search.Text.Trim();
            string vDriverName = "";
            string vSQLStr_Temp = "select [Name] from  Employee where EmpNo = '" + vDriver + "' ";
            vDriverName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName == "")
            {
                //如果工號抓不到姓名，可能是因為輸入的是姓名，所以改用姓名抓工號
                vDriverName = vDriver;
                vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vDriverName + "' order by EmpNo DESC";
                vDriver = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Search.Text = vDriver.Trim();
            eDriverName_Search.Text = vDriverName.Trim();
        }

        /// <summary>
        /// 取回查詢資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSearch_Click(object sender, EventArgs e)
        {
            CaseDataBind();
        }

        /// <summary>
        /// 匯出 EXCEL 檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
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
            string vSelStr = GetSelStr_Excel();
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
                    try
                    {
                        //寫入標題列
                        vLinesNo = 0;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            vHeaderText = (drExcel.GetName(i) == "CaseNo") ? "肇事單號" :
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
                                          (drExcel.GetName(i) == "CaseOccurrence") ? "肇事經過" :
                                          (drExcel.GetName(i) == "ERPCouseNo") ? "ERP單號" :
                                          (drExcel.GetName(i) == "CaseClose") ? "是否和解" :
                                          (drExcel.GetName(i) == "IsExemption") ? "是否裁定免責" :
                                          (drExcel.GetName(i) == "PaidAmount") ? "已自付總額" :
                                          (drExcel.GetName(i) == "Penalty") ? "罰款分擔金額" :
                                          (drExcel.GetName(i) == "PenaltyRatio") ? "罰款分擔比例" :
                                          (drExcel.GetName(i) == "InsuAmount") ? "保險理賠金" :
                                          (drExcel.GetName(i) == "IDCardNo") ? "身分證號" :
                                          (drExcel.GetName(i) == "Birthday") ? "出生日期" :
                                          (drExcel.GetName(i) == "Assumeday") ? "到職日期" :
                                          (drExcel.GetName(i) == "TelephoneNo") ? "連絡電話" :
                                          (drExcel.GetName(i) == "Address") ? "戶籍地址" :
                                          (drExcel.GetName(i) == "PersonDamage") ? "體傷" :
                                          (drExcel.GetName(i) == "CarDamage") ? "車 (財) 損" :
                                          (drExcel.GetName(i) == "ReportDate") ? "報告日期" :
                                          (drExcel.GetName(i) == "CaseDate") ? "事件日期" :
                                          (drExcel.GetName(i) == "CaseTime") ? "事件時間" :
                                          (drExcel.GetName(i) == "OutReportNo") ? "登記聯單編號" :
                                          (drExcel.GetName(i) == "CasePosition") ? "事件地點" :
                                          (drExcel.GetName(i) == "PoliceUnit") ? "警方受理單位" :
                                          (drExcel.GetName(i) == "PoliceName") ? "承辦員警" :
                                          (drExcel.GetName(i) == "HasVideo") ? "是否有行車影像" :
                                          (drExcel.GetName(i) == "NoVideoReason") ? "無行車影像原因" :
                                          (drExcel.GetName(i) == "HasCaseData") ? "是否已申請案件資料" :
                                          (drExcel.GetName(i) == "ModifyMan") ? "修改人" :
                                          (drExcel.GetName(i) == "ModifyDate") ? "修改日期" :
                                          (drExcel.GetName(i) == "HasAccReport") ? "是否申請鑑定" : drExcel.GetName(i);
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
                                if (((drExcel.GetName(i).ToUpper() == "DEDUCTIONDATE") ||
                                     (drExcel.GetName(i).ToUpper() == "BUILDDATE") ||
                                     (drExcel.GetName(i).ToUpper() == "BIRTHDAY") ||
                                     (drExcel.GetName(i).ToUpper() == "ASSUMEDAY") ||
                                     (drExcel.GetName(i).ToUpper() == "REPORTDATE") ||
                                     (drExcel.GetName(i).ToUpper() == "CASEDATE") ||
                                     (drExcel.GetName(i).ToUpper() == "MODIFYDATE")) && (drExcel[i].ToString() != ""))
                                {
                                    //日期欄位
                                    string vTempStr = drExcel[i].ToString();
                                    vBuDate = DateTime.Parse(drExcel[i].ToString());
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                }
                                else if ((drExcel.GetName(i).ToUpper() == "ANECDOTALRESRATIO") ||
                                         (drExcel.GetName(i).ToUpper() == "PAIDAMOUNT") ||
                                         (drExcel.GetName(i).ToUpper() == "PENALY") ||
                                         (drExcel.GetName(i).ToUpper() == "PENALYRATIO") ||
                                         (drExcel.GetName(i).ToUpper() == "INSUAMOUNT") && (drExcel[i].ToString() != ""))
                                {
                                    //浮點數欄位
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
                        eErrorMSG_Main.Text = eMessage.Message;
                        eErrorMSG_Main.Visible = true;
                    }
                }
                else
                {
                    eErrorMSG_Main.Text = "查無資料可匯出！";
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 回主畫面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCancel_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// A 檔 FormView 已經完成資料繫結
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void fvAnecdoteCaseA_Data_DataBound(object sender, EventArgs e)
        {
            Label eCaseNo_A;
            TextBox eBuildDate;
            TextBox eReportDate;
            TextBox eDeductionDate;
            TextBox eBirthday;
            TextBox eCaseDate;
            TextBox eBuildMan;
            CheckBox cbHasInsurance;
            CheckBox cbCaseClose;
            CheckBox cbIsNoDeduction;
            CheckBox cbHasVideo;
            CheckBox cbHasAccReport;
            CheckBox cbIsExemption;
            Label eHasInsurance;
            Label eCaseClose;
            Label eIsNoDeduction;
            Label eHasVideo;
            Label eHasCaseData;
            Label eHasAccReport;
            Label eIsExemption;
            Label eBuildManName;
            RadioButtonList rbHasCaseData;

            bool vAllowDep = ((vLoginID.ToUpper() == "SUPERVISOR") || (vLoginDepNo == "09") || (vLoginDepNo == "06"));
            switch (fvAnecdoteCaseA_Data.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    plDetailDataShow.Visible = ((gridAnecdoteCaseA_List.Rows.Count > 0) && (gridAnecdoteCaseA_List.SelectedIndex >= 0));
                    //在畫面上尋找 bbExportERP_List 這個按鈕，有被實體化才往下做
                    Button bbExportERP = (Button)fvAnecdoteCaseA_Data.FindControl("bbExportERP_List");
                    if (bbExportERP != null)
                    {
                        Label eERPCouseNo = (Label)fvAnecdoteCaseA_Data.FindControl("eERPCauseNo_List");
                        Button bbDelERP = (Button)fvAnecdoteCaseA_Data.FindControl("bbDelERP_List");
                        bbExportERP.Visible = ((vAllowDep) && (eERPCouseNo.Text.Trim() == ""));
                        bbDelERP.Visible = ((vAllowDep) && (eERPCouseNo.Text.Trim() != ""));
                        //設定 CheckBox 勾選狀態
                        cbHasInsurance = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasInsurance_List");
                        eHasInsurance = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsurance_List");
                        cbHasInsurance.Checked = (eHasInsurance.Text.Trim().ToUpper() == "TRUE") ? true : false;
                        cbCaseClose = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbCaseClose_List");
                        eCaseClose = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseClose_List");
                        cbCaseClose.Checked = (eCaseClose.Text.Trim().ToUpper() == "TRUE") ? true : false;
                        cbIsNoDeduction = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsNoDeduction_List");
                        eIsNoDeduction = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeduction_List");
                        cbIsNoDeduction.Checked = (eIsNoDeduction.Text.Trim().ToUpper() == "TRUE") ? true : false;
                        cbHasVideo = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasVideo_List");
                        eHasVideo = (Label)fvAnecdoteCaseA_Data.FindControl("eHasVideo_List");
                        cbHasVideo.Checked = (eHasVideo.Text.Trim().ToUpper() == "Y") ? true : false;
                        rbHasCaseData = (RadioButtonList)fvAnecdoteCaseA_Data.FindControl("rbHasCaseData_List");
                        eHasCaseData = (Label)fvAnecdoteCaseA_Data.FindControl("eHasCaseData_List");
                        rbHasCaseData.SelectedIndex = rbHasCaseData.Items.IndexOf(rbHasCaseData.Items.FindByValue(eHasCaseData.Text.Trim().ToUpper()));
                        cbHasAccReport = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasAccReport_List");
                        eHasAccReport = (Label)fvAnecdoteCaseA_Data.FindControl("eHasAccReport_List");
                        cbHasAccReport.Checked = (eHasAccReport.Text.Trim().ToUpper() == "Y") ? true : false;
                        cbIsExemption = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsExemption_List");
                        eIsExemption = (Label)fvAnecdoteCaseA_Data.FindControl("eIsExemption_List");
                        cbIsExemption.Checked = (eIsExemption.Text.Trim().ToUpper() == "Y") ? true : false;
                    }
                    break;
                case FormViewMode.Edit:
                    //先在畫面上尋找 eCaseNo_A 這個欄位，這個欄位有被實體化才往下做
                    eCaseNo_A = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_A_Edit");
                    if (eCaseNo_A != null)
                    {
                        //設定幾個日期輸入欄位的動作
                        eBuildDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildDate_Edit");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eBuildDate.Attributes["onClick"] = vCaseDateScript;
                        eReportDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eReportDate_Edit");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eReportDate.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eReportDate.Attributes["onClick"] = vCaseDateScript;
                        eDeductionDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDeductionDate_Edit");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDeductionDate.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eDeductionDate.Attributes["onClick"] = vCaseDateScript;
                        eBirthday = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBirthday_Edit");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBirthday.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eBirthday.Attributes["onClick"] = vCaseDateScript;
                        eCaseDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseDate_Edit");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eCaseDate.Attributes["onClick"] = vCaseDateScript;
                        //設定 CheckBox 勾選狀態
                        cbHasInsurance = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasInsurance_Edit");
                        eHasInsurance = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsurance_Edit");
                        cbHasInsurance.Checked = (eHasInsurance.Text.Trim().ToUpper() == "TRUE");
                        cbCaseClose = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbCaseClose_Edit");
                        eCaseClose = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseClose_Edit");
                        cbCaseClose.Checked = (eCaseClose.Text.Trim().ToUpper() == "TRUE");
                        cbIsNoDeduction = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsNoDeduction_Edit");
                        eIsNoDeduction = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeduction_Edit");
                        cbIsNoDeduction.Checked = (eIsNoDeduction.Text.Trim().ToUpper() == "TRUE");
                        cbHasVideo = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasVideo_Edit");
                        eHasVideo = (Label)fvAnecdoteCaseA_Data.FindControl("eHasVideo_Edit");
                        cbHasVideo.Checked = (eHasVideo.Text.Trim().ToUpper() == "Y");
                        rbHasCaseData = (RadioButtonList)fvAnecdoteCaseA_Data.FindControl("rbHasCaseData_Edit");
                        eHasCaseData = (Label)fvAnecdoteCaseA_Data.FindControl("eHasCaseData_Edit");
                        rbHasCaseData.SelectedIndex = rbHasCaseData.Items.IndexOf(rbHasCaseData.Items.FindByValue(eHasCaseData.Text.Trim().ToUpper()));
                        cbHasAccReport = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasAccReport_Edit");
                        eHasAccReport = (Label)fvAnecdoteCaseA_Data.FindControl("eHasAccReport_Edit");
                        cbHasAccReport.Checked = (eHasAccReport.Text.Trim().ToUpper() == "Y");
                        cbIsExemption = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsExemption_Edit");
                        eIsExemption = (Label)fvAnecdoteCaseA_Data.FindControl("eIsExemption_Edit");
                        cbIsExemption.Checked = (eIsExemption.Text.Trim().ToUpper() == "Y");
                    }
                    break;
                case FormViewMode.Insert:
                    //先在畫面上尋找 eCaseNo_A 這個欄位，這個欄位有被實體化才往下做
                    eCaseNo_A = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_A_INS");
                    if (eCaseNo_A != null)
                    {
                        //建檔人直接預帶登入人
                        eBuildMan = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildMan_INS");
                        eBuildMan.Text = vLoginID;
                        eBuildManName = (Label)fvAnecdoteCaseA_Data.FindControl("eBuildManName_INS");
                        eBuildManName.Text = vLoginName;
                        //設定幾個日期輸入欄位的動作
                        eBuildDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildDate_INS");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eBuildDate.Attributes["onClick"] = vCaseDateScript;
                        eBuildDate.Text = DateTime.Today.ToShortDateString(); //出險日期預帶當天
                        eReportDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eReportDate_INS");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eReportDate.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eReportDate.Attributes["onClick"] = vCaseDateScript;
                        eDeductionDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDeductionDate_INS");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDeductionDate.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eDeductionDate.Attributes["onClick"] = vCaseDateScript;
                        eBirthday = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBirthday_INS");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBirthday.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eBirthday.Attributes["onClick"] = vCaseDateScript;
                        eCaseDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseDate_INS");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eCaseDate.Attributes["onClick"] = vCaseDateScript;
                        //設定 CheckBox 勾選狀態
                        cbHasInsurance = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasInsurance_INS");
                        eHasInsurance = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsurance_INS");
                        cbHasInsurance.Checked = false;
                        eHasInsurance.Text = "false";
                        cbCaseClose = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbCaseClose_INS");
                        eCaseClose = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseClose_INS");
                        cbCaseClose.Checked = false;
                        eCaseClose.Text = "false";
                        cbIsNoDeduction = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsNoDeduction_INS");
                        eIsNoDeduction = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeduction_INS");
                        cbIsNoDeduction.Checked = false;
                        eIsNoDeduction.Text = "false";
                        cbHasVideo = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasVideo_INS");
                        eHasVideo = (Label)fvAnecdoteCaseA_Data.FindControl("eHasVideo_INS");
                        cbHasVideo.Checked = false;
                        eHasVideo.Text = "N";
                        rbHasCaseData = (RadioButtonList)fvAnecdoteCaseA_Data.FindControl("rbHasCaseData_INS");
                        eHasCaseData = (Label)fvAnecdoteCaseA_Data.FindControl("eHasCaseData_INS");
                        rbHasCaseData.SelectedIndex = 0;
                        eHasCaseData.Text = "Y";
                        cbHasAccReport = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasAccReport_INS");
                        eHasAccReport = (Label)fvAnecdoteCaseA_Data.FindControl("eHasAccReport_INS");
                        cbHasAccReport.Checked = false;
                        eHasAccReport.Text = "N";
                        cbIsExemption = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsExemption_INS");
                        eIsExemption = (Label)fvAnecdoteCaseA_Data.FindControl("eIsExemption_INS");
                        cbIsExemption.Checked = false;
                        eIsExemption.Text = "N";
                    }
                    break;
            }
        }

        /// <summary>
        /// 主檔確定編輯
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            Label eCaseNo_A = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_A_Edit");
            try
            {
                if ((eCaseNo_A != null) && (eCaseNo_A.Text.Trim() != ""))
                {
                    DateTime vModifyDate = DateTime.Today;
                    string vCaseNo_Temp = eCaseNo_A.Text.Trim(); //先取回單號
                    //寫入記錄
                    string vRecordNote = "異動肇事單主檔資料_" + vCaseNo_Temp + Environment.NewLine +
                                         "異動日期：" + vModifyDate.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                         "肇事單號：[" + vCaseNo_Temp + "]" + Environment.NewLine +
                                         "異動人工號：" + vLoginID + Environment.NewLine +
                                         "AnecdoteCase.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //複製舊資料到異動檔
                    string vHistoryNo = vModifyDate.Year.ToString("D4") + vModifyDate.Month.ToString("D2") + "MODA";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%'";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndex_Str = (vMaxNo.Trim() != "") ? vMaxNo.Replace(vHistoryNo, "").Trim() : "0";
                    int vNewIndex = Int32.Parse(vIndex_Str) + 1;
                    string vNewHistoryNo = vHistoryNo + vNewIndex.ToString("D4");
                    vSQLStr_Temp = "insert into AnecdoteCaseHistory " + Environment.NewLine +
                                   "       (HistoryNo, CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, DriverName, InsuMan, " + Environment.NewLine +
                                   "        AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, " + Environment.NewLine +
                                   "        PaidAmount, Penalty, PenaltyRatio, InsuAmount, IDCardNo, Birthday, Assumeday, TelephoneNo, Address, " + Environment.NewLine +
                                   "        PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, OutReportNo, PoliceUnit, PoliceName, HasVideo, " + Environment.NewLine +
                                   "        NoVideoReason, HasCaseData, ModifyType, ModifyMan, ModifyDate) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, DriverName, InsuMan, " + Environment.NewLine +
                                   "       AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, " + Environment.NewLine +
                                   "       PaidAmount, Penalty, PenaltyRatio, InsuAmount, IDCardNo, Birthday, Assumeday, TelephoneNo, Address, " + Environment.NewLine +
                                   "       PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, OutReportNo, PoliceUnit, PoliceName, HasVideo, " + Environment.NewLine +
                                   "       NoVideoReason, HasCaseData, 'EDIT', '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                   "  from AnecdoteCase " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //更新主檔資料
                    TextBox eBuildDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildDate_Edit");
                    TextBox eDepNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_Edit");
                    Label eHasInsurance = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsurance_Edit");
                    Label eDepName = (Label)fvAnecdoteCaseA_Data.FindControl("eDepName_Edit");
                    TextBox eCar_ID = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCar_ID_Edit");
                    TextBox eDriver = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_Edit");
                    Label eDriverName = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_Edit");
                    TextBox eInsuMan = (TextBox)fvAnecdoteCaseA_Data.FindControl("eInsuMan_Edit");
                    TextBox eAnecdotalResRatio = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAnecdotalResRatio_Edit");
                    Label eIsNoDeduction = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeduction_Edit");
                    TextBox eDeductionDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDeductionDate_Edit");
                    TextBox eRemark_A = (TextBox)fvAnecdoteCaseA_Data.FindControl("eRemark_A_Edit");
                    TextBox eCaseOccurrence = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseOccurrence_Edit");
                    Label eERPCouseNo = (Label)fvAnecdoteCaseA_Data.FindControl("eERPCouseNo_Edit");
                    Label eCaseClose = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseClose_Edit");
                    Label eIsExemption = (Label)fvAnecdoteCaseA_Data.FindControl("eIsExemption_Edit");
                    TextBox ePaidAmount = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePaidAmount_Edit");
                    TextBox ePenalty = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePenalty_Edit");
                    TextBox ePenaltyRatio = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePenaltyRatio_Edit");
                    TextBox eInsuAmount = (TextBox)fvAnecdoteCaseA_Data.FindControl("eInsuAmount_Edit");
                    TextBox eIDCardNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eIDCardNo_Edit");
                    TextBox eBirthday = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBirthday_Edit");
                    Label eAssumeday = (Label)fvAnecdoteCaseA_Data.FindControl("eAssumeday_Edit");
                    TextBox eTelephoneNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eTelephoneNo_Edit");
                    TextBox eAddress = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAddress_Edit");
                    TextBox ePersonDamage = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePersonDamage_Edit");
                    TextBox eCarDamage = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCarDamage_Edit");
                    TextBox eReportDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eReportDate_Edit");
                    TextBox eCaseDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseDate_Edit");
                    TextBox eCaseTime = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseTime_Edit");
                    TextBox eOutReportNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eOutReportNo_Edit");
                    TextBox ePoliceUnit = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePoliceUnit_Edit");
                    TextBox ePoliceName = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePoliceName_Edit");
                    Label eHasVideo = (Label)fvAnecdoteCaseA_Data.FindControl("eHasVideo_Edit");
                    TextBox eNoVideoReason = (TextBox)fvAnecdoteCaseA_Data.FindControl("eNoVideoReason_Edit");
                    Label eHasCaseData = (Label)fvAnecdoteCaseA_Data.FindControl("eHasCaseData_Edit");
                    TextBox eBuildMan = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildMan_Edit");

                    vSQLStr_Temp = "update [dbo].[AnecdoteCase] " + Environment.NewLine +
                                   "   set HasInsurance = @HasInsurance, DepNo = @DepNo, DepName = @DepName, BuildDate = @BuildDate, Car_ID = @Car_ID, " + Environment.NewLine +
                                   "       Driver = @Driver, DriverName = @DriverName, InsuMan = @InsuMan, AnecdotalResRatio = @AnecdotalResRatio, " + Environment.NewLine +
                                   "       IsNoDeduction = @IsNoDeduction, DeductionDate = @DeductionDate, Remark = @Remark, ModifyDate = GetDate(), " + Environment.NewLine +
                                   "       ModifyMan = @ModifyMan, CaseOccurrence = @CaseOccurrence, ERPCouseNo = @ERPCouseNo, CaseClose = @CaseClose, " + Environment.NewLine +
                                   "       IsExemption = @IsExemption, PaidAmount = @PaidAmount, Penalty = @Penalty, PenaltyRatio = @PenaltyRatio, " + Environment.NewLine +
                                   "       InsuAmount = @InsuAmount, IDCardNo = @IDCardNo, Birthday = @Birthday, Assumeday = @Assumeday, TelephoneNo = @TelephoneNo, " + Environment.NewLine +
                                   "       Address = @Address, PersonDamage = @PersonDamage, CarDamage = @CarDamage, ReportDate = @ReportDate, CaseDate = @CaseDate, " + Environment.NewLine +
                                   "       CaseTime = @CaseTime, OutReportNo = @OutReportNo, PoliceUnit = @PoliceUnit, PoliceName = @PoliceName, HasVideo = @HasVideo, " + Environment.NewLine +
                                   "       NoVideoReason = @NoVideoReason, HasCaseData = @HasCaseData, BuildMan = @BuildMan " + Environment.NewLine +
                                   " where CaseNo = @CaseNo";
                    sdsAnecdoteCaseA_Data.UpdateCommand = vSQLStr_Temp;
                    DateTime vTempDate;
                    double vTempFloat;
                    sdsAnecdoteCaseA_Data.UpdateParameters.Clear();
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("HasInsurance", DbType.Boolean, (eHasInsurance.Text.Trim() != "") ? eHasInsurance.Text.Trim() : "False"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("DepNo", DbType.String, eDepNo.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("DepName", DbType.String, eDepName.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("BuildDate", DbType.Date, DateTime.TryParse(eBuildDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Car_ID", DbType.String, eCar_ID.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Driver", DbType.String, eDriver.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("DriverName", DbType.String, eDriverName.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("InsuMan", DbType.String, eInsuMan.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("AnecdotalResRatio", DbType.Double, double.TryParse(eAnecdotalResRatio.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("IsNoDeduction", DbType.Boolean, (eIsNoDeduction.Text.Trim() != "") ? eIsNoDeduction.Text.Trim() : "False"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("DeductionDate", DbType.Date, DateTime.TryParse(eDeductionDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Remark", DbType.String, eRemark_A.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("CaseOccurrence", DbType.String, eCaseOccurrence.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("ERPCouseNo", DbType.String, eERPCouseNo.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("CaseClose", DbType.Boolean, (eCaseClose.Text.Trim() != "") ? eCaseClose.Text.Trim() : "False"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("IsExemption", DbType.String, (eIsExemption.Text.Trim() != "") ? eIsExemption.Text.Trim() : "N"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("PaidAmount", DbType.Double, double.TryParse(ePaidAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Penalty", DbType.Double, double.TryParse(ePenalty.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("PenaltyRatio", DbType.Double, double.TryParse(ePenaltyRatio.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("InsuAmount", DbType.Double, double.TryParse(eInsuAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("IDCardNo", DbType.String, eIDCardNo.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Birthday", DbType.Date, DateTime.TryParse(eBirthday.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Assumeday", DbType.Date, DateTime.TryParse(eAssumeday.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("TelephoneNo", DbType.String, eTelephoneNo.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("Address", DbType.String, eAddress.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("PersonDamage", DbType.String, ePersonDamage.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("CarDamage", DbType.String, eCarDamage.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("ReportDate", DbType.Date, DateTime.TryParse(eReportDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("CaseDate", DbType.Date, DateTime.TryParse(eCaseDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("CaseTime", DbType.String, eCaseTime.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("OutReportNo", DbType.String, eOutReportNo.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("PoliceUnit", DbType.String, ePoliceUnit.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("PoliceName", DbType.String, ePoliceName.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("HasVideo", DbType.String, (eHasVideo.Text.Trim() != "") ? eHasVideo.Text.Trim() : "N"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("NoVideoReson", DbType.String, eNoVideoReason.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("HasCaseData", DbType.String, (eHasCaseData.Text.Trim() != "") ? eHasCaseData.Text.Trim() : "X"));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("BuildMan", DbType.String, eBuildMan.Text.Trim()));
                    sdsAnecdoteCaseA_Data.UpdateParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp));
                    sdsAnecdoteCaseA_Data.Update();
                    gridAnecdoteCaseA_List.DataBind();
                    fvAnecdoteCaseA_Data.ChangeMode(FormViewMode.ReadOnly);
                }
            }
            catch (Exception eMessage)
            {
                eErrorMSG_A.Text = eMessage.Message;
                eErrorMSG_A.Visible = true;
            }
        }

        /// <summary>
        /// 編輯頁面選擇是否出險
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void cbHasInsurance_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbHasInsurance = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasInsurance_Edit");
            Label eHasInsurance = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsurance_Edit");
            if (cbHasInsurance != null)
            {
                eHasInsurance.Text = cbHasInsurance.Checked ? "True" : "False";
            }
        }

        /// <summary>
        /// 編輯頁面選擇是否和解
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void cbCaseClose_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbCaseClose = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbCaseClose_Edit");
            Label eCaseClose = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseClose_Edit");
            if (cbCaseClose != null)
            {
                eCaseClose.Text = cbCaseClose.Checked ? "True" : "False";
            }
        }

        /// <summary>
        /// 編輯頁面輸入建檔人
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eBuildMan_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eBuildMan = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildMan_Edit");
            Label eBuildManName = (Label)fvAnecdoteCaseA_Data.FindControl("eBuildManName_Edit");
            if (eBuildMan != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vBuildMan = eBuildMan.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vBuildMan + "' ";
                string vBuildManName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vBuildManName == "")
                {
                    vBuildManName = eBuildMan.Text.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vBuildManName + "' order by Assumeday DESC";
                    vBuildMan = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eBuildMan.Text = vBuildMan.Trim();
                eBuildManName.Text = vBuildManName.Trim();
            }
        }

        /// <summary>
        /// 編輯頁面選擇是否免扣精勤
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void cbIsNoDeduction_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsNoDeduction = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsNoDeduction_Edit");
            Label eIsNoDeduction = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeduction_Edit");
            if (cbIsNoDeduction != null)
            {
                eIsNoDeduction.Text = cbIsNoDeduction.Checked ? "True" : "False";
            }
        }

        /// <summary>
        /// 編輯畫面輸入車號
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eCar_ID_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eCar_ID = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCar_ID_Edit");
            TextBox eDepNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_Edit");
            Label eDepName = (Label)fvAnecdoteCaseA_Data.FindControl("eDepName_Edit");
            TextBox eDriver = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_Edit");
            Label eDriverName = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_Edit");
            TextBox eIDCardNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eIDCardNo_Edit");
            TextBox eBirthday = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBirthday_Edit");
            Label eAssumeday = (Label)fvAnecdoteCaseA_Data.FindControl("eAssumeday_Edit");
            TextBox eTelephoneNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eTelephoneNo_Edit");
            TextBox eAddress = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAddress_Edit");
            DateTime vBirthday;
            DateTime vAssumeday;
            if (eCar_ID != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCar_ID = eCar_ID.Text.Trim();
                string vSQLStr_Temp = "select isnull(a.Driver, '') + ',' + isnull(e.[Name], '') + ',' + isnull(a.CompanyNo, '') + ',' + isnull(d.[Name], '') as ResultStr " + Environment.NewLine +
                                      "  from Car_infoA a left join Department d on d.DepNo = a.CompanyNo " + Environment.NewLine +
                                      "                   left join Employee e on e.EmpNo = a.Driver " + Environment.NewLine +
                                      " where a.Car_ID = '" + vCar_ID + "' ";
                string vResultStr = PF.GetValue(vConnStr, vSQLStr_Temp, "ResultStr");
                string[] vaResult = vResultStr.Split(',');
                if (vaResult[0].ToString() != "")
                {
                    eDriver.Text = vaResult[0].ToString();
                    eDriverName.Text = vaResult[1].ToString();
                    vSQLStr_Temp = "select (IDCardNo + ',' + convert(varchar(10), Birthday, 111) + ',' + convert(varchar(10), Assumeday, 111) + ',' + " + Environment.NewLine +
                                   "       case when isnull(CellPhone, '') <> '' then CellPhone " + Environment.NewLine +
                                   "            when isnull(TelNo1, '') <> '' then TelNo1 " + Environment.NewLine +
                                   "            when isnull(TelNo2, '') <> '' then TelNo2 else '  ' end + ',' + Addr1) as ResultStr " + Environment.NewLine +
                                   "  from Employee " + Environment.NewLine +
                                   " where EmpNo = '" + vaResult[0].ToString() + "' ";
                    vResultStr = PF.GetValue(vConnStr, vSQLStr_Temp, "ResultStr");
                    string[] vaDriverData = vResultStr.Split(',');
                    eIDCardNo.Text = vaDriverData[0].ToString();
                    eBirthday.Text = DateTime.TryParse(vaDriverData[1].ToString(), out vBirthday) ? (vBirthday.Year - 3822).ToString() + '/' + vBirthday.ToString("MM/dd") : "";
                    eAssumeday.Text = DateTime.TryParse(vaDriverData[2].ToString(), out vAssumeday) ? (vAssumeday.Year - 3822).ToString() + '/' + vAssumeday.ToString("MM/dd") : "";
                    eTelephoneNo.Text = vaDriverData[3].ToString();
                    eAddress.Text = vaDriverData[4].ToString();
                }
                eDepNo.Text = vaResult[2].ToString();
                eDepName.Text = vaResult[3].ToString();
            }
        }

        /// <summary>
        /// 編輯頁面輸入所屬單位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_Edit");
            Label eDepName = (Label)fvAnecdoteCaseA_Data.FindControl("eDepName_Edit");
            if (eDepNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] = '" + vDepName + "' order by DepNo";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo.Text = vDepNo.Trim();
                eDepName.Text = vDepName.Trim();
            }
        }

        /// <summary>
        /// 編輯頁面輸入駕駛員資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eDriver_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eDriver = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_Edit");
            Label eDriverName = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_Edit");
            TextBox eIDCardNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eIDCardNo_Edit");
            TextBox eBirthday = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBirthday_Edit");
            Label eAssumeday = (Label)fvAnecdoteCaseA_Data.FindControl("eAssumeday_Edit");
            TextBox eTelephoneNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eTelephoneNo_Edit");
            TextBox eAddress = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAddress_Edit");
            DateTime vBirthday;
            DateTime vAssumeday;
            if (eDriver != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vDriver = eDriver.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver + "' ";
                string vDriverName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDriverName == "")
                {
                    vDriverName = eDriver.Text.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName + "' order by Assumeday DESC";
                    vDriver = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eDriver.Text = vDriver.Trim();
                eDriverName.Text = vDriverName.Trim();
                if (eDriver.Text.Trim() != "")
                {
                    vSQLStr_Temp = "select (IDCardNo + ',' + convert(varchar(10), Birthday, 111) + ',' + convert(varchar(10), Assumeday, 111) + ',' + " + Environment.NewLine +
                                   "       case when isnull(CellPhone, '') <> '' then CellPhone " + Environment.NewLine +
                                   "            when isnull(TelNo1, '') <> '' then TelNo1 " + Environment.NewLine +
                                   "            when isnull(TelNo2, '') <> '' then TelNo2 else '  ' end + ',' + Addr1) as ResultStr " + Environment.NewLine +
                                   "  from Employee " + Environment.NewLine +
                                   " where EmpNo = '" + vDriver.Trim() + "' ";
                    string vResultStr = PF.GetValue(vConnStr, vSQLStr_Temp, "ResultStr");
                    string[] vaDriverData = vResultStr.Split(',');
                    eIDCardNo.Text = vaDriverData[0].ToString();
                    eBirthday.Text = DateTime.TryParse(vaDriverData[1].ToString(), out vBirthday) ? (vBirthday.Year - 3822).ToString() + '/' + vBirthday.ToString("MM/dd") : "";
                    eAssumeday.Text = DateTime.TryParse(vaDriverData[2].ToString(), out vAssumeday) ? (vAssumeday.Year - 3822).ToString() + '/' + vAssumeday.ToString("MM/dd") : "";
                    eTelephoneNo.Text = vaDriverData[3].ToString();
                    eAddress.Text = vaDriverData[4].ToString();
                }
            }
        }

        protected void cbHasVideo_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbHasVideo = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasVideo_Edit");
            Label eHasVideo = (Label)fvAnecdoteCaseA_Data.FindControl("eHasVideo_Edit");
            if (cbHasVideo != null)
            {
                eHasVideo.Text = cbHasVideo.Checked ? "Y" : "N";
            }
        }

        protected void rbHasCaseData_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            RadioButtonList rbHasCaseData = (RadioButtonList)fvAnecdoteCaseA_Data.FindControl("rbHasCaseData_Edit");
            Label eHasCaseData = (Label)fvAnecdoteCaseA_Data.FindControl("eHasCaseData_Edit");
            if (rbHasCaseData != null)
            {
                eHasCaseData.Text = rbHasCaseData.SelectedValue.Trim();
            }
        }

        protected void cbHasAccReport_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbHasAccReport = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasAccReport_Edit");
            Label eHasAccReport = (Label)fvAnecdoteCaseA_Data.FindControl("eHasAccReport_Edit");
            if (cbHasAccReport != null)
            {
                eHasAccReport.Text = cbHasAccReport.Checked ? "Y" : "N";
            }
        }

        protected void cbIsExemption_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsExemption = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsExemption_Edit");
            Label eIsExemption = (Label)fvAnecdoteCaseA_Data.FindControl("eIsExemption_Edit");
            if (cbIsExemption != null)
            {
                eIsExemption.Text = cbIsExemption.Checked ? "Y" : "N";
            }
        }

        /// <summary>
        /// 主檔新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            TextBox eBuildDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildDate_INS");
            TextBox eDepNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_INS");
            try
            {
                if ((eBuildDate != null) && (eDepNo != null) && (eDepNo.Text.Trim() != ""))
                {
                    //產生新單號
                    string vCaseNo_INS;
                    string vCaseNo_FirstCode = (DateTime.Today.Year - 1911).ToString() + DateTime.Today.Month.ToString("D2") + "A";
                    string vSQLStr_INS = "select max(CaseNo) MaxNo from AnecdoteCase where CaseNo like '" + vCaseNo_FirstCode + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_INS, "MaxNo");
                    string vIndexStr = (vMaxNo.Trim() != "") ? vMaxNo.Replace(vCaseNo_FirstCode, "").Trim() : "0";
                    int vIndex = Int32.Parse(vIndexStr) + 1;
                    vCaseNo_INS = vCaseNo_FirstCode + vIndex.ToString("D4");
                    DateTime vToday = DateTime.Today;
                    //寫入記錄
                    string vRecordNote = "新增肇事單主檔資料_" + vCaseNo_INS + Environment.NewLine +
                                         "建檔日期：" + vToday.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                         "肇事單號：[" + vCaseNo_INS + "]" + Environment.NewLine +
                                         "建檔人工號：" + vLoginID + Environment.NewLine +
                                         "AnecdoteCase.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //取得欄位
                    Label eHasInsurance = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsurance_INS");
                    Label eDepName = (Label)fvAnecdoteCaseA_Data.FindControl("eDepName_INS");
                    TextBox eCar_ID = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCar_ID_INS");
                    TextBox eDriver = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_INS");
                    Label eDriverName = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_INS");
                    TextBox eInsuMan = (TextBox)fvAnecdoteCaseA_Data.FindControl("eInsuMan_INS");
                    TextBox eAnecdotalResRatio = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAnecdotalResRatio_INS");
                    Label eIsNoDeduction = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeduction_INS");
                    TextBox eDeductionDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDeductionDate_INS");
                    TextBox eRemark_A = (TextBox)fvAnecdoteCaseA_Data.FindControl("eRemark_A_INS");
                    TextBox eCaseOccurrence = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseOccurrence_INS");
                    Label eERPCouseNo = (Label)fvAnecdoteCaseA_Data.FindControl("eERPCouseNo_INS");
                    Label eCaseClose = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseClose_INS");
                    Label eIsExemption = (Label)fvAnecdoteCaseA_Data.FindControl("eIsExemption_INS");
                    TextBox ePaidAmount = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePaidAmount_INS");
                    TextBox ePenalty = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePenalty_INS");
                    TextBox ePenaltyRatio = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePenaltyRatio_INS");
                    TextBox eInsuAmount = (TextBox)fvAnecdoteCaseA_Data.FindControl("eInsuAmount_INS");
                    TextBox eIDCardNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eIDCardNo_INS");
                    TextBox eBirthday = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBirthday_INS");
                    Label eAssumeday = (Label)fvAnecdoteCaseA_Data.FindControl("eAssumeday_INS");
                    TextBox eTelephoneNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eTelephoneNo_INS");
                    TextBox eAddress = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAddress_INS");
                    TextBox ePersonDamage = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePersonDamage_INS");
                    TextBox eCarDamage = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCarDamage_INS");
                    TextBox eReportDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eReportDate_INS");
                    TextBox eCaseDate = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseDate_INS");
                    TextBox eCaseTime = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCaseTime_INS");
                    TextBox eOutReportNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eOutReportNo_INS");
                    TextBox ePoliceUnit = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePoliceUnit_INS");
                    TextBox ePoliceName = (TextBox)fvAnecdoteCaseA_Data.FindControl("ePoliceName_INS");
                    Label eHasVideo = (Label)fvAnecdoteCaseA_Data.FindControl("eHasVideo_INS");
                    TextBox eNoVideoReason = (TextBox)fvAnecdoteCaseA_Data.FindControl("eNoVideoReason_INS");
                    Label eHasCaseData = (Label)fvAnecdoteCaseA_Data.FindControl("eHasCaseData_INS");

                    //開始寫入
                    vSQLStr_INS = "insert into AnecdoteCase " + Environment.NewLine +
                                  "       (CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, DriverName, InsuMan, " + Environment.NewLine +
                                  "        AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, " + Environment.NewLine +
                                  "        PaidAmount, Penalty, PenaltyRatio, InsuAmount, IDCardNo, Birthday, Assumeday, TelephoneNo, Address, " + Environment.NewLine +
                                  "        PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, OutReportNo, PoliceUnit, PoliceName, HasVideo, " + Environment.NewLine +
                                  "        NoVideoReason, HasCaseData) " + Environment.NewLine +
                                  "values (@CaseNo, @HasInsurance, @DepNo, @DepName, @BuildDate, @BuildMan, @Car_ID, @Driver, @DriverName, @InsuMan, " + Environment.NewLine +
                                  "        @AnecdotalResRatio, @IsNoDeduction, @DeductionDate, @Remark, @CaseOccurrence, @ERPCouseNo, @CaseClose, @IsExemption, " + Environment.NewLine +
                                  "        @PaidAmount, @Penalty, @PenaltyRatio, @InsuAmount, @IDCardNo, @Birthday, @Assumeday, @TelephoneNo, @Address, " + Environment.NewLine +
                                  "        @PersonDamage, @CarDamage, @ReportDate, @CaseDate, @CaseTime, @OutReportNo, @PoliceUnit, @PoliceName, @HasVideo, " + Environment.NewLine +
                                  "        @NoVideoReason, @HasCaseData)";
                    sdsAnecdoteCaseA_Data.InsertCommand = vSQLStr_INS;
                    sdsAnecdoteCaseA_Data.InsertParameters.Clear();
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("HasInsurance", DbType.Boolean, (eHasInsurance.Text.Trim() != "") ? eHasInsurance.Text.Trim() : "False"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("DepName", DbType.String, (eDepName.Text.Trim() != "") ? eDepName.Text.Trim() : String.Empty));
                    DateTime vTempDate;
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, DateTime.TryParse(eBuildDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vLoginID));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Car_ID", DbType.String, (eCar_ID.Text.Trim() != "") ? eCar_ID.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Driver", DbType.String, (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("DriverName", DbType.String, (eDriverName.Text.Trim() != "") ? eDriverName.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("InsuMan", DbType.String, (eInsuMan.Text.Trim() != "") ? eInsuMan.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("AnecdotalResRatio", DbType.Double, (eAnecdotalResRatio.Text.Trim() != "") ? eAnecdotalResRatio.Text.Trim() : "0.0"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("IsNoDeduction", DbType.Boolean, (eIsNoDeduction.Text.Trim() != "") ? eIsNoDeduction.Text.Trim() : "False"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("DeductionDate", DbType.Date, DateTime.TryParse(eDeductionDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark_A.Text.Trim() != "") ? eRemark_A.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("CaseOccurrence", DbType.String, (eCaseOccurrence.Text.Trim() != "") ? eCaseOccurrence.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("ERPCouseNo", DbType.String, (eERPCouseNo.Text.Trim() != "") ? eERPCouseNo.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("CaseClose", DbType.Boolean, (eCaseClose.Text.Trim() != "") ? eCaseClose.Text.Trim() : "False"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("IsExemption", DbType.String, (eIsExemption.Text.Trim() != "") ? eIsExemption.Text.Trim() : "N"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("PaidAmount", DbType.Double, (ePaidAmount.Text.Trim() != "") ? ePaidAmount.Text.Trim() : "0.0"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Penalty", DbType.Double, (ePenalty.Text.Trim() != "") ? ePenalty.Text.Trim() : "0.0"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("PenaltyRatio", DbType.Double, (ePenaltyRatio.Text.Trim() != "") ? ePenaltyRatio.Text.Trim() : "0.0"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("InsuAmount", DbType.Double, (eInsuAmount.Text.Trim() != "") ? eInsuAmount.Text.Trim() : "0.0"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("IDCardNo", DbType.String, (eIDCardNo.Text.Trim() != "") ? eIDCardNo.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Birthday", DbType.Date, DateTime.TryParse(eBirthday.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Assumeday", DbType.Date, DateTime.TryParse(eAssumeday.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("TelePhoneNo", DbType.String, (eTelephoneNo.Text.Trim() != "") ? eTelephoneNo.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("Address", DbType.String, (eAddress.Text.Trim() != "") ? eAddress.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("PersonDamage", DbType.String, (ePersonDamage.Text.Trim() != "") ? ePersonDamage.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("CarDamage", DbType.String, (eCarDamage.Text.Trim() != "") ? eCarDamage.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("ReportDate", DbType.Date, DateTime.TryParse(eReportDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("CaseDate", DbType.Date, DateTime.TryParse(eCaseDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("CaseTime", DbType.String, (eCaseTime.Text.Trim() != "") ? eCaseTime.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("OutReportNo", DbType.String, (eOutReportNo.Text.Trim() != "") ? eOutReportNo.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("PoliceUnit", DbType.String, (ePoliceUnit.Text.Trim() != "") ? ePoliceUnit.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("PoliceName", DbType.String, (ePoliceName.Text.Trim() != "") ? ePoliceName.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("HasVideo", DbType.String, (eHasVideo.Text.Trim() != "") ? eHasVideo.Text.Trim() : "N"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("NoVideoReason", DbType.String, (eNoVideoReason.Text.Trim() != "") ? eNoVideoReason.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("HasCaseData", DbType.String, (eHasCaseData.Text.Trim() != "") ? eHasCaseData.Text.Trim() : "X"));
                    sdsAnecdoteCaseA_Data.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_INS));
                    sdsAnecdoteCaseA_Data.Insert();
                    fvAnecdoteCaseA_Data.ChangeMode(FormViewMode.ReadOnly);
                    gridAnecdoteCaseA_List.DataBind();
                    fvAnecdoteCaseA_Data.DataBind();
                }
                else
                {
                    eErrorMSG_A.Text = "站別不可空白！";
                    eErrorMSG_A.Visible = true;
                    eDepNo.Focus();
                }
            }
            catch (Exception eMessage)
            {
                eErrorMSG_A.Text = eMessage.Message;
                eErrorMSG_A.Visible = true;
            }
        }

        /// <summary>
        /// 新增畫面變更 "已出險" 選項
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void cbHasInsurance_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbHasInsurance = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasInsurance_INS");
            Label eHasInsurance = (Label)fvAnecdoteCaseA_Data.FindControl("eHasInsurance_INS");
            if (cbHasInsurance != null)
            {
                eHasInsurance.Text = cbHasInsurance.Checked ? "True" : "False";
            }
        }

        /// <summary>
        /// 新增畫面變更 "已和解" 選項
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void cbCaseClose_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbCaseClose = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbCaseClose_INS");
            Label eCaseClose = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseClose_INS");
            if (cbCaseClose != null)
            {
                eCaseClose.Text = cbCaseClose.Checked ? "True" : "False";
            }
        }

        /// <summary>
        /// 新增畫面建檔人變更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eBuildMan_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eBuildMan = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBuildMan_INS");
            Label eBuildManName = (Label)fvAnecdoteCaseA_Data.FindControl("eBuildManName_INS");
            if (eBuildMan != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vBuildMan = eBuildMan.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vBuildMan + "' ";
                string vBuildManName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vBuildManName == "")
                {
                    vBuildManName = eBuildMan.Text.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vBuildManName + "' order by Assumeday DESC";
                    vBuildMan = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eBuildMan.Text = vBuildMan.Trim();
                eBuildManName.Text = vBuildManName.Trim();
            }
        }

        /// <summary>
        /// 新增畫面變更 "不扣精勤" 選項
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void cbIsNoDeduction_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsNoDeduction = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsNoDeduction_INS");
            Label eIsNoDeduction = (Label)fvAnecdoteCaseA_Data.FindControl("eIsNoDeduction_INS");
            if (cbIsNoDeduction != null)
            {
                eIsNoDeduction.Text = cbIsNoDeduction.Checked ? "True" : "False";
            }
        }

        /// <summary>
        /// 新增畫面變更車號
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eCar_ID_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eCar_ID = (TextBox)fvAnecdoteCaseA_Data.FindControl("eCar_ID_INS");
            TextBox eDepNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_INS");
            Label eDepName = (Label)fvAnecdoteCaseA_Data.FindControl("eDepName_INS");
            TextBox eDriver = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_INS");
            Label eDriverName = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_INS");
            TextBox eIDCardNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eIDCardNo_INS");
            TextBox eBirthday = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBirthday_INS");
            Label eAssumeday = (Label)fvAnecdoteCaseA_Data.FindControl("eAssumeday_INS");
            TextBox eTelephoneNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eTelephoneNo_INS");
            TextBox eAddress = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAddress_INS");
            DateTime vBirthday;
            DateTime vAssumeday;
            if (eCar_ID != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCar_ID = eCar_ID.Text.Trim();
                string vSQLStr_Temp = "select isnull(a.Driver, '') + ',' + isnull(e.[Name], '') + ',' + isnull(a.CompanyNo, '') + ',' + isnull(d.[Name], '') as ResultStr " + Environment.NewLine +
                                      "  from Car_infoA a left join Department d on d.DepNo = a.CompanyNo " + Environment.NewLine +
                                      "                   left join Employee e on e.EmpNo = a.Driver " + Environment.NewLine +
                                      " where a.Car_ID = '" + vCar_ID + "' ";
                string vResultStr = PF.GetValue(vConnStr, vSQLStr_Temp, "ResultStr");
                string[] vaResult = vResultStr.Split(',');
                if (vaResult[0].ToString() != "")
                {
                    eDriver.Text = vaResult[0].ToString();
                    eDriverName.Text = vaResult[1].ToString();
                    vSQLStr_Temp = "select (IDCardNo + ',' + convert(varchar(10), Birthday, 111) + ',' + convert(varchar(10), Assumeday, 111) + ',' + " + Environment.NewLine +
                                   "       case when isnull(CellPhone, '') <> '' then CellPhone " + Environment.NewLine +
                                   "            when isnull(TelNo1, '') <> '' then TelNo1 " + Environment.NewLine +
                                   "            when isnull(TelNo2, '') <> '' then TelNo2 else '  ' end + ',' + Addr1) as ResultStr " + Environment.NewLine +
                                   "  from Employee " + Environment.NewLine +
                                   " where EmpNo = '" + vaResult[0].ToString() + "' ";
                    vResultStr = PF.GetValue(vConnStr, vSQLStr_Temp, "ResultStr");
                    string[] vaDriverData = vResultStr.Split(',');
                    eIDCardNo.Text = vaDriverData[0].ToString();
                    eBirthday.Text = DateTime.TryParse(vaDriverData[1].ToString(), out vBirthday) ? (vBirthday.Year - 3822).ToString() + '/' + vBirthday.ToString("MM/dd") : "";
                    eAssumeday.Text = DateTime.TryParse(vaDriverData[2].ToString(), out vAssumeday) ? (vAssumeday.Year - 3822).ToString() + '/' + vAssumeday.ToString("MM/dd") : "";
                    eTelephoneNo.Text = vaDriverData[3].ToString();
                    eAddress.Text = vaDriverData[4].ToString();
                }
                eDepNo.Text = vaResult[2].ToString();
                eDepName.Text = vaResult[3].ToString();
            }
        }

        /// <summary>
        /// 新增畫面變更所屬單位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDepNo_INS");
            Label eDepName = (Label)fvAnecdoteCaseA_Data.FindControl("eDepName_INS");
            if (eDepNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] = '" + vDepName + "' order by DepNo";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo.Text = vDepNo.Trim();
                eDepName.Text = vDepName.Trim();
            }
        }

        /// <summary>
        /// 新增畫面變更駕駛員
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eDriver_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eDriver = (TextBox)fvAnecdoteCaseA_Data.FindControl("eDriver_INS");
            Label eDriverName = (Label)fvAnecdoteCaseA_Data.FindControl("eDriverName_INS");
            TextBox eIDCardNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eIDCardNo_INS");
            TextBox eBirthday = (TextBox)fvAnecdoteCaseA_Data.FindControl("eBirthday_INS");
            Label eAssumeday = (Label)fvAnecdoteCaseA_Data.FindControl("eAssumeday_INS");
            TextBox eTelephoneNo = (TextBox)fvAnecdoteCaseA_Data.FindControl("eTelephoneNo_INS");
            TextBox eAddress = (TextBox)fvAnecdoteCaseA_Data.FindControl("eAddress_INS");
            DateTime vBirthday;
            DateTime vAssumeday;
            if (eDriver != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vDriver = eDriver.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver + "' ";
                string vDriverName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDriverName == "")
                {
                    vDriverName = eDriver.Text.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName + "' order by Assumeday DESC";
                    vDriver = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eDriver.Text = vDriver.Trim();
                eDriverName.Text = vDriverName.Trim();
                if (eDriver.Text.Trim() != "")
                {
                    vSQLStr_Temp = "select (IDCardNo + ',' + convert(varchar(10), Birthday, 111) + ',' + convert(varchar(10), Assumeday, 111) + ',' + " + Environment.NewLine +
                                   "       case when isnull(CellPhone, '') <> '' then CellPhone " + Environment.NewLine +
                                   "            when isnull(TelNo1, '') <> '' then TelNo1 " + Environment.NewLine +
                                   "            when isnull(TelNo2, '') <> '' then TelNo2 else '  ' end + ',' + Addr1) as ResultStr " + Environment.NewLine +
                                   "  from Employee " + Environment.NewLine +
                                   " where EmpNo = '" + vDriver.Trim() + "' ";
                    string vResultStr = PF.GetValue(vConnStr, vSQLStr_Temp, "ResultStr");
                    string[] vaDriverData = vResultStr.Split(',');
                    eIDCardNo.Text = vaDriverData[0].ToString();
                    eBirthday.Text = DateTime.TryParse(vaDriverData[1].ToString(), out vBirthday) ? (vBirthday.Year - 3822).ToString() + '/' + vBirthday.ToString("MM/dd") : "";
                    eAssumeday.Text = DateTime.TryParse(vaDriverData[2].ToString(), out vAssumeday) ? (vAssumeday.Year - 3822).ToString() + '/' + vAssumeday.ToString("MM/dd") : "";
                    eTelephoneNo.Text = vaDriverData[3].ToString();
                    eAddress.Text = vaDriverData[4].ToString();
                }
            }
        }

        protected void cbHasVideo_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbHasVideo = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasVideo_INS");
            Label eHasVideo = (Label)fvAnecdoteCaseA_Data.FindControl("eHasVideo_INS");
            if (cbHasVideo != null)
            {
                eHasVideo.Text = cbHasVideo.Checked ? "Y" : "N";
            }
        }

        protected void rbHasCaseData_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            RadioButtonList rbHasCaseData = (RadioButtonList)fvAnecdoteCaseA_Data.FindControl("rbHasCaseData_INS");
            Label eHasCaseData = (Label)fvAnecdoteCaseA_Data.FindControl("eHasCaseData_INS");
            if (rbHasCaseData != null)
            {
                eHasCaseData.Text = rbHasCaseData.SelectedValue.Trim();
            }
        }

        protected void cbHasAccReport_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbHasAccReport = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbHasAccReport_INS");
            Label eHasAccReport = (Label)fvAnecdoteCaseA_Data.FindControl("eHasAccReport_INS");
            if (cbHasAccReport != null)
            {
                eHasAccReport.Text = cbHasAccReport.Checked ? "Y" : "N";
            }
        }

        protected void cbIsExemption_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsExemption = (CheckBox)fvAnecdoteCaseA_Data.FindControl("cbIsExemption_INS");
            Label eIsExemption = (Label)fvAnecdoteCaseA_Data.FindControl("eIsExemption_INS");
            if (cbIsExemption != null)
            {
                eIsExemption.Text = cbIsExemption.Checked ? "Y" : "N";
            }
        }

        protected void gridAnecdoteCaseA_List_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridAnecdoteCaseA_List.PageIndex = e.NewPageIndex;
            gridAnecdoteCaseA_List.DataBind();
        }

        /// <summary>
        /// 主檔刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDel_Main_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            Label eCaseNo_List = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_A_List");
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
                    DateTime vDelDate = DateTime.Today;
                    //記錄操作
                    string vRecordNote = "刪除肇事單資料_" + vCaseNo_Temp + Environment.NewLine +
                                         "刪除日期：" + vDelDate.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                         "肇事單號：[" + vCaseNo_Temp + "]" + Environment.NewLine +
                                         "刪除人工號：" + vLoginID + Environment.NewLine +
                                         "AnecdoteCase.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //因為是刪主檔，所以要連 B 檔和 C 檔都一起刪，可是 C 檔沒有異動檔所以只針對 B 檔做異動
                    string vHistoryNo = vDelDate.Year.ToString("D4") + vDelDate.Month.ToString("D2") + "DELA";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndex_Str = vMaxNo.Replace(vHistoryNo, "").Trim();
                    int vIndex = (vIndex_Str != "") ? Int32.Parse(vIndex_Str) + 1 : 1;
                    string vNewHistoryNo = vHistoryNo + vIndex.ToString("D4");
                    vSQLStr_Temp = "select max(ItemsH) MaxItem from AnecdoteCaseBHistory where HistoryNo = '" + vNewHistoryNo + "' ";
                    string vMaxItemsH = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    string vNewItemsH = (vMaxItemsH.Trim() != "") ? Int32.Parse(vMaxItemsH).ToString("D4") : "0001";
                    string vNewHistoryNoItems = vNewHistoryNo.Trim() + vNewItemsH.Trim();
                    //複製 B 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseBHistory] " + Environment.NewLine +
                                   "            ([HistoryNo], [ItemsH], [HistoryNoItems], [CaseNo], [Items], [CaseNoItems], [Relationship], [RelCar_ID], " + Environment.NewLine +
                                   "             [EstimatedAmount], [ThirdInsurance], [CompInsurance], [DriverSharing], [CompanySharing], [CarDamageAMT], " + Environment.NewLine +
                                   "             [PersonDamageAMT], [RelationComp], [ReconciliationDate], [Remark], [ModifyType], [ModifyDate], [ModifyMan], " + Environment.NewLine +
                                   "             [PassengerInsu], [RelGender], [RelTelNo1], [RelTelNo2], [RelCarType], [RelPersonDamage], [RelCarDamage], " + Environment.NewLine +
                                   "             [RelNote])" + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', '" + vNewItemsH + "', '" + vNewHistoryNoItems + "', [CaseNo], [Items], [CaseNoItems], [Relationship], " + Environment.NewLine +
                                   "       [RelCar_ID], [EstimatedAmount], [ThirdInsurance], [CompInsurance], [DriverSharing], [CompanySharing], [CarDamageAMT], " + Environment.NewLine +
                                   "       [PersonDamageAMT], [RelationComp], [ReconciliationDate], [Remark], 'DELA', GetDate(), '" + vLoginID + "', [PassengerInsu], " + Environment.NewLine +
                                   "       [RelGender], [RelTelNo1], [RelTelNo2], [RelCarType], [RelPersonDamage], [RelCarDamage], [RelNote] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCaseB] " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //複製 A 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseHistory] " + Environment.NewLine +
                                   "            ([HistoryNo], [CaseNo], [HasInsurance], [DepNo], [DepName], [BuildDate], [BuildMan], [Car_ID], [Driver], " + Environment.NewLine +
                                   "             [DriverName], [InsuMan], [AnecdotalResRatio], [IsNoDeduction], [DeductionDate], [Remark], [ModifyType], " + Environment.NewLine +
                                   "             [ModifyDate], [ModifyMan], [CaseOccurrence], [ERPCouseNo], [CaseClose], [IsExemption], [PaidAmount], " + Environment.NewLine +
                                   "             [Penalty], [PenaltyRatio], [InsuAmount], [IDCardNo], [Birthday], [Assumeday], [TelephoneNo], [Address], " + Environment.NewLine +
                                   "             [PersonDamage], [CarDamage], [ReportDate], [CaseDate], [CaseTime], [OutReportNo], [PoliceUnit], [PoliceName], " + Environment.NewLine +
                                   "             [HasVideo], [NoVideoReason], [HasCaseData]) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', [CaseNo], [HasInsurance], [DepNo], [DepName], [BuildDate], [BuildMan], [Car_ID], [Driver], " + Environment.NewLine +
                                   "       [DriverName], [InsuMan], [AnecdotalResRatio], [IsNoDeduction], [DeductionDate], [Remark], 'DELA', GetDate(), '" + vLoginID + "', " + Environment.NewLine +
                                   "       [CaseOccurrence], [ERPCouseNo], [CaseClose], [IsExemption], [PaidAmount], [Penalty], [PenaltyRatio], [InsuAmount], [IDCardNo], " + Environment.NewLine +
                                   "       [Birthday], [Assumeday], [TelephoneNo], [Address], [PersonDamage], [CarDamage], [ReportDate], [CaseDate], [CaseTime], " + Environment.NewLine +
                                   "       [OutReportNo], [PoliceUnit], [PoliceName], [HasVideo], [NoVideoReason], [HasCaseData] " + Environment.NewLine +
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
                    CaseDataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Visible = true;
                    eErrorMSG_A.Text = eMessage.Message;
                }
            }
            else
            {
                eErrorMSG_A.Text = "請先選擇要刪除的肇事單";
                eErrorMSG_A.Visible = true;
            }
        }

        /// <summary>
        /// 明細 FormView 完成資料繫結
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void fvAnecdoteCaseB_Data_DataBound(object sender, EventArgs e)
        {
            string vTempDateURL = "";
            string vTempDateScript = "";
            DropDownList ddlRelCarType;
            Label eRelCarType;
            TextBox eReconciliationDate;
            switch (fvAnecdoteCaseB_Data.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    ddlRelCarType = (DropDownList)fvAnecdoteCaseB_Data.FindControl("ddlRelCarType_List");
                    if (ddlRelCarType != null)
                    {
                        eRelCarType = (Label)fvAnecdoteCaseB_Data.FindControl("eRelCarType_List");
                        if (eRelCarType.Text.Trim() == "")
                        {
                            eRelCarType.Text = "00";
                        }
                        ddlRelCarType.SelectedIndex = ddlRelCarType.Items.IndexOf(ddlRelCarType.Items.FindByValue(eRelCarType.Text.Trim()));
                    }
                    break;
                case FormViewMode.Edit:
                    eReconciliationDate = (TextBox)fvAnecdoteCaseB_Data.FindControl("eReconciliationDate_Edit");
                    if (eReconciliationDate != null)
                    {
                        vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eReconciliationDate.ClientID;
                        vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eReconciliationDate.Attributes["onClick"] = vTempDateScript;
                    }
                    ddlRelCarType = (DropDownList)fvAnecdoteCaseB_Data.FindControl("ddlRelCarType_Edit");
                    eRelCarType = (Label)fvAnecdoteCaseB_Data.FindControl("eRelCarType_Edit");
                    if (eRelCarType.Text.Trim() == "")
                    {
                        eRelCarType.Text = "00";
                    }
                    ddlRelCarType.SelectedIndex = ddlRelCarType.Items.IndexOf(ddlRelCarType.Items.FindByValue(eRelCarType.Text.Trim()));
                    break;
                case FormViewMode.Insert:
                    eReconciliationDate = (TextBox)fvAnecdoteCaseB_Data.FindControl("eReconciliationDate_INS");
                    if (eReconciliationDate != null)
                    {
                        vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eReconciliationDate.ClientID;
                        vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eReconciliationDate.Attributes["onClick"] = vTempDateScript;
                    }
                    ddlRelCarType = (DropDownList)fvAnecdoteCaseB_Data.FindControl("ddlRelCarType_INS");
                    eRelCarType = (Label)fvAnecdoteCaseB_Data.FindControl("eRelCarType_INS");
                    ddlRelCarType.SelectedIndex = 0;
                    eRelCarType.Text = "00";
                    break;
            }
        }

        /// <summary>
        /// 編輯明細畫面確定修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDetailOK_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eCaseNo_B = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseNo_B_Edit");
            if (eCaseNo_B != null)
            {
                string vSQLStr_Temp = "";
                int vTempINT;
                DateTime vTempDate;
                double vTempFloat;
                Label eItems = (Label)fvAnecdoteCaseB_Data.FindControl("eItems_Edit");
                Label eCaseNoItems = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseNoItems_Edit");
                string vCaseNo = eCaseNo_B.Text.Trim();
                string vItems = eItems.Text.Trim();
                string vCaseNoItems = eCaseNoItems.Text.Trim();
                try
                {
                    //寫入記錄
                    string vRecordNote = "修改肇事明細資料_" + vCaseNo + Environment.NewLine +
                                         "建檔日期：" + DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                         "肇事單號：[ " + vCaseNo + " ]" + Environment.NewLine +
                                         "明細項次：[ " + vItems + " ]" + Environment.NewLine +
                                         "異動人工號：" + vLoginID + Environment.NewLine +
                                         "AnecdoteCase.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //在主檔的異動記錄加入一筆資料
                    DateTime vModifyDate = DateTime.Today;
                    //先產生異動單號
                    string vHistoryNo = vModifyDate.Year.ToString("D4") + vModifyDate.Month.ToString("D2") + "MODB";
                    vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    Int32 vIndex;
                    string vIndex_Str = Int32.TryParse(vMaxNo.Replace(vHistoryNo, "").Trim(), out vIndex) ? (vIndex + 1).ToString("D4") : "0001";
                    string vNewHistoryNo = vHistoryNo + vIndex_Str;
                    //寫入主檔異動資料
                    vSQLStr_Temp = "INSERT INTO AnecdoteCaseHistory " + Environment.NewLine +
                                   "            (HistoryNo, CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, " + Environment.NewLine +
                                   "             DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, ModifyType, " + Environment.NewLine +
                                   "             ModifyDate, ModifyMan, CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, PaidAmount, " + Environment.NewLine +
                                   "             Penalty, PenaltyRatio, InsuAmount, IDCardNo, Birthday, Assumeday, TelephoneNo, Address, " + Environment.NewLine +
                                   "             PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, OutReportNo, PoliceUnit, PoliceName, " + Environment.NewLine +
                                   "             HasVideo, NoVideoReason, HasCaseData) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, " + Environment.NewLine +
                                   "       DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, 'MODB', GetDate(), '" + vLoginID + "', " + Environment.NewLine +
                                   "       CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, PaidAmount, Penalty, PenaltyRatio, InsuAmount, IDCardNo, " + Environment.NewLine +
                                   "       Birthday, Assumeday, TelephoneNo, Address, PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, " + Environment.NewLine +
                                   "       OutReportNo, PoliceUnit, PoliceName, HasVideo, NoVideoReason, HasCaseData " + Environment.NewLine +
                                   "  from AnecdoteCase " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //產生明細異動記錄
                    vSQLStr_Temp = "select max(Items) MaxItem from AnecdoteCaseBHistory where HistoryNo = '" + vNewHistoryNo + "' ";
                    vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    int vNewItems = (vMaxNo != "") ? Int32.Parse(vMaxNo) + 1 : 1;
                    string vNewItemsStr = vNewItems.ToString("D4");
                    //寫入明細異動資料
                    vSQLStr_Temp = "INSERT INTO AnecdoteCaseBHistory " + Environment.NewLine +
                                   "            (HistoryNo, ItemsH, HistoryNoItems, CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, " + Environment.NewLine +
                                   "             EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, " + Environment.NewLine +
                                   "             PersonDamageAMT, RelationComp, ReconciliationDate, Remark, ModifyType, ModifyDate, ModifyMan, " + Environment.NewLine +
                                   "             PassengerInsu, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, RelNote)" + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', '" + vNewItemsStr + "', '" + vNewHistoryNo + vNewItemsStr + "', " + Environment.NewLine +
                                   "       CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, " + Environment.NewLine +
                                   "       CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, Remark, 'MODB', GetDate(),  " + Environment.NewLine +
                                   "       '" + vLoginID + "', PassengerInsu, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, RelNote " + Environment.NewLine +
                                   "  from AnecdoteCaseB " + Environment.NewLine +
                                   " where CaseNoItems = '" + vCaseNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //開始更新明細資料
                    vSQLStr_Temp = "update AnecdoteCaseB " + Environment.NewLine +
                                   "   set Relationship = @Relationship, RelCar_ID = @RelCar_ID, EstimatedAmount = @EstimatedAmount, ThirdInsurance = @ThirdInsurance, " + Environment.NewLine +
                                   "       CompInsurance = @CompInsurance, DriverSharing = @DriverSharing, CompanySharing = @CompanySharing, CarDamageAMT = @CarDamageAMT, " + Environment.NewLine +
                                   "       PersonDamageAMT = @PersonDamageAMT, RelationComp = @RelationComp, ReconciliationDate = @ReconciliationDate, Remark = @Remark, " + Environment.NewLine +
                                   "       ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "', PassengerInsu = @PassengerInsu, RelGender = @RelGender, " + Environment.NewLine +
                                   "       RelTelNo1 = @RelTelNo1, RelTelNo2 = @RelTelNo2, RelCarType = @RelCarType, RelPersonDamage = @RelPersonDamage, " + Environment.NewLine +
                                   "       RelCarDamage = @RelCarDamage, RelNote = @RelNote " + Environment.NewLine +
                                   " where CaseNoItems = @CaseNoItems ";
                    sdsAnecdoteCaseB_Data.UpdateCommand = vSQLStr_Temp;
                    sdsAnecdoteCaseB_Data.UpdateParameters.Clear();
                    TextBox eRelationship = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelationship_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("Relationship", DbType.String, (eRelationship.Text.Trim() != "") ? eRelationship.Text.Trim() : String.Empty));
                    TextBox eRelCar_ID = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelCar_ID_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelCar_ID", DbType.String, (eRelCar_ID.Text.Trim() != "") ? eRelCar_ID.Text.Trim() : String.Empty));
                    TextBox eEstimatedAmount = (TextBox)fvAnecdoteCaseB_Data.FindControl("eEstimatedAmount_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("EstimatedAmount", DbType.Int32, Int32.TryParse(eEstimatedAmount.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eThirdInsurance = (TextBox)fvAnecdoteCaseB_Data.FindControl("eThirdInsurance_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("ThirdInsurance", DbType.Int32, Int32.TryParse(eThirdInsurance.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eCompInsurance = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCompInsurance_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("CompInsurance", DbType.Int32, Int32.TryParse(eCompInsurance.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eDriverSharing = (TextBox)fvAnecdoteCaseB_Data.FindControl("eDriverSharing_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("DriverSharing", DbType.Int32, Int32.TryParse(eDriverSharing.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eCompanySharing = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCompanySharing_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("CompanySharing", DbType.Int32, Int32.TryParse(eCompanySharing.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eCarDamageAMT = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCarDamageAMT_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("CarDamageAMT", DbType.Int32, Int32.TryParse(eCarDamageAMT.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox ePersonDamageAMT = (TextBox)fvAnecdoteCaseB_Data.FindControl("ePersonDamageAMT_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("PersonDamageAMT", DbType.Int32, Int32.TryParse(ePersonDamageAMT.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eRelationComp = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelationComp_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelationComp", DbType.Int32, Int32.TryParse(eRelationComp.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eReconciliationDate = (TextBox)fvAnecdoteCaseB_Data.FindControl("eReconciliationDate_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("ReconciliationDate", DbType.Date, DateTime.TryParse(eReconciliationDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    TextBox ePassengerInsu = (TextBox)fvAnecdoteCaseB_Data.FindControl("ePassengerInsu_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("PassengerInsu", DbType.Double, double.TryParse(ePassengerInsu.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    TextBox eRemark = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRemark_B_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark.Text.Trim() != "") ? eRemark.Text.Trim() : String.Empty));
                    TextBox eRelGender = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelGender_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelGender", DbType.String, (eRelGender.Text.Trim() != "") ? eRelGender.Text.Trim() : String.Empty));
                    TextBox eRelTelNo1 = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelTelNo1_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelTelNo1", DbType.String, (eRelTelNo1.Text.Trim() != "") ? eRelTelNo1.Text.Trim() : String.Empty));
                    TextBox eRelTelNo2 = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelTelNo2_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelTelNo2", DbType.String, (eRelTelNo2.Text.Trim() != "") ? eRelTelNo2.Text.Trim() : String.Empty));
                    Label eRelCarType = (Label)fvAnecdoteCaseB_Data.FindControl("eRelCarType_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelCarType", DbType.String, (eRelCarType.Text.Trim() != "") ? eRelCarType.Text.Trim() : "00"));
                    TextBox eRelPersonDamage = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelPersonDamage_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelPersonDamage", DbType.String, (eRelPersonDamage.Text.Trim() != "") ? eRelPersonDamage.Text.Trim() : String.Empty));
                    TextBox eRelCarDamage = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelCarDamage_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelCarDamage", DbType.String, (eRelCarDamage.Text.Trim() != "") ? eRelCarDamage.Text.Trim() : String.Empty));
                    TextBox eRelNote = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelNote_Edit");
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("RelNote", DbType.String, (eRelNote.Text.Trim() != "") ? eRelNote.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.UpdateParameters.Add(new Parameter("CaseNoItems", DbType.String, vCaseNoItems));
                    sdsAnecdoteCaseB_Data.Update();
                    fvAnecdoteCaseB_Data.ChangeMode(FormViewMode.ReadOnly);
                    gridAnecdoteCaseB_List.DataBind();
                    fvAnecdoteCaseB_Data.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        /// <summary>
        /// 編輯明細畫面選擇車種
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlRelCarType_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlRelCarType = (DropDownList)fvAnecdoteCaseB_Data.FindControl("ddlRelCarType_Edit");
            Label eRelCarType = (Label)fvAnecdoteCaseB_Data.FindControl("eRelCarType_Edit");
            if (ddlRelCarType != null)
            {
                eRelCarType.Text = ddlRelCarType.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 新增明細畫面確定新增
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDetailOK_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            try
            {
                Label eCaseNo_B = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseNo_B_INS");
                DateTime vTempDate;
                double vTempFloat;
                if (eCaseNo_B != null)
                {
                    string vSQLStr_INSB = "select max(Items) MaxItems from AnecdoteCaseB where CaseNo = '" + eCaseNo_B.Text.Trim() + "' ";
                    string vMaxItems = PF.GetValue(vConnStr, vSQLStr_INSB, "MaxItems");
                    int vTempINT;
                    string vNewItems = Int32.TryParse(vMaxItems, out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    string vCaseNoItems = eCaseNo_B.Text.Trim() + vNewItems;
                    //寫入記錄
                    string vRecordNote = "新增肇事單明細資料_" + eCaseNo_B.Text.Trim() + Environment.NewLine +
                                         "建檔日期：" + DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                         "肇事單號：[" + eCaseNo_B.Text.Trim() + "]" + Environment.NewLine +
                                         "明細項次：[" + vNewItems + "]" + Environment.NewLine +
                                         "建檔人工號：" + vLoginID + Environment.NewLine +
                                         "AnecdoteCase.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    vSQLStr_INSB = "insert into AnecdoteCaseB " + Environment.NewLine +
                                   "       (CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, " + Environment.NewLine +
                                   "        CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, PassengerInsu, Remark, RelGender, RelTelNo1, RelTelNo2, RelCarType, " + Environment.NewLine +
                                   "        RelPersonDamage, RelCarDamage, RelNote)" + Environment.NewLine +
                                   "values (@CaseNo, @Items, @CaseNoItems, @Relationship, @RelCar_ID, @EstimatedAmount, @ThirdInsurance, @CompInsurance, @DriverSharing, @CompanySharing, " + Environment.NewLine +
                                   "        @CarDamageAMT, @PersonDamageAMT, @RelationComp, @ReconciliationDate, @PassengerInsu, @Remark, @RelGender, @RelTelNo1, @RelTelNo2, @RelCarType, " + Environment.NewLine +
                                   "        @RelPersonDamage, @RelCarDamage, @RelNote)";
                    sdsAnecdoteCaseB_Data.InsertCommand = vSQLStr_INSB;
                    sdsAnecdoteCaseB_Data.InsertParameters.Clear();
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CaseNo", DbType.String, eCaseNo_B.Text.Trim()));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("Items", DbType.String, vNewItems));
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CaseNoItems", DbType.String, vCaseNoItems));
                    TextBox eRelationship = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelationship_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("Relationship", DbType.String, (eRelationship.Text.Trim() != "") ? eRelationship.Text.Trim() : String.Empty));
                    TextBox eRelCar_ID = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelCar_ID_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelCar_ID", DbType.String, (eRelCar_ID.Text.Trim() != "") ? eRelCar_ID.Text.Trim() : String.Empty));
                    TextBox eEstimatedAmount = (TextBox)fvAnecdoteCaseB_Data.FindControl("eEstimatedAmount_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("EstimatedAmount", DbType.Int32, Int32.TryParse(eEstimatedAmount.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eThirdInsurance = (TextBox)fvAnecdoteCaseB_Data.FindControl("eThirdInsurance_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("ThirdInsurance", DbType.Int32, Int32.TryParse(eThirdInsurance.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eCompInsurance = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCompInsurance_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CompInsurance", DbType.Int32, Int32.TryParse(eCompInsurance.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eDriverSharing = (TextBox)fvAnecdoteCaseB_Data.FindControl("eDriverSharing_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("DriverSharing", DbType.Int32, Int32.TryParse(eDriverSharing.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eCompanySharing = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCompanySharing_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CompanySharing", DbType.Int32, Int32.TryParse(eCompanySharing.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eCarDamageAMT = (TextBox)fvAnecdoteCaseB_Data.FindControl("eCarDamageAMT_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("CarDamageAMT", DbType.Int32, Int32.TryParse(eCarDamageAMT.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox ePersonDamageAMT = (TextBox)fvAnecdoteCaseB_Data.FindControl("ePersonDamageAMT_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("PersonDamageAMT", DbType.Int32, Int32.TryParse(ePersonDamageAMT.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eRelationComp = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelationComp_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelationComp", DbType.Int32, Int32.TryParse(eRelationComp.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                    TextBox eReconciliationDate = (TextBox)fvAnecdoteCaseB_Data.FindControl("eReconciliationDate_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("ReconciliationDate", DbType.Date, DateTime.TryParse(eReconciliationDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    TextBox ePassengerInsu = (TextBox)fvAnecdoteCaseB_Data.FindControl("ePassengerInsu_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("PassengerInsu", DbType.Double, double.TryParse(ePassengerInsu.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    TextBox eRemark = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRemark_B_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark.Text.Trim() != "") ? eRemark.Text.Trim() : String.Empty));
                    TextBox eRelGender = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelGender_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelGender", DbType.String, (eRelGender.Text.Trim() != "") ? eRelGender.Text.Trim() : String.Empty));
                    TextBox eRelTelNo1 = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelTelNo1_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelTelNo1", DbType.String, (eRelTelNo1.Text.Trim() != "") ? eRelTelNo1.Text.Trim() : String.Empty));
                    TextBox eRelTelNo2 = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelTelNo2_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelTelNo2", DbType.String, (eRelTelNo2.Text.Trim() != "") ? eRelTelNo2.Text.Trim() : String.Empty));
                    Label eRelCarType = (Label)fvAnecdoteCaseB_Data.FindControl("eRelCarType_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelCarType", DbType.String, (eRelCarType.Text.Trim() != "") ? eRelCarType.Text.Trim() : "00"));
                    TextBox eRelPersonDamage = (TextBox)fvAnecdoteCaseB_Data.FindControl("RelPersonDamage");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelPersonDamage", DbType.String, (eRelPersonDamage.Text.Trim() != "") ? eRelPersonDamage.Text.Trim() : String.Empty));
                    TextBox eRelCarDamage = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelCarDamage_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelCarDamage", DbType.String, (eRelCarDamage.Text.Trim() != "") ? eRelCarDamage.Text.Trim() : String.Empty));
                    TextBox eRelNote = (TextBox)fvAnecdoteCaseB_Data.FindControl("eRelNote_INS");
                    sdsAnecdoteCaseB_Data.InsertParameters.Add(new Parameter("RelNote", DbType.String, (eRelNote.Text.Trim() != "") ? eRelNote.Text.Trim() : String.Empty));
                    sdsAnecdoteCaseB_Data.Insert();
                    fvAnecdoteCaseB_Data.ChangeMode(FormViewMode.ReadOnly);
                    gridAnecdoteCaseB_List.DataBind();
                    fvAnecdoteCaseB_Data.DataBind();
                }
            }
            catch (Exception eMessage)
            {
                eErrorMSG_B.Text = eMessage.Message;
                eErrorMSG_B.Visible = true;
            }
        }

        /// <summary>
        /// 新增明細畫面選擇車種
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlRelCarType_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlRelCarType = (DropDownList)fvAnecdoteCaseB_Data.FindControl("ddlRelCarType_INS");
            Label eRelCarType = (Label)fvAnecdoteCaseB_Data.FindControl("eRelCarType_INS");
            if (ddlRelCarType != null)
            {
                eRelCarType.Text = ddlRelCarType.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 明細刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDel_Detail_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            try
            {
                Label eCaseNoItems = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseNoItems_List");
                Label eCaseNo = (Label)fvAnecdoteCaseB_Data.FindControl("eCaseNo_B_List");
                if ((eCaseNoItems != null) && (eCaseNoItems.Text.Trim() != ""))
                {
                    //找不到明細序號或是序號是空白就不進行刪除的動作
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    //雖然是刪除明細，還是要在主檔的異動記錄加入一筆資料
                    DateTime vModifyDate = DateTime.Today;
                    //先取回異動記錄序號的最大值
                    string vHistoryNo = vModifyDate.Year.ToString("D4") + vModifyDate.Month.ToString("D2") + "DELB";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndexStr = (vMaxNo.Replace(vHistoryNo, "").Trim() != "") ? vMaxNo.Replace(vHistoryNo, "").Trim() : "0";
                    int vIndex = Int32.Parse(vIndexStr) + 1;
                    //產生新的異動記錄序號
                    string vNewHistoryNo = vHistoryNo + vIndex.ToString("D4");
                    string vCaseNo_Temp = eCaseNo.Text.Trim();
                    //寫入主檔異動記錄
                    vSQLStr_Temp = "INSERT INTO AnecdoteCaseHistory " + Environment.NewLine +
                                   "            (HistoryNo, CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, " + Environment.NewLine +
                                   "             DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, ModifyType, " + Environment.NewLine +
                                   "             ModifyDate, ModifyMan, CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, PaidAmount, " + Environment.NewLine +
                                   "             Penalty, PenaltyRatio, InsuAmount, IDCardNo, Birthday, Assumeday, TelephoneNo, Address, " + Environment.NewLine +
                                   "             PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, OutReportNo, PoliceUnit, PoliceName, " + Environment.NewLine +
                                   "             HasVideo, NoVideoReason, HasCaseData) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, " + Environment.NewLine +
                                   "       DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, 'EDITA', GetDate(), '" + vLoginID + "', " + Environment.NewLine +
                                   "       CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, PaidAmount, Penalty, PenaltyRatio, InsuAmount, IDCardNo, " + Environment.NewLine +
                                   "       Birthday, Assumeday, TelephoneNo, Address, PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, " + Environment.NewLine +
                                   "       OutReportNo, PoliceUnit, PoliceName, HasVideo, NoVideoReason, HasCaseData " + Environment.NewLine +
                                   "  from AnecdoteCase " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入明細異動記錄
                    //先取回異動記錄序號的最大值
                    vSQLStr_Temp = "select max(ItemsH) MaxItems from AnecdoteCaseBHistory where HistoryNo = '" + vNewHistoryNo + "' ";
                    vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems");
                    int vItems = (vMaxNo != "") ? Int32.Parse(vMaxNo) + 1 : 1;
                    //複製明細原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO AnecdoteCaseBHistory " + Environment.NewLine +
                                   "            (HistoryNo, ItemsH, HistoryNoItems, CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, " + Environment.NewLine +
                                   "             EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, " + Environment.NewLine +
                                   "             PersonDamageAMT, RelationComp, ReconciliationDate, Remark, ModifyType, ModifyDate, ModifyMan, " + Environment.NewLine +
                                   "             PassengerInsu, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, RelNote)" + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', '" + vItems.ToString("D4") + "', '" + vNewHistoryNo + vItems.ToString("D4") + "', " + Environment.NewLine +
                                   "       CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, " + Environment.NewLine +
                                   "       CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, Remark, 'DELB', GetDate(),  " + Environment.NewLine +
                                   "       '" + vLoginID + "', PassengerInsu, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, RelNote " + Environment.NewLine +
                                   "  from AnecdoteCaseB " + Environment.NewLine +
                                   " where CaseNoItems = '" + eCaseNoItems.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //開始刪除明細
                    string vDelStr_B = "delete from AnecdoteCaseB where CaseNoItems = @CaseNoItems";
                    sdsAnecdoteCaseB_Data.DeleteCommand = vDelStr_B;
                    sdsAnecdoteCaseB_Data.DeleteParameters.Clear();
                    sdsAnecdoteCaseB_Data.DeleteParameters.Add(new Parameter("CaseNoItems", DbType.String, eCaseNoItems.Text.Trim()));
                    sdsAnecdoteCaseB_Data.Delete();
                    gridAnecdoteCaseB_List.DataBind();
                    fvAnecdoteCaseB_Data.DataBind();
                }
            }
            catch (Exception eMessage)
            {
                eErrorMSG_B.Text = eMessage.Message;
                eErrorMSG_B.Visible = true;
            }
        }

        /// <summary>
        /// 取回審議通知需要的資料
        /// </summary>
        /// <returns></returns>
        private DataTable GetReportData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DataTable dtResult = new DataTable();
            Label eCaseNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_List");
            string vResultStr = (eCaseNo_Temp != null) ?
                                "select DepName, Car_ID, CaseOccurrence, Driver, DriverName, IsExemption, PaidAmount, InsuAmount, PenaltyRatio, Penalty " + Environment.NewLine +
                                "  from AnecdoteCase " + Environment.NewLine +
                                " where CaseNo = @CaseNo " : "";
            if (vResultStr != "")
            {
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daTemp = new SqlDataAdapter(vResultStr, connTemp);
                    daTemp.SelectCommand.Parameters.Clear();
                    daTemp.SelectCommand.Parameters.Add(new SqlParameter("CaseNo", eCaseNo_Temp.Text.Trim()));
                    connTemp.Open();
                    daTemp.Fill(dtResult);
                }
            }
            return dtResult;
        }

        /// <summary>
        /// 列印審議通知
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintNote_Click(object sender, EventArgs e)
        {
            DataTable dtReport = GetReportData();
            ReportDataSource rdsPrint = new ReportDataSource("AnecdoteCaseP", dtReport);

            rvPrint.LocalReport.DataSources.Clear();
            rvPrint.LocalReport.ReportPath = @"Report\AnecdoteCaseP2.rdlc";
            rvPrint.LocalReport.DataSources.Add(rdsPrint);
            rvPrint.LocalReport.Refresh();
            plPrint.Visible = true;
            plMainDataShow.Visible = false;
            plDetailDataShow.Visible = false;
            plSearch.Visible = false;
        }

        /// <summary>
        /// 列印肇事報告表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintReport_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            Label eCaseNo = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_A_List");
            if (eCaseNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSQLStr_Print = "select b.CaseNoItems, a.DepName, convert(varchar(10), a.ReportDate, 111) ReportDate, a.Driver, a.DriverName, a.IDCardNo, " + Environment.NewLine +
                                       "       convert(varchar(10), a.Birthday, 111) Birthday, convert(varchar(10), a.Assumeday, 111) Assumeday, a.Car_ID, " + Environment.NewLine +
                                       "	   a.TelephoneNo, a.[Address], a.PersonDamage, a.CarDamage, b.Relationship, b.RelGender, b.RelTelNo1,b.RelTelNo2, " + Environment.NewLine +
                                       "	   case when b.RelCarType = '00' then '' when b.RelCarType = '01' then '行人 (含輪椅)' " + Environment.NewLine +
                                       "            when b.RelCarType = '02' then '腳踏車' when b.RelCarType = '03' then '機車 (含電動二輪)' " + Environment.NewLine +
                                       "            when b.RelCarType = '04' then '自小客貨' when b.RelCarType = '05' then '營小客貨' " + Environment.NewLine +
                                       "            when b.RelCarType = '06' then '大客車' else '其他' end RelCarType, " + Environment.NewLine +
                                       "       b.RelCar_ID, b.RelPersonDamage, b.RelCarDamage, convert(varchar(10), a.CaseDate, 111) CaseDate, a.CaseTime, " + Environment.NewLine +
                                       "	   a.OutReportNo, a.CasePosition, a.PoliceUnit, a.PoliceName, a.CaseOccurrence, b.RelNote, " + Environment.NewLine +
                                       "	   a.HasVideo, a.NoVideoReason, a.HasCaseData, a.HasAccReport " + Environment.NewLine +
                                       "  from AnecdoteCase a left join AnecdoteCaseB b on b.CaseNo = a.CaseNo " + Environment.NewLine +
                                       " where a.CaseNo = @CaseNo ";
                using (SqlConnection connPrint=new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daPrint = new SqlDataAdapter(vSQLStr_Print, connPrint);
                    daPrint.SelectCommand.Parameters.Clear();
                    daPrint.SelectCommand.Parameters.Add(new SqlParameter("CaseNo", eCaseNo.Text.Trim()));
                    connPrint.Open();
                    DataTable dtPrint = new DataTable();
                    daPrint.Fill(dtPrint);
                    ReportDataSource rdsPrint = new ReportDataSource("AnecdoteReportP", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\AnecdoteReportP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plMainDataShow.Visible = false;
                    plDetailDataShow.Visible = false;
                    plSearch.Visible = false;
                }
            }
        }

        /// <summary>
        /// ERP同步
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbERPSync_Click(object sender, EventArgs e)
        {
            string vSQLStr_Temp = "";
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            Label eCaseNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_A_List");
            if (eCaseNo_Temp != null)
            {
                Label eDepNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eDepNo_List");
                Label eBuildDate_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eBuildDate_List");
                Label eCarID_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eCar_ID_List");
                Label eBuildMan_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eBuildMan_List");
                Label eDriver_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eDriver_List");
                TextBox eRemark_Temp = (TextBox)fvAnecdoteCaseA_Data.FindControl("eRemark_A_List");
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
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 取消同步
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelERP_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            Label eCaseNo_Temp = (Label)fvAnecdoteCaseA_Data.FindControl("eCaseNo_A_List");
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
                    using (SqlConnection connUpdate = new SqlConnection(vConnStr))
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
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 結束報表預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plSearch.Visible = true;
            plMainDataShow.Visible = true;
            plDetailDataShow.Visible = true;
        }
    }
}