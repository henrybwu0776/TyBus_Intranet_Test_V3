using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;


namespace TyBus_Intranet_Test_V3
{
    public partial class LinesBounds : Page
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
        private string vSQLStr_Main = "";
        private string[] vStrArray = { "02", "03", "04", "99" }; //在轉入表一和國道台鐵轉乘優惠資料時要排除不轉入的卡種

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

                UnobtrusiveValidationMode = UnobtrusiveValidationMode.None;

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        eImportYear.Text = (DateTime.Now.Year - 1911).ToString();
                        ddlImportMonth.SelectedIndex = ((DateTime.Now.Month - 2) < 0) ? 12 : DateTime.Now.Month - 2;
                        plShowData.Visible = false;
                        vSQLStr_Main = "select Content from Sysflag where FormName = 'unSysFlag' and ControlItem = 'a_TicketPrice_City' ";
                        eTicketPrice_City.Text = PF.GetValue(vConnStr, vSQLStr_Main, "Content");
                        vSQLStr_Main = "select Content from Sysflag where FormName = 'unSysFlag' and ControlItem = 'a_TicketPrice_Highway' ";
                        eTicketPrice_Highway.Text = PF.GetValue(vConnStr, vSQLStr_Main, "Content");
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
        /// 建立轉入表三 CSV 檔時的中繼資料表 (LINQ 用)
        /// </summary>
        /// <returns></returns>
        private DataTable CreateDataForLinq_List3()
        {
            DataTable dtTemp = new DataTable();
            DataColumn dcTemp = new DataColumn();
            dcTemp.ColumnName = "IndexNo";
            dcTemp.DataType = Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "CalYM";
            dcTemp.DataType = Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "DepNo";
            dcTemp.DataType = Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "TicketLineNo";
            dcTemp.DataType = Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "TicketType";
            dcTemp.DataType = Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "TicketPCount";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "TotalAmount";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "RealIncome";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_8KM";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Limited";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Student";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Social";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_MRT";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Area";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Highway";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Others";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Dift";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            return dtTemp;
        }

        /// <summary>
        /// 建立轉入表三 CSV 檔時的中繼資料表
        /// </summary>
        /// <returns></returns>
        private DataTable CreateDataTable()
        {
            DataTable dtTemp = new DataTable();
            DataColumn dcTemp = new DataColumn();
            dcTemp.ColumnName = "IndexNo";
            dcTemp.DataType = System.Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "CalYM";
            dcTemp.DataType = System.Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "TicketType";
            dcTemp.DataType = System.Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "TicketLineNo";
            dcTemp.DataType = System.Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "LinesGovNo";
            dcTemp.DataType = System.Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "TicketPCount";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "TotalAmount";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "RealIncome";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_8KM";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Limited";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Student";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Social";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_MRT";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Area";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Highway";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Others";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Dift";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);


            dcTemp = new DataColumn();
            dcTemp.ColumnName = "BuMan";
            dcTemp.DataType = System.Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "BuDate";
            dcTemp.DataType = typeof(DateTime);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "ModifyMan";
            dcTemp.DataType = System.Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "ModifyDate";
            dcTemp.DataType = typeof(DateTime);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "DepNo";
            dcTemp.DataType = System.Type.GetType("System.String");
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "DistanceCost_Car";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Ticket_D";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "TicketIncome";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "CashIncome";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "DeductAmount";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_TWTrip";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_NTOld";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_TPMPass";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_NTBus";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Student2";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_BusTran";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Trial";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Trip501";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Trip502";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_TrialPre";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Highway2";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "Subsidy_Old2";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "DailyCostByCar";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "DailyIncomeByCar";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "DepCarCount";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "INC_ZZ";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "COS301";
            dcTemp.DataType = typeof(double);
            dtTemp.Columns.Add(dcTemp);

            dcTemp = new DataColumn();
            dcTemp.ColumnName = "MonthDays";
            dcTemp.DataType = typeof(int);
            dtTemp.Columns.Add(dcTemp);

            return dtTemp;
        }

        /// <summary>
        /// 轉入刷卡資料
        /// </summary>
        private void ImportFromExcel_CardIn()
        {
            if (fuExcel.FileName=="")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr=="")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";
                int vSheetCount = 0;
                int vTempINT = 0;
                Double vTempFloat = 0.0;
                string vErrorStr = "";
                DateTime vDate_S = (Int32.Parse(eImportYear.Text.Trim()) < 1911) ?
                                   new DateTime(Int32.Parse(eImportYear.Text.Trim()) + 1911, Int32.Parse(ddlImportMonth.SelectedValue.Trim()), 1) :
                                   new DateTime(Int32.Parse(eImportYear.Text.Trim()), Int32.Parse(ddlImportMonth.SelectedValue.Trim()), 1);
                DateTime vDate_E = vDate_S.AddMonths(1).AddDays(-1);

                string vIndexNo = "";
                string vCalYM = "";
                string vLinesNo = "";
                string vTicketLineNo = "";
                string vLinesGovNo = "";

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式

                        break;

                    case ".xlsx": //新版 EXCEL (2010 以後) 格式

                        break;
                }
            }
        }

        /// <summary>
        /// 轉入表一資料
        /// </summary>
        private void ImportFromExcel_List1()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";
                int vSheetCount = 0;
                int vTempINT = 0;
                Double vTempFloat = 0.0;
                string vErrorStr = "";
                DateTime vDate_S = (Int32.Parse(eImportYear.Text.Trim()) < 1911) ?
                                   new DateTime(Int32.Parse(eImportYear.Text.Trim()) + 1911, Int32.Parse(ddlImportMonth.SelectedValue.Trim()), 1) :
                                   new DateTime(Int32.Parse(eImportYear.Text.Trim()), Int32.Parse(ddlImportMonth.SelectedValue.Trim()), 1);
                DateTime vDate_E = vDate_S.AddMonths(1).AddDays(-1);

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vLinesNo = "";
                string vTicketLineNo = "";
                string vLinesGovNo = "";
                string vTicketType = "";
                string vTicketPCount = "";
                string vTotalAmount = "";
                string vRealIncome = "";
                string vSubsidy_8KM = "";
                string vSubsidy_Limited = "";
                string vSubsidy_Student = "";
                string vSubsidy_Social = "";
                string vSubsidy_MRT = "";
                string vSubsidy_Area = "";
                string vSubsidy_Highway = "";
                string vSubsidy_Others = "";
                string vSubsidy_Dift = "";
                string vDepNo = "";
                string vDailyCostByCar = "0";
                string vDailyIncomeByCar = "0";
                int vDepCarCount = 0;
                double vCarCount = 0.0;
                int vMonthDays = 0;
                double vINC_ZZ = 0.0;
                double vCOS301 = 0.0;
                double vTotalKM = 0.0;
                string[] vaTempStr;
                string vTempStr = "";

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        vSheetCount = wbExcel_H.NumberOfSheets;
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0); //取回第一個工作表
                        for (int i = sheetExcel_H.FirstRowNum + 2; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if ((vRowTemp_H != null) && (vRowTemp_H.Cells.Count >= 15))
                            {
                                switch (vRowTemp_H.Cells[2].ToString().Trim())
                                {
                                    default:
                                        vTicketType = "";
                                        break;
                                    case "一般卡(含員工卡)":
                                    case "一般卡":
                                        vTicketType = "01";
                                        break;
                                    case "學生卡":
                                        vTicketType = "02";
                                        break;
                                    case "愛心卡":
                                        vTicketType = "03";
                                        break;
                                    case "敬老卡":
                                    case "博愛卡":
                                        vTicketType = "04";
                                        break;
                                    case "愛陪卡(有陪)":
                                        vTicketType = "05";
                                        break;
                                    case "愛陪卡(未陪)":
                                    case "愛陪卡(沒陪)":
                                        vTicketType = "06";
                                        break;
                                }
                                vTicketLineNo = (Int32.TryParse(vRowTemp_H.Cells[0].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : ""; //驗票機路線代碼
                                if ((vTicketType != "") && (vTicketLineNo != "")) //如果對應不到驗票機路線代碼或票種就不轉入
                                {
                                    vSQLStrTemp = "select (DepNo + ',' + ERPLinesNo) LinesData from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                    vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LinesData");
                                    vaTempStr = vTempStr.Split(',');
                                    if (Int32.TryParse(vaTempStr[0].ToString().Trim(), out vTempINT))
                                    {
                                        vDepNo = vTempINT.ToString("D2");
                                        vSQLStrTemp = "select CarCount from CarCount where DepNoYM = '" + vCalYM.Trim() + vDepNo.Trim() + "' ";
                                        vDepCarCount = (Int32.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "CarCount"), out vTempINT)) ? vTempINT : 1;
                                        vLinesNo = vaTempStr[1].ToString().Trim();
                                        vSQLStrTemp = "select sum(ActualKM) as TotalKM " + Environment.NewLine +
                                                      "  from RunsheetB " + Environment.NewLine +
                                                      " where LinesNo = '" + vLinesNo + "' " + Environment.NewLine +
                                                      "   and AssignNo in (select AssignNo from RunSheetA where BuDate between '" + vDate_S.ToString("yyyy/MM/dd") + " 00:00:00' and '" + vDate_E.ToString("yyyy/MM/dd") + " 23:59:59') " + Environment.NewLine +
                                                      "   and isnull(ReduceReason, '') = ''";
                                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "TotalKM");
                                        vTotalKM = (vTempStr.Trim() != "") ? Double.Parse(vTempStr) : 0.0;
                                        vMonthDays = (Int32.Parse(eImportYear.Text.Trim()) < 1911) ?
                                                     DateTime.DaysInMonth(Int32.Parse(eImportYear.Text.Trim()) + 1911, Int32.Parse(ddlImportMonth.SelectedValue.Trim())) :
                                                     DateTime.DaysInMonth(Int32.Parse(eImportYear.Text.Trim()), Int32.Parse(ddlImportMonth.SelectedValue.Trim()));
                                        vTicketLineNo = Int32.Parse(vTicketLineNo).ToString("D6");
                                        vIndexNo = vCalYM.Trim() + vDepNo.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        vLinesGovNo = vRowTemp_H.Cells[1].ToString().Trim(); //政府路線編碼

                                        if (Array.IndexOf(vStrArray, vTicketType) != -1)
                                        {
                                            vTicketPCount = "0"; //刷卡量
                                            vRealIncome = "0"; //實際營收金額
                                            vSubsidy_8KM = "0"; //基本里程補貼金額
                                            vSubsidy_Limited = "0"; //票價上限補貼金額
                                            vSubsidy_Student = "0"; //學生卡25折優惠
                                            vSubsidy_Social = "0"; //社福補助
                                            vSubsidy_MRT = "0"; //捷運轉乘優惠補助
                                            vSubsidy_Area = "0"; //A21-A23區間優惠
                                            vSubsidy_Highway = "0"; //國道及台鐵轉乘優惠
                                            vSubsidy_Others = "0"; //其他補貼
                                        }
                                        else
                                        {
                                            vTicketPCount = (Int32.TryParse(vRowTemp_H.Cells[3].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //刷卡量
                                            vRealIncome = (Int32.TryParse(vRowTemp_H.Cells[5].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //實際營收金額
                                            vSubsidy_8KM = (Int32.TryParse(vRowTemp_H.Cells[6].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //基本里程補貼金額
                                            vSubsidy_Limited = (Int32.TryParse(vRowTemp_H.Cells[7].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //票價上限補貼金額
                                            vSubsidy_Student = (Int32.TryParse(vRowTemp_H.Cells[8].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //學生卡25折優惠
                                            vSubsidy_Social = (Int32.TryParse(vRowTemp_H.Cells[9].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //社福補助
                                            vSubsidy_MRT = (Int32.TryParse(vRowTemp_H.Cells[10].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //捷運轉乘優惠補助
                                            vSubsidy_Area = (Int32.TryParse(vRowTemp_H.Cells[11].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //A21-A23區間優惠
                                            vSubsidy_Highway = (Int32.TryParse(vRowTemp_H.Cells[12].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //國道及台鐵轉乘優惠
                                            vSubsidy_Others = (Int32.TryParse(vRowTemp_H.Cells[13].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //其他補貼
                                        }
                                        vSubsidy_Dift = (Int32.TryParse(vRowTemp_H.Cells[14].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //法定差額補貼
                                        vTotalAmount = (double.Parse(vRealIncome) + double.Parse(vSubsidy_Others) + double.Parse(vSubsidy_Dift)).ToString(); //應收總金額

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, TotalKM, CarCount, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                                              "       DailyCostByCar, DailyIncomeByCar, DepCarCount, MonthDays, INC_ZZ, COS301, BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, @TotalKM, 0, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, @RealIncome, @Subsidy_8KM, @Subsidy_Limited, @Subsidy_Student, " + Environment.NewLine +
                                                              "       @Subsidy_Social, @Subsidy_MRT, @Subsidy_Area, @Subsidy_Highway, @Subsidy_Others, @Subsidy_Dift, " + Environment.NewLine +
                                                              "       0, 0, @DepCarCount, @MonthDays, 0, 0, @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalKM", DbType.Double, vTotalKM.ToString()));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Dift", DbType.Double, vSubsidy_Dift));
                                                dsTemp.InsertParameters.Add(new Parameter("DepCarCount", DbType.Int32, vDepCarCount.ToString()));
                                                dsTemp.InsertParameters.Add(new Parameter("MonthDays", DbType.Int32, vMonthDays.ToString()));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    vErrorStr += (Environment.NewLine + eMessage.Message + " IndexNo = [" + vIndexNo + "]");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       RealIncome = RealIncome + @RealIncome, Subsidy_8KM = Subsidy_8KM + @Subsidy_8KM, " + Environment.NewLine +
                                                              "       Subsidy_Limited = Subsidy_Limited + @Subsidy_Limited, Subsidy_Student = Subsidy_Student + @Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social = Subsidy_Social + @Subsidy_Social, Subsidy_MRT = Subsidy_MRT + @Subsidy_MRT, " + Environment.NewLine +
                                                              "       Subsidy_Area = Subsidy_Area + @Subsidy_Area, Subsidy_Highway = Subsidy_Highway + @Subsidy_Highway, " + Environment.NewLine +
                                                              "       Subsidy_Others = Subsidy_Others + @Subsidy_Others, Subsidy_Dift = Subsidy_Dift + @Subsidy_Dift, " + Environment.NewLine +
                                                              "       DepCarCount = @DepCarCount, MonthDays = @MonthDays, TotalKM = @TotalKM, LinesGovNo = @LinesGovNo, " + Environment.NewLine +
                                                              "       ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Dift", DbType.Double, vSubsidy_Dift));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.UpdateParameters.Add(new Parameter("DepCarCount", DbType.Int32, vDepCarCount.ToString()));
                                                dsTemp.UpdateParameters.Add(new Parameter("MonthDays", DbType.Int32, vMonthDays.ToString()));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalKM", DbType.Double, vTotalKM.ToString()));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    vErrorStr += (Environment.NewLine + eMessage.Message + " IndexNo = [" + vIndexNo + "]");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    vErrorStr += (Environment.NewLine + "驗票機路線代碼 [" + vTicketLineNo + "] 對應不到主責站別！");
                                }
                            }
                        }
                        //計算各路線使用車輛數和成本資料
                        using (SqlConnection connCalCost = new SqlConnection())
                        {
                            if (vConnStr == "")
                            {
                                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                            }
                            connCalCost.ConnectionString = vConnStr;
                            vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                            vSQLStrTemp = "select TicketLineNo, DepNo, DepCarCount, sum(TotalKM) TotalKM_T " + Environment.NewLine +
                                          "  from LinesBounds " + Environment.NewLine +
                                          " where CalYM = '" + vCalYM + "' " + Environment.NewLine +
                                          " group by TicketLineNo, DepNo, DepCarCount " + Environment.NewLine +
                                          " order by TicketLineNo, DepNo";
                            SqlCommand cmdCalCost = new SqlCommand(vSQLStrTemp, connCalCost);
                            connCalCost.Open();
                            SqlDataReader drCalCost = cmdCalCost.ExecuteReader();
                            while (drCalCost.Read())
                            {
                                vTicketLineNo = drCalCost["TicketLineNo"].ToString().Trim();
                                vDepNo = drCalCost["DepNo"].ToString().Trim();
                                vTotalKM = (drCalCost["TotalKM_T"].ToString().Trim() != "") ? double.Parse(drCalCost["TotalKM_T"].ToString().Trim()) : 0.0;
                                vDepCarCount = (drCalCost["DepCarCount"].ToString().Trim() != "") ? Int32.Parse(drCalCost["DepCarCount"].ToString().Trim()) : 0;
                                vSQLStrTemp = "select (sum(t.DepKM) - sum(t.DepKM_2)) as RealDepKM " + Environment.NewLine +
                                              "  from( " + Environment.NewLine +
                                              "        select sum(ActualKM) as DepKM, cast(0.0 as float) DepKM_2 " + Environment.NewLine +
                                              "          from RunSheetA " + Environment.NewLine +
                                              "         where DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                              "           and BUDATE between '" + vDate_S.ToString("yyyy/MM/dd") + " 00:00:00' and '" + vDate_E.ToString("yyyy/MM/dd") + " 23:59:59' " + Environment.NewLine +
                                              "         union all " + Environment.NewLine +
                                              "        select cast(0.0 as float) DepKM, sum(ActualKM) DepKM_2 " + Environment.NewLine +
                                              "          from RunSheetB " + Environment.NewLine +
                                              "         where AssignNo in (select AssignNo from RunSheetA where DepNo = '" + vDepNo + "' and BUDATE between '" + vDate_S.ToString("yyyy/MM/dd") + " 00:00:00' and '" + vDate_E.ToString("yyyy/MM/dd") + " 23:59:59') " + Environment.NewLine +
                                              "           and isnull(ReduceReason, '') = '' " + Environment.NewLine +
                                              "           and LinesNo = '99990' " + Environment.NewLine +
                                              "       ) t";
                                vTempFloat = double.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RealDepKM").Trim());
                                vCarCount = (vTotalKM / vTempFloat) * (double)vDepCarCount;
                                vSQLStrTemp = "select (SubjectAMT + OtherAMT + ShareAMT) AMT_COS301 " + Environment.NewLine +
                                              "  from DepMonthIncome " + Environment.NewLine +
                                              " where IncomeYM = '" + vCalYM + "' " + Environment.NewLine +
                                              "   and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                              "   and ChartBarCode = 'COS301' ";
                                vCOS301 = double.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "AMT_COS301"), out vTempFloat) ? vTempFloat : 0.0;
                                vSQLStrTemp = "select (SubjectAMT + OtherAMT + ShareAMT) AMT_INCZZ " + Environment.NewLine +
                                              "  from DepMonthIncome " + Environment.NewLine +
                                              " where IncomeYM = '" + vCalYM + "' " + Environment.NewLine +
                                              "   and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                              "   and ChartBarCode = 'INC ZZ' ";
                                vINC_ZZ = double.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "AMT_INCZZ"), out vTempFloat) ? vTempFloat : 0.0;
                                vTempFloat = (vCarCount * (double)vMonthDays);
                                vDailyCostByCar = (vTempFloat == 0.0) ? "0" : (vCOS301 / vTempFloat).ToString();
                                vDailyIncomeByCar = (vTempFloat == 0.0) ? "0" : (vINC_ZZ / vTempFloat).ToString();

                                using (SqlDataSource dsCalCost = new SqlDataSource())
                                {
                                    dsCalCost.ConnectionString = vConnStr;
                                    vSQLStrTemp = "update LinesBounds set CarCount = @CarCount, COS301 = @COS301, INC_ZZ = @INC_ZZ, " + Environment.NewLine +
                                                  "                       DailyCostByCar = @DailyCostByCar, DailyIncomeByCar = @DailyIncomeByCar " + Environment.NewLine +
                                                  " where TicketLineNo = @TicketLineNo and DepNo = @DepNo";
                                    dsCalCost.UpdateCommand = vSQLStrTemp;
                                    dsCalCost.UpdateParameters.Clear();
                                    dsCalCost.UpdateParameters.Add(new Parameter("CarCount", DbType.Double, vCarCount.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("COS301", DbType.Double, vCOS301.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("INC_ZZ", DbType.Double, vINC_ZZ.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("DailyCostByCar", DbType.Double, vDailyCostByCar.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("DailyIncomeByCar", DbType.Double, vDailyIncomeByCar.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                    dsCalCost.UpdateParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                    try
                                    {
                                        dsCalCost.Update();
                                    }
                                    catch (Exception eMessage)
                                    {
                                        Response.Write("<Script language='Javascript'>");
                                        Response.Write("alert(" + eMessage.Message + ")");
                                        Response.Write("</" + "Script>");
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2010 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        vSheetCount = wbExcel_X.NumberOfSheets;
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0); //取回第一個工作表
                        for (int i = sheetExcel_X.FirstRowNum + 2; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if ((vRowTemp_X != null) && (vRowTemp_X.Cells.Count >= 15))
                            {
                                switch (vRowTemp_X.Cells[2].ToString().Trim())
                                {
                                    default:
                                        vTicketType = "";
                                        break;
                                    case "一般卡(含員工卡)":
                                    case "一般卡":
                                        vTicketType = "01";
                                        break;
                                    case "學生卡":
                                        vTicketType = "02";
                                        break;
                                    case "愛心卡":
                                        vTicketType = "03";
                                        break;
                                    case "敬老卡":
                                    case "博愛卡":
                                        vTicketType = "04";
                                        break;
                                    case "愛陪卡(有陪)":
                                        vTicketType = "05";
                                        break;
                                    case "愛陪卡(未陪)":
                                    case "愛陪卡(沒陪)":
                                        vTicketType = "06";
                                        break;
                                }
                                vTicketLineNo = (Int32.TryParse(vRowTemp_X.Cells[0].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : ""; //驗票機路線代碼
                                if ((vTicketType != "") && (vTicketLineNo != "")) //如果對應不到驗票機路線代碼或票種就不轉入
                                {
                                    vSQLStrTemp = "select (DepNo + ',' + ERPLinesNo) LinesData from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                    vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LinesData");
                                    vaTempStr = vTempStr.Split(',');
                                    if (Int32.TryParse(vaTempStr[0].ToString().Trim(), out vTempINT))
                                    {
                                        vDepNo = vTempINT.ToString("D2");
                                        vSQLStrTemp = "select CarCount from CarCount where DepNoYM = '" + vCalYM.Trim() + vDepNo.Trim() + "' ";
                                        vDepCarCount = (Int32.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "CarCount"), out vTempINT)) ? vTempINT : 1;
                                        vLinesNo = vaTempStr[1].ToString().Trim();
                                        vSQLStrTemp = "select sum(ActualKM) as TotalKM " + Environment.NewLine +
                                                      "  from RunsheetB " + Environment.NewLine +
                                                      " where LinesNo = '" + vLinesNo + "' " + Environment.NewLine +
                                                      "   and AssignNo in (select AssignNo from RunSheetA where BuDate between '" + vDate_S.ToString("yyyy/MM/dd") + " 00:00:00' and '" + vDate_E.ToString("yyyy/MM/dd") + " 23:59:59') " + Environment.NewLine +
                                                      "   and isnull(ReduceReason, '') = ''";
                                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "TotalKM");
                                        vTotalKM = (vTempStr.Trim() != "") ? Double.Parse(vTempStr) : 0.0;
                                        vMonthDays = (Int32.Parse(eImportYear.Text.Trim()) < 1911) ?
                                                     DateTime.DaysInMonth(Int32.Parse(eImportYear.Text.Trim()) + 1911, Int32.Parse(ddlImportMonth.SelectedValue.Trim())) :
                                                     DateTime.DaysInMonth(Int32.Parse(eImportYear.Text.Trim()), Int32.Parse(ddlImportMonth.SelectedValue.Trim()));
                                        vTicketLineNo = Int32.Parse(vTicketLineNo).ToString("D6");
                                        vIndexNo = vCalYM.Trim() + vDepNo.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        vLinesGovNo = vRowTemp_X.Cells[1].ToString().Trim(); //政府路線編碼

                                        if (Array.IndexOf(vStrArray, vTicketType) != -1)
                                        {
                                            vTicketPCount = "0"; //刷卡量
                                            vRealIncome = "0"; //實際營收金額
                                            vSubsidy_8KM = "0"; //基本里程補貼金額
                                            vSubsidy_Limited = "0"; //票價上限補貼金額
                                            vSubsidy_Student = "0"; //學生卡25折優惠
                                            vSubsidy_Social = "0"; //社福補助
                                            vSubsidy_MRT = "0"; //捷運轉乘優惠補助
                                            vSubsidy_Area = "0"; //A21-A23區間優惠
                                            vSubsidy_Highway = "0"; //國道及台鐵轉乘優惠
                                            vSubsidy_Others = "0"; //其他補貼
                                        }
                                        else
                                        {
                                            vTicketPCount = (Int32.TryParse(vRowTemp_X.Cells[3].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //刷卡量                                            
                                            vRealIncome = (Int32.TryParse(vRowTemp_X.Cells[5].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //實際營收金額
                                            vSubsidy_8KM = (Int32.TryParse(vRowTemp_X.Cells[6].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //基本里程補貼金額
                                            vSubsidy_Limited = (Int32.TryParse(vRowTemp_X.Cells[7].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //票價上限補貼金額
                                            vSubsidy_Student = (Int32.TryParse(vRowTemp_X.Cells[8].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //學生卡25折優惠
                                            vSubsidy_Social = (Int32.TryParse(vRowTemp_X.Cells[9].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //社福補助
                                            vSubsidy_MRT = (Int32.TryParse(vRowTemp_X.Cells[10].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //捷運轉乘優惠補助
                                            vSubsidy_Area = (Int32.TryParse(vRowTemp_X.Cells[11].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //A21-A23區間優惠
                                            vSubsidy_Highway = (Int32.TryParse(vRowTemp_X.Cells[12].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //國道及台鐵轉乘優惠
                                            vSubsidy_Others = (Int32.TryParse(vRowTemp_X.Cells[13].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //其他補貼
                                        }
                                        vSubsidy_Dift = (Int32.TryParse(vRowTemp_X.Cells[14].ToString().Replace(",", "").Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //法定差額補貼
                                        vTotalAmount = (double.Parse(vRealIncome) + double.Parse(vSubsidy_Others) + double.Parse(vSubsidy_Dift)).ToString(); //應收總金額

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, TotalKM, CarCount, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                                              "       DailyCostByCar, DailyIncomeByCar, DepCarCount, MonthDays, INC_ZZ, COS301, BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, @TotalKM, 0, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, @RealIncome, @Subsidy_8KM, @Subsidy_Limited, @Subsidy_Student, " + Environment.NewLine +
                                                              "       @Subsidy_Social, @Subsidy_MRT, @Subsidy_Area, @Subsidy_Highway, @Subsidy_Others, @Subsidy_Dift, " + Environment.NewLine +
                                                              "       0, 0, @DepCarCount, @MonthDays, 0, 0, @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalKM", DbType.Double, vTotalKM.ToString()));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Dift", DbType.Double, vSubsidy_Dift));
                                                dsTemp.InsertParameters.Add(new Parameter("DepCarCount", DbType.Int32, vDepCarCount.ToString()));
                                                dsTemp.InsertParameters.Add(new Parameter("MonthDays", DbType.Int32, vMonthDays.ToString()));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    vErrorStr += (Environment.NewLine + eMessage.Message + " IndexNo = [" + vIndexNo + "]");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       RealIncome = RealIncome + @RealIncome, Subsidy_8KM = Subsidy_8KM + @Subsidy_8KM, " + Environment.NewLine +
                                                              "       Subsidy_Limited = Subsidy_Limited + @Subsidy_Limited, Subsidy_Student = Subsidy_Student + @Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social = Subsidy_Social + @Subsidy_Social, Subsidy_MRT = Subsidy_MRT + @Subsidy_MRT, " + Environment.NewLine +
                                                              "       Subsidy_Area = Subsidy_Area + @Subsidy_Area, Subsidy_Highway = Subsidy_Highway + @Subsidy_Highway, " + Environment.NewLine +
                                                              "       Subsidy_Others = Subsidy_Others + @Subsidy_Others, Subsidy_Dift = Subsidy_Dift + @Subsidy_Dift, " + Environment.NewLine +
                                                              "       DepCarCount = @DepCarCount, MonthDays = @MonthDays, TotalKM = @TotalKM, " + Environment.NewLine +
                                                              "       ModifyMan = @ModifyMan, ModifyDate = GetDate(), LinesGovNo = @LinesGovNo " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Dift", DbType.Double, vSubsidy_Dift));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.UpdateParameters.Add(new Parameter("DepCarCount", DbType.Int32, vDepCarCount.ToString()));
                                                dsTemp.UpdateParameters.Add(new Parameter("MonthDays", DbType.Int32, vMonthDays.ToString()));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalKM", DbType.Double, vTotalKM.ToString()));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    vErrorStr += (Environment.NewLine + eMessage.Message + " IndexNo = [" + vIndexNo + "]");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    vErrorStr += (Environment.NewLine + "驗票機路線代碼 [" + vTicketLineNo + "] 對應不到主責站別！");
                                }
                            }
                        }
                        //計算各路線使用車輛數和成本資料
                        using (SqlConnection connCalCost = new SqlConnection())
                        {
                            if (vConnStr == "")
                            {
                                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                            }
                            connCalCost.ConnectionString = vConnStr;
                            vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                            vSQLStrTemp = "select TicketLineNo, DepNo, DepCarCount, sum(TotalKM) TotalKM_T " + Environment.NewLine +
                                          "  from LinesBounds " + Environment.NewLine +
                                          " where CalYM = '" + vCalYM + "' " + Environment.NewLine +
                                          " group by TicketLineNo, DepNo, DepCarCount " + Environment.NewLine +
                                          " order by TicketLineNo, DepNo";
                            SqlCommand cmdCalCost = new SqlCommand(vSQLStrTemp, connCalCost);
                            connCalCost.Open();
                            SqlDataReader drCalCost = cmdCalCost.ExecuteReader();
                            while (drCalCost.Read())
                            {
                                vTicketLineNo = drCalCost["TicketLineNo"].ToString().Trim();
                                vDepNo = drCalCost["DepNo"].ToString().Trim();
                                vTotalKM = (drCalCost["TotalKM_T"].ToString().Trim() != "") ? double.Parse(drCalCost["TotalKM_T"].ToString().Trim()) : 0.0;
                                vDepCarCount = (drCalCost["DepCarCount"].ToString().Trim() != "") ? Int32.Parse(drCalCost["DepCarCount"].ToString().Trim()) : 0;
                                vSQLStrTemp = "select (sum(t.DepKM) - sum(t.DepKM_2)) as RealDepKM " + Environment.NewLine +
                                              "  from( " + Environment.NewLine +
                                              "        select sum(ActualKM) as DepKM, cast(0.0 as float) DepKM_2 " + Environment.NewLine +
                                              "          from RunSheetA " + Environment.NewLine +
                                              "         where DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                              "           and BUDATE between '" + vDate_S.ToString("yyyy/MM/dd") + " 00:00:00' and '" + vDate_E.ToString("yyyy/MM/dd") + " 23:59:59' " + Environment.NewLine +
                                              "         union all " + Environment.NewLine +
                                              "        select cast(0.0 as float) DepKM, sum(ActualKM) DepKM_2 " + Environment.NewLine +
                                              "          from RunSheetB " + Environment.NewLine +
                                              "         where AssignNo in (select AssignNo from RunSheetA where DepNo = '" + vDepNo + "' and BUDATE between '" + vDate_S.ToString("yyyy/MM/dd") + " 00:00:00' and '" + vDate_E.ToString("yyyy/MM/dd") + " 23:59:59') " + Environment.NewLine +
                                              "           and isnull(ReduceReason, '') = '' " + Environment.NewLine +
                                              "           and LinesNo = '99990' " + Environment.NewLine +
                                              "       ) t";
                                vTempFloat = double.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RealDepKM").Trim());
                                vCarCount = (vTotalKM / vTempFloat) * (double)vDepCarCount;
                                vSQLStrTemp = "select (SubjectAMT + OtherAMT + ShareAMT) AMT_COS301 " + Environment.NewLine +
                                              "  from DepMonthIncome " + Environment.NewLine +
                                              " where IncomeYM = '" + vCalYM + "' " + Environment.NewLine +
                                              "   and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                              "   and ChartBarCode = 'COS301' ";
                                vCOS301 = double.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "AMT_COS301"), out vTempFloat) ? vTempFloat : 0.0;
                                vSQLStrTemp = "select (SubjectAMT + OtherAMT + ShareAMT) AMT_INCZZ " + Environment.NewLine +
                                              "  from DepMonthIncome " + Environment.NewLine +
                                              " where IncomeYM = '" + vCalYM + "' " + Environment.NewLine +
                                              "   and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                              "   and ChartBarCode = 'INC ZZ' ";
                                vINC_ZZ = double.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "AMT_INCZZ"), out vTempFloat) ? vTempFloat : 0.0;
                                vTempFloat = (vCarCount * (double)vMonthDays);
                                vDailyCostByCar = (vTempFloat == 0.0) ? "0" : (vCOS301 / vTempFloat).ToString();
                                vDailyIncomeByCar = (vTempFloat == 0.0) ? "0" : (vINC_ZZ / vTempFloat).ToString();

                                using (SqlDataSource dsCalCost = new SqlDataSource())
                                {
                                    dsCalCost.ConnectionString = vConnStr;
                                    vSQLStrTemp = "update LinesBounds set CarCount = @CarCount, COS301 = @COS301, INC_ZZ = @INC_ZZ, " + Environment.NewLine +
                                                  "                       DailyCostByCar = @DailyCostByCar, DailyIncomeByCar = @DailyIncomeByCar " + Environment.NewLine +
                                                  " where TicketLineNo = @TicketLineNo and DepNo = @DepNo";
                                    dsCalCost.UpdateCommand = vSQLStrTemp;
                                    dsCalCost.UpdateParameters.Clear();
                                    dsCalCost.UpdateParameters.Add(new Parameter("CarCount", DbType.Double, vCarCount.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("COS301", DbType.Double, vCOS301.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("INC_ZZ", DbType.Double, vINC_ZZ.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("DailyCostByCar", DbType.Double, vDailyCostByCar.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("DailyIncomeByCar", DbType.Double, vDailyIncomeByCar.ToString()));
                                    dsCalCost.UpdateParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                    dsCalCost.UpdateParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                    try
                                    {
                                        dsCalCost.Update();
                                    }
                                    catch (Exception eMessage)
                                    {
                                        Response.Write("<Script language='Javascript'>");
                                        Response.Write("alert(" + eMessage.Message + ")");
                                        Response.Write("</" + "Script>");
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
        }

        /// <summary>
        /// 從表三取回資料
        /// </summary>
        private void ImportFromExcel_List3()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";
                int vTempINT = 0;
                Double vTempFloat = 0.0;
                string vTempStr = "";
                string[] vaTempStr;
                DataTable dtTemp = CreateDataTable();
                DataTable dtTempForLinq = CreateDataForLinq_List3();
                DataRow drTemp;
                DateTime vDate_S = (Int32.Parse(eImportYear.Text.Trim()) < 1911) ?
                                   new DateTime(Int32.Parse(eImportYear.Text.Trim()) + 1911, Int32.Parse(ddlImportMonth.SelectedValue.Trim()), 1) :
                                   new DateTime(Int32.Parse(eImportYear.Text.Trim()), Int32.Parse(ddlImportMonth.SelectedValue.Trim()), 1);
                DateTime vDate_E = vDate_S.AddMonths(1).AddDays(-1);


                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vTicketLineNo = "";
                string vLinesNo = "";
                string vLinesGovNo = "";
                string vTicketType = "";
                string vTicketPCount = "";
                string vTotalAmount = "";
                string vRealIncome = "";
                string vSubsidy_8KM = "";
                string vSubsidy_Limited = "";
                string vSubsidy_Student = "";
                string vSubsidy_Social = "";
                string vSubsidy_MRT = "";
                string vSubsidy_Area = "";
                string vSubsidy_Highway = "";
                string vSubsidy_Others = "";
                string vSubsidy_Dift = "";
                string vCar_ID = "";
                string vDepNo = "";
                string vDailyCostByCar = "0";
                string vDailyIncomeByCar = "0";
                double vCarCount = 0.0;
                int vDepCarCount = 0;
                int vMonthDays = 0;
                double vINC_ZZ = 0.0;
                double vCOS301 = 0.0;
                double vTotalKM = 0.0;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003)
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        for (int i = 0; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if ((vRowExcel_H != null) && (vRowExcel_H.Cells.Count > 26) && (vRowExcel_H.Cells[3].ToString().Replace("車號", "") != ""))
                            {
                                drTemp = dtTempForLinq.NewRow();
                                drTemp["TicketLineNo"] = (Int32.TryParse(vRowExcel_H.Cells[4].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : ""; //驗票機路線代碼
                                vCar_ID = vRowExcel_H.Cells[3].ToString().Trim(); //牌照號碼
                                if ((vCar_ID.Substring(0, 1) != "4") && (vCar_ID.ToUpper().IndexOf("-U7") > 0))
                                {
                                    vCar_ID = vCar_ID.ToUpper().Replace("-U7", "-U-7");
                                }
                                vSQLStrTemp = "select CompanyNo from Car_InfoA where Car_ID = '" + vCar_ID + "' ";
                                vDepNo = PF.GetValue(vConnStr, vSQLStrTemp, "CompanyNo"); //車輛站別
                                drTemp["DepNo"] = (Int32.TryParse(vDepNo, out vTempINT)) ? vTempINT.ToString("D2") : "XX"; //車輛站別

                                switch (vRowExcel_H.Cells[9].ToString().Trim().Replace("桃園市民卡-", ""))
                                {
                                    default:
                                        vTicketType = "99";
                                        break;
                                    case "一般":
                                    case "一般卡(含員工卡)":
                                    case "一般卡":
                                        vTicketType = "01";
                                        break;
                                    case "國小":
                                    case "市立國中":
                                    case "市立高中":
                                    case "大專":
                                    case "學生卡":
                                        vTicketType = "02";
                                        break;
                                    case "愛心":
                                    case "愛心卡":
                                        vTicketType = "03";
                                        break;
                                    case "敬老":
                                    case "敬老卡":
                                        vTicketType = "04";
                                        break;
                                    case "陪伴(有陪)":
                                    case "愛陪卡(有陪)":
                                        vTicketType = "05";
                                        break;
                                    case "陪伴(未陪)":
                                    case "陪伴(沒陪)":
                                    case "愛陪卡(未陪)":
                                    case "愛陪卡(沒陪)":
                                        vTicketType = "06";
                                        break;
                                }
                                drTemp["IndexNo"] = vCalYM.Trim() + drTemp["DepNo"].ToString().Trim() + vTicketType + Int32.Parse(drTemp["TicketLineNo"].ToString().Trim()).ToString("D6");
                                drTemp["TicketType"] = vTicketType;
                                if ((drTemp["TicketLineNo"].ToString().Trim() != "") && (vTicketType != ""))
                                {
                                    if (Array.IndexOf(vStrArray, vTicketType) != -1)
                                    {
                                        drTemp["TicketPCount"] = 0; //刷卡量
                                        drTemp["TotalAmount"] = 0.0; //應收總金額
                                        drTemp["RealIncome"] = 0; //實際營收金額
                                        drTemp["Subsidy_8KM"] = 0; //基本里程補貼金額
                                        drTemp["Subsidy_Limited"] = 0; //票價上限補貼金額
                                        drTemp["Subsidy_Student"] = 0; //學生卡25折優惠
                                        drTemp["Subsidy_Social"] = 0; //社福補助
                                        drTemp["Subsidy_MRT"] = 0; //捷運轉乘優惠補助
                                        drTemp["Subsidy_Area"] = 0; //A21-A23區間優惠
                                        drTemp["Subsidy_Highway"] = 0; //國道及台鐵轉乘優惠
                                        drTemp["Subsidy_Others"] = 0; //其他補貼
                                        drTemp["Subsidy_Dift"] = Double.TryParse(vRowExcel_H.Cells[26].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"; //法定差額補貼
                                        dtTempForLinq.Rows.Add(drTemp);
                                    }
                                    else
                                    {
                                        drTemp["TicketPCount"] = "1"; //刷卡量
                                        drTemp["RealIncome"] = Int32.TryParse(vRowExcel_H.Cells[17].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //實際營收金額
                                        drTemp["Subsidy_8KM"] = Int32.TryParse(vRowExcel_H.Cells[20].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //基本里程補貼金額
                                        drTemp["Subsidy_Limited"] = Int32.TryParse(vRowExcel_H.Cells[21].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //票價上限補貼金額
                                        drTemp["Subsidy_Student"] = Int32.TryParse(vRowExcel_H.Cells[18].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; ; //學生卡25折優惠
                                        drTemp["Subsidy_Social"] = Int32.TryParse(vRowExcel_H.Cells[19].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; ; //社福補助
                                        drTemp["Subsidy_MRT"] = Int32.TryParse(vRowExcel_H.Cells[22].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //捷運轉乘優惠補助
                                        drTemp["Subsidy_Area"] = Int32.TryParse(vRowExcel_H.Cells[23].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //A21-A23區間優惠
                                        drTemp["Subsidy_Highway"] = Int32.TryParse(vRowExcel_H.Cells[24].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //國道及台鐵轉乘優惠
                                        drTemp["Subsidy_Others"] = Int32.TryParse(vRowExcel_H.Cells[25].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; ; //其他補貼*/
                                        drTemp["Subsidy_Dift"] = Double.TryParse(vRowExcel_H.Cells[26].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"; //法定差額補貼
                                        drTemp["TotalAmount"] = "0.0"; //應收總金額
                                        dtTempForLinq.Rows.Add(drTemp);
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2007 以後)
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        for (int i = 0; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if ((vRowExcel_X != null) && (vRowExcel_X.Cells.Count > 26) && (vRowExcel_X.Cells[3].ToString().Replace("車號", "") != ""))
                            {
                                drTemp = dtTempForLinq.NewRow();
                                drTemp["TicketLineNo"] = (Int32.TryParse(vRowExcel_X.Cells[4].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : ""; //驗票機路線代碼
                                vCar_ID = vRowExcel_X.Cells[3].ToString().Trim(); //牌照號碼
                                if ((vCar_ID.Substring(0, 1) != "4") && (vCar_ID.ToUpper().IndexOf("-U7") > 0))
                                {
                                    vCar_ID = vCar_ID.ToUpper().Replace("-U7", "-U-7");
                                }
                                vSQLStrTemp = "select CompanyNo from Car_InfoA where Car_ID = '" + vCar_ID + "' ";
                                vDepNo = PF.GetValue(vConnStr, vSQLStrTemp, "CompanyNo"); //車輛站別
                                drTemp["DepNo"] = (Int32.TryParse(vDepNo, out vTempINT)) ? vTempINT.ToString("D2") : "XX"; //車輛站別

                                switch (vRowExcel_X.Cells[9].ToString().Trim().Replace("桃園市民卡-", ""))
                                {
                                    default:
                                        vTicketType = "99";
                                        break;
                                    case "一般":
                                    case "一般卡(含員工卡)":
                                    case "一般卡":
                                        vTicketType = "01";
                                        break;
                                    case "國小":
                                    case "市立國中":
                                    case "市立高中":
                                    case "大專":
                                    case "學生卡":
                                        vTicketType = "02";
                                        break;
                                    case "愛心":
                                    case "愛心卡":
                                        vTicketType = "03";
                                        break;
                                    case "敬老":
                                    case "敬老卡":
                                        vTicketType = "04";
                                        break;
                                    case "陪伴(有陪)":
                                    case "愛陪卡(有陪)":
                                        vTicketType = "05";
                                        break;
                                    case "陪伴(未陪)":
                                    case "陪伴(沒陪)":
                                    case "愛陪卡(未陪)":
                                    case "愛陪卡(沒陪)":
                                        vTicketType = "06";
                                        break;
                                }
                                drTemp["IndexNo"] = vCalYM.Trim() + drTemp["DepNo"].ToString().Trim() + vTicketType + Int32.Parse(drTemp["TicketLineNo"].ToString().Trim()).ToString("D6");
                                drTemp["TicketType"] = vTicketType;
                                if ((drTemp["TicketLineNo"].ToString().Trim() != "") && (vTicketType != ""))
                                {
                                    if (Array.IndexOf(vStrArray, vTicketType) != -1)
                                    {
                                        drTemp["TicketPCount"] = 0; //刷卡量
                                        drTemp["TotalAmount"] = 0.0; //應收總金額
                                        drTemp["RealIncome"] = 0; //實際營收金額
                                        drTemp["Subsidy_8KM"] = 0; //基本里程補貼金額
                                        drTemp["Subsidy_Limited"] = 0; //票價上限補貼金額
                                        drTemp["Subsidy_Student"] = 0; //學生卡25折優惠
                                        drTemp["Subsidy_Social"] = 0; //社福補助
                                        drTemp["Subsidy_MRT"] = 0; //捷運轉乘優惠補助
                                        drTemp["Subsidy_Area"] = 0; //A21-A23區間優惠
                                        drTemp["Subsidy_Highway"] = 0; //國道及台鐵轉乘優惠
                                        drTemp["Subsidy_Others"] = 0; //其他補貼
                                        drTemp["Subsidy_Dift"] = Double.TryParse(vRowExcel_X.Cells[26].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"; //法定差額補貼
                                        dtTempForLinq.Rows.Add(drTemp);
                                    }
                                    else
                                    {
                                        drTemp["TicketPCount"] = "1"; //刷卡量
                                        drTemp["TotalAmount"] = "0.0"; //應收總金額
                                        drTemp["RealIncome"] = Int32.TryParse(vRowExcel_X.Cells[17].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //實際營收金額
                                        drTemp["Subsidy_8KM"] = Int32.TryParse(vRowExcel_X.Cells[20].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //基本里程補貼金額
                                        drTemp["Subsidy_Limited"] = Int32.TryParse(vRowExcel_X.Cells[21].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //票價上限補貼金額
                                        drTemp["Subsidy_Student"] = Int32.TryParse(vRowExcel_X.Cells[18].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; ; //學生卡25折優惠
                                        drTemp["Subsidy_Social"] = Int32.TryParse(vRowExcel_X.Cells[19].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; ; //社福補助
                                        drTemp["Subsidy_MRT"] = Int32.TryParse(vRowExcel_X.Cells[22].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //捷運轉乘優惠補助
                                        drTemp["Subsidy_Area"] = Int32.TryParse(vRowExcel_X.Cells[23].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //A21-A23區間優惠
                                        drTemp["Subsidy_Highway"] = Int32.TryParse(vRowExcel_X.Cells[24].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; //國道及台鐵轉乘優惠
                                        drTemp["Subsidy_Others"] = Int32.TryParse(vRowExcel_X.Cells[25].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0"; ; //其他補貼*/
                                        drTemp["Subsidy_Dift"] = Double.TryParse(vRowExcel_X.Cells[26].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"; //法定差額補貼
                                        dtTempForLinq.Rows.Add(drTemp);
                                    }
                                }
                            }
                        }
                        break;

                    case ".csv": //用逗號隔開的檔案
                        var vSourceFile = new StreamReader(fuExcel.FileContent, Encoding.Default);
                        while (!vSourceFile.EndOfStream)
                        {
                            string vSourceData = vSourceFile.ReadLine();
                            if (vSourceData.Trim() != "")
                            {
                                string[] vaSourceData = vSourceData.Split(',');
                                if ((vaSourceData.Length > 26) && (vaSourceData[3].Replace("車號", "") != ""))
                                {
                                    drTemp = dtTempForLinq.NewRow();
                                    vCar_ID = vaSourceData[3].ToString().Trim().Replace("_B", ""); //牌照號碼
                                    if ((vCar_ID.Substring(0, 1) != "4") && (vCar_ID.ToUpper().IndexOf("-U7") > 0))
                                    {
                                        vCar_ID = vCar_ID.ToUpper().Replace("-U7", "-U-7");
                                    }
                                    vSQLStrTemp = "select CompanyNo from Car_InfoA where Car_ID = '" + vCar_ID + "' ";
                                    vDepNo = PF.GetValue(vConnStr, vSQLStrTemp, "CompanyNo");
                                    drTemp["DepNo"] = (Int32.TryParse(vDepNo, out vTempINT)) ? vTempINT.ToString("D2") : "XX"; //車輛站別
                                    drTemp["TicketLineNo"] = (Int32.TryParse(vaSourceData[4].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : ""; //驗票機路線代碼
                                    vTempStr = vaSourceData[9].ToString().Trim().Replace("桃園市民卡-", "");
                                    switch (vTempStr)
                                    {
                                        default:
                                            vTicketType = "99";
                                            break;
                                        case "一般":
                                        case "一般卡(含員工卡)":
                                        case "一般卡":
                                            vTicketType = "01";
                                            break;
                                        case "國小":
                                        case "市立國中":
                                        case "市立高中":
                                        case "大專":
                                        case "學生卡":
                                            vTicketType = "02";
                                            break;
                                        case "愛心":
                                        case "愛心卡":
                                            vTicketType = "03";
                                            break;
                                        case "敬老":
                                        case "敬老卡":
                                            vTicketType = "04";
                                            break;
                                        case "陪伴(有陪)":
                                        case "愛陪卡(有陪)":
                                            vTicketType = "05";
                                            break;
                                        case "陪伴(未陪)":
                                        case "陪伴(沒陪)":
                                        case "愛陪卡(未陪)":
                                        case "愛陪卡(沒陪)":
                                            vTicketType = "06";
                                            break;
                                    }
                                    drTemp["IndexNo"] = vCalYM.Trim() + drTemp["DepNo"].ToString().Trim() + vTicketType + Int32.Parse(drTemp["TicketLineNo"].ToString().Trim()).ToString("D6");
                                    drTemp["TicketType"] = vTicketType;
                                    if ((drTemp["TicketLineNo"].ToString().Trim() != "") && (vTicketType != ""))
                                    {
                                        if (Array.IndexOf(vStrArray, vTicketType) != -1)
                                        {
                                            drTemp["TicketPCount"] = 0; //刷卡量
                                            drTemp["TotalAmount"] = 0.0; //應收總金額
                                            drTemp["RealIncome"] = 0; //實際營收金額
                                            drTemp["Subsidy_8KM"] = 0; //基本里程補貼金額
                                            drTemp["Subsidy_Limited"] = 0; //票價上限補貼金額
                                            drTemp["Subsidy_Student"] = 0; //學生卡25折優惠
                                            drTemp["Subsidy_Social"] = 0; //社福補助
                                            drTemp["Subsidy_MRT"] = 0; //捷運轉乘優惠補助
                                            drTemp["Subsidy_Area"] = 0; //A21-A23區間優惠
                                            drTemp["Subsidy_Highway"] = 0; //國道及台鐵轉乘優惠
                                            drTemp["Subsidy_Others"] = 0; //其他補貼
                                            drTemp["Subsidy_Dift"] = double.TryParse(vaSourceData[26].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"; //法定差額補貼
                                            dtTempForLinq.Rows.Add(drTemp);
                                        }
                                        else
                                        {
                                            drTemp["TicketPCount"] = 1; //刷卡量
                                            drTemp["TotalAmount"] = "0.0"; //應收總金額
                                            drTemp["RealIncome"] = Int32.TryParse(vaSourceData[17].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0";
                                            drTemp["Subsidy_8KM"] = Int32.TryParse(vaSourceData[20].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0";
                                            drTemp["Subsidy_Limited"] = Int32.TryParse(vaSourceData[21].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0";
                                            drTemp["Subsidy_Student"] = Int32.TryParse(vaSourceData[18].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0";
                                            drTemp["Subsidy_Social"] = Int32.TryParse(vaSourceData[19].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0";
                                            drTemp["Subsidy_MRT"] = Int32.TryParse(vaSourceData[22].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0";
                                            drTemp["Subsidy_Area"] = Int32.TryParse(vaSourceData[23].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0";
                                            drTemp["Subsidy_Highway"] = Int32.TryParse(vaSourceData[24].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0";
                                            drTemp["Subsidy_Others"] = Int32.TryParse(vaSourceData[25].ToString().Trim(), out vTempINT) ? vTempINT.ToString() : "0";
                                            drTemp["Subsidy_Dift"] = double.TryParse(vaSourceData[26].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"; //法定差額補貼
                                            dtTempForLinq.Rows.Add(drTemp);
                                        }
                                    }
                                }
                            }
                        }
                        break;
                }
                if (dtTempForLinq.Rows.Count > 0)
                {
                    //2022.05.12 第一次試著實作 LINQ 查詢
                    var TestTable = from TT in dtTempForLinq.AsEnumerable()
                                    group TT by TT.Field<string>("IndexNo")
                                    into groupTemp
                                    select new
                                    {
                                        _IndexNo = groupTemp.Key,
                                        _TicketPCount = groupTemp.Sum(TT => TT.Field<Int32>("TicketPCount")),
                                        _TotalAmount = groupTemp.Sum(TT => TT.Field<Double>("TotalAmount")),
                                        _RealIncome = groupTemp.Sum(TT => TT.Field<Int32>("RealIncome")),
                                        _Subsidy_8KM = groupTemp.Sum(TT => TT.Field<Int32>("Subsidy_8KM")),
                                        _Subsidy_Limited = groupTemp.Sum(TT => TT.Field<Int32>("Subsidy_Limited")),
                                        _Subsidy_Student = groupTemp.Sum(TT => TT.Field<Int32>("Subsidy_Student")),
                                        _Subsidy_Social = groupTemp.Sum(TT => TT.Field<Int32>("Subsidy_Social")),
                                        _Subsidy_MRT = groupTemp.Sum(TT => TT.Field<Int32>("Subsidy_MRT")),
                                        _Subsidy_Area = groupTemp.Sum(TT => TT.Field<Int32>("Subsidy_Area")),
                                        _Subsidy_Highway = groupTemp.Sum(TT => TT.Field<Int32>("Subsidy_Highway")),
                                        _Subsidy_Others = groupTemp.Sum(TT => TT.Field<Int32>("Subsidy_Others")),
                                        _Subsidy_Dift = groupTemp.Sum(TT => TT.Field<Double>("Subsidy_Dift"))
                                    };

                    foreach (var TT2 in TestTable)
                    {
                        vIndexNo = TT2._IndexNo.Trim();
                        vCalYM = vIndexNo.Substring(0, 5);
                        vDepNo = vIndexNo.Substring(5, 2);
                        vTicketType = vIndexNo.Substring(7, 2);
                        vTicketLineNo = vIndexNo.Substring(9, 6); //驗票機路線代碼
                        vSQLStrTemp = "select (ERPLinesNo + ',' + GOVLinesNo) LinesData from LinesNoChart where TicketLinesNo = '" + Int32.Parse(vTicketLineNo).ToString() + "' ";
                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LinesData");
                        vaTempStr = vTempStr.Split(',');
                        vLinesNo = (vaTempStr.Length >= 2) ? vaTempStr[0].ToString().Trim() : "";
                        vLinesGovNo = (vaTempStr.Length >= 2) ? vaTempStr[1].ToString().Trim() : ""; //路線政府編碼
                        vSQLStrTemp = "select CarCount from CarCount where DepNoYM = '" + vCalYM.Trim() + vDepNo.Trim() + "' ";
                        vDepCarCount = (Int32.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "CarCount"), out vTempINT)) ? vTempINT : 1;
                        vSQLStrTemp = "select sum(ActualKM) as TotalKM " + Environment.NewLine +
                                      "  from RunsheetB " + Environment.NewLine +
                                      " where LinesNo = '" + vLinesNo + "' " + Environment.NewLine +
                                      "   and AssignNo in (select AssignNo from RunSheetA where BuDate between '" + vDate_S.ToString("yyyy/MM/dd") + " 00:00:00' and '" + vDate_E.ToString("yyyy/MM/dd") + " 23:59:59') " + Environment.NewLine +
                                      "   and isnull(ReduceReason, '') = ''";
                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "TotalKM");
                        vTotalKM = (vTempStr.Trim() != "") ? Double.Parse(vTempStr) : 0.0;
                        vMonthDays = (Int32.Parse(eImportYear.Text.Trim()) < 1911) ?
                                     DateTime.DaysInMonth(Int32.Parse(eImportYear.Text.Trim()) + 1911, Int32.Parse(ddlImportMonth.SelectedValue.Trim())) :
                                     DateTime.DaysInMonth(Int32.Parse(eImportYear.Text.Trim()), Int32.Parse(ddlImportMonth.SelectedValue.Trim()));

                        vTicketPCount = TT2._TicketPCount.ToString();
                        vTotalAmount = ((double)TT2._RealIncome + (double)TT2._Subsidy_Others + TT2._Subsidy_Dift).ToString();
                        vRealIncome = TT2._RealIncome.ToString();
                        vSubsidy_8KM = TT2._Subsidy_8KM.ToString();
                        vSubsidy_Limited = TT2._Subsidy_Limited.ToString();
                        vSubsidy_Student = TT2._Subsidy_Student.ToString();
                        vSubsidy_Social = TT2._Subsidy_Social.ToString();
                        vSubsidy_MRT = TT2._Subsidy_MRT.ToString();
                        vSubsidy_Area = TT2._Subsidy_Area.ToString();
                        vSubsidy_Highway = TT2._Subsidy_Highway.ToString();
                        vSubsidy_Others = TT2._Subsidy_Others.ToString();
                        vSubsidy_Dift = TT2._Subsidy_Dift.ToString();

                        using (SqlDataSource dsTemp = new SqlDataSource())
                        {
                            if (vConnStr == "")
                            {
                                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                            }
                            dsTemp.ConnectionString = vConnStr;
                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                            if (vRCount == 0)
                            {
                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, TotalKM, " + Environment.NewLine +
                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, CarCount, " + Environment.NewLine +
                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                              "       DailyCostByCar, DailyIncomeByCar, DepCarCount, MonthDays, INC_ZZ, COS301, BuMan, BuDate)" + Environment.NewLine +
                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, @TotalKM, " + Environment.NewLine +
                                              "       @TicketPCount, @TotalAmount, @RealIncome, @Subsidy_8KM, @Subsidy_Limited, @Subsidy_Student, 0, " + Environment.NewLine +
                                              "       @Subsidy_Social, @Subsidy_MRT, @Subsidy_Area, @Subsidy_Highway, @Subsidy_Others, @Subsidy_Dift, " + Environment.NewLine +
                                              "       0, 0, @DepCarCount, @MonthDays, 0, 0, @BuMan, GetDate())";
                                dsTemp.InsertCommand = vSQLStrTemp;
                                dsTemp.InsertParameters.Clear();
                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                dsTemp.InsertParameters.Add(new Parameter("TotalKM", DbType.Double, vTotalKM.ToString()));
                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                dsTemp.InsertParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Dift", DbType.Double, vSubsidy_Dift));
                                dsTemp.InsertParameters.Add(new Parameter("DepCarCount", DbType.Int32, vDepCarCount.ToString()));
                                dsTemp.InsertParameters.Add(new Parameter("MonthDays", DbType.Int32, vMonthDays.ToString()));
                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));

                                try
                                {
                                    dsTemp.Insert();
                                }
                                catch (Exception eMessage)
                                {
                                    Response.Write("<Script language='Javascript'>");
                                    Response.Write("alert('" + eMessage.Message + "')");
                                    Response.Write("</" + "Script>");
                                }
                            }
                            else if (vRCount == 1)
                            {
                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                              "       RealIncome = RealIncome + @RealIncome, Subsidy_8KM = Subsidy_8KM + @Subsidy_8KM, " + Environment.NewLine +
                                              "       Subsidy_Limited = Subsidy_Limited + @Subsidy_Limited, Subsidy_Student = Subsidy_Student + @Subsidy_Student, " + Environment.NewLine +
                                              "       Subsidy_Social = Subsidy_Social + @Subsidy_Social, Subsidy_MRT = Subsidy_MRT + @Subsidy_MRT, " + Environment.NewLine +
                                              "       Subsidy_Area = Subsidy_Area + @Subsidy_Area, Subsidy_Highway = Subsidy_Highway + @Subsidy_Highway, " + Environment.NewLine +
                                              "       Subsidy_Others = Subsidy_Others + @Subsidy_Others, Subsidy_Dift = Subsidy_Dift + @Subsidy_Dift, " + Environment.NewLine +
                                              "       DepCarCount = @DepCarCount, TotalKM = @TotalKM, MonthDays = @MonthDays, " + Environment.NewLine +
                                              "       ModifyMan = @ModifyMan, ModifyDate = GetDate(), LinesGovNo = @LinesGovNo " + Environment.NewLine +
                                              " where IndexNo = @IndexNo ";
                                dsTemp.UpdateCommand = vSQLStrTemp;
                                dsTemp.UpdateParameters.Clear();
                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                dsTemp.UpdateParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Dift", DbType.Double, vSubsidy_Dift));
                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                dsTemp.UpdateParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                dsTemp.UpdateParameters.Add(new Parameter("DepCarCount", DbType.Int32, vDepCarCount.ToString()));
                                dsTemp.UpdateParameters.Add(new Parameter("TotalKM", DbType.Double, vTotalKM.ToString()));
                                dsTemp.UpdateParameters.Add(new Parameter("MonthDays", DbType.Int32, vMonthDays.ToString()));
                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));

                                try
                                {
                                    dsTemp.Update();
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
                    //計算各路線使用車輛數和成本資料
                    using (SqlConnection connCalCost = new SqlConnection())
                    {
                        if (vConnStr == "")
                        {
                            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                        }
                        connCalCost.ConnectionString = vConnStr;
                        vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                        vSQLStrTemp = "select TicketLineNo, DepNo, DepCarCount, sum(TotalKM) TotalKM_T " + Environment.NewLine +
                                      "  from LinesBounds " + Environment.NewLine +
                                      " where CalYM = '" + vCalYM + "' " + Environment.NewLine +
                                      " group by TicketLineNo, DepNo, DepCarCount " + Environment.NewLine +
                                      " order by TicketLineNo, DepNo";
                        SqlCommand cmdCalCost = new SqlCommand(vSQLStrTemp, connCalCost);
                        connCalCost.Open();
                        SqlDataReader drCalCost = cmdCalCost.ExecuteReader();
                        while (drCalCost.Read())
                        {
                            vTicketLineNo = drCalCost["TicketLineNo"].ToString().Trim();
                            vDepNo = drCalCost["DepNo"].ToString().Trim();
                            vTotalKM = (drCalCost["TotalKM_T"].ToString().Trim() != "") ? double.Parse(drCalCost["TotalKM_T"].ToString().Trim()) : 0.0;
                            vDepCarCount = (drCalCost["DepCarCount"].ToString().Trim() != "") ? Int32.Parse(drCalCost["DepCarCount"].ToString().Trim()) : 0;
                            vSQLStrTemp = "select (sum(t.DepKM) - sum(t.DepKM_2)) as RealDepKM " + Environment.NewLine +
                                          "  from( " + Environment.NewLine +
                                          "        select sum(ActualKM) as DepKM, cast(0.0 as float) DepKM_2 " + Environment.NewLine +
                                          "          from RunSheetA " + Environment.NewLine +
                                          "         where DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "           and BUDATE between '" + vDate_S.ToString("yyyy/MM/dd") + " 00:00:00' and '" + vDate_E.ToString("yyyy/MM/dd") + " 23:59:59' " + Environment.NewLine +
                                          "         union all " + Environment.NewLine +
                                          "        select cast(0.0 as float) DepKM, sum(ActualKM) DepKM_2 " + Environment.NewLine +
                                          "          from RunSheetB " + Environment.NewLine +
                                          "         where AssignNo in (select AssignNo from RunSheetA where DepNo = '" + vDepNo + "' and BUDATE between '" + vDate_S.ToString("yyyy/MM/dd") + " 00:00:00' and '" + vDate_E.ToString("yyyy/MM/dd") + " 23:59:59') " + Environment.NewLine +
                                          "           and isnull(ReduceReason, '') = '' " + Environment.NewLine +
                                          "           and LinesNo = '99990' " + Environment.NewLine +
                                          "       ) t";
                            vTempFloat = double.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RealDepKM").Trim());
                            vCarCount = (vTotalKM / vTempFloat) * (double)vDepCarCount;
                            vSQLStrTemp = "select (SubjectAMT + OtherAMT + ShareAMT) AMT_COS301 " + Environment.NewLine +
                                          "  from DepMonthIncome " + Environment.NewLine +
                                          " where IncomeYM = '" + vCalYM + "' " + Environment.NewLine +
                                          "   and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "   and ChartBarCode = 'COS301' ";
                            vCOS301 = double.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "AMT_COS301"), out vTempFloat) ? vTempFloat : 0.0;
                            vSQLStrTemp = "select (SubjectAMT + OtherAMT + ShareAMT) AMT_INCZZ " + Environment.NewLine +
                                          "  from DepMonthIncome " + Environment.NewLine +
                                          " where IncomeYM = '" + vCalYM + "' " + Environment.NewLine +
                                          "   and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                          "   and ChartBarCode = 'INC ZZ' ";
                            vINC_ZZ = double.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "AMT_INCZZ"), out vTempFloat) ? vTempFloat : 0.0;
                            vTempFloat = (vCarCount * (double)vMonthDays);
                            vDailyCostByCar = (vTempFloat == 0.0) ? "0" : (vCOS301 / vTempFloat).ToString();
                            vDailyIncomeByCar = (vTempFloat == 0.0) ? "0" : (vINC_ZZ / vTempFloat).ToString();

                            using (SqlDataSource dsCalCost = new SqlDataSource())
                            {
                                dsCalCost.ConnectionString = vConnStr;
                                vSQLStrTemp = "update LinesBounds set CarCount = @CarCount, COS301 = @COS301, INC_ZZ = @INC_ZZ, " + Environment.NewLine +
                                              "                       DailyCostByCar = @DailyCostByCar, DailyIncomeByCar = @DailyIncomeByCar " + Environment.NewLine +
                                              " where TicketLineNo = @TicketLineNo and DepNo = @DepNo";
                                dsCalCost.UpdateCommand = vSQLStrTemp;
                                dsCalCost.UpdateParameters.Clear();
                                dsCalCost.UpdateParameters.Add(new Parameter("CarCount", DbType.Double, vCarCount.ToString()));
                                dsCalCost.UpdateParameters.Add(new Parameter("COS301", DbType.Double, vCOS301.ToString()));
                                dsCalCost.UpdateParameters.Add(new Parameter("INC_ZZ", DbType.Double, vINC_ZZ.ToString()));
                                dsCalCost.UpdateParameters.Add(new Parameter("DailyCostByCar", DbType.Double, vDailyCostByCar.ToString()));
                                dsCalCost.UpdateParameters.Add(new Parameter("DailyIncomeByCar", DbType.Double, vDailyIncomeByCar.ToString()));
                                dsCalCost.UpdateParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                dsCalCost.UpdateParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                try
                                {
                                    dsCalCost.Update();
                                }
                                catch (Exception eMessage)
                                {
                                    Response.Write("<Script language='Javascript'>");
                                    Response.Write("alert(" + eMessage.Message + ")");
                                    Response.Write("</" + "Script>");
                                }
                            }
                        }
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('無資料')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// 轉入學生卡資料
        /// </summary>
        private void ImportFromExcel_Student()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";
                int SheetCount = 0;
                int vTempINT = 0;
                Double vTempFloat = 0.0;
                string vTempStr = "";

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vTicketLineNo = "";
                string vLinesGovNo = "";
                string vTicketType = "02"; //學生卡是 02
                string vTicketPCount = "";
                string vTotalAmount = "";
                string vRealIncome = "";
                string vSubsidy_8KM = "";
                string vSubsidy_Limited = "";
                string vSubsidy_Student = "";
                string vSubsidy_MRT = "";
                string vSubsidy_Area = "";
                string vSubsidy_Highway = "";
                string vSubsidy_Others = "";
                string vDepNo = "";
                string[] vaTempArray;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        SheetCount = wbExcel_H.NumberOfSheets;
                        for (int vSC = 0; vSC < SheetCount; vSC++) //逐一取回工作表
                        {
                            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(vSC);
                            for (int i = sheetExcel_H.FirstRowNum + 2; i <= sheetExcel_H.LastRowNum; i++)
                            {
                                HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                if ((vRowTemp_H != null) && (vRowTemp_H.Cells.Count >= 15))
                                {
                                    vTicketLineNo = (Int32.TryParse(vRowTemp_H.Cells[0].ToString().Trim(), out vTempINT)) ? vTempINT.ToString("D6") : ""; //驗票機路線代碼
                                    if (vTicketLineNo != "") //如果對應不到驗票機路線代碼就不轉入
                                    {
                                        vSQLStrTemp = "select (DepNo + ',' + GOVLinesNo) LineData from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                        vaTempArray = vTempStr.Split(',');
                                        vDepNo = vaTempArray[0].ToString().Trim();
                                        //vIndexNo = vCalYM.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        vIndexNo = vCalYM.Trim() + vDepNo.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        //vLinesGovNo = vRowTemp_H.Cells[1].ToString().Trim(); //政府路線編碼
                                        vLinesGovNo = vaTempArray[1].ToString().Trim();
                                        vTicketPCount = (Int32.TryParse(vRowTemp_H.Cells[3].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //刷卡量
                                        vTotalAmount = (Double.TryParse(vRowTemp_H.Cells[4].ToString().Trim(), out vTempFloat)) ? vTempFloat.ToString() : "0.0"; //應收總金額
                                        vRealIncome = (Int32.TryParse(vRowTemp_H.Cells[5].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //實際營收金額
                                        vSubsidy_8KM = (Int32.TryParse(vRowTemp_H.Cells[6].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //基本里程補貼金額
                                        vSubsidy_Limited = (Int32.TryParse(vRowTemp_H.Cells[7].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //票價上限補貼金額
                                        vSubsidy_Student = (Int32.TryParse(vRowTemp_H.Cells[8].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //學生卡25折優惠
                                        vSubsidy_MRT = (Int32.TryParse(vRowTemp_H.Cells[9].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //捷運轉乘優惠補助
                                        vSubsidy_Area = (Int32.TryParse(vRowTemp_H.Cells[10].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //A21-A23區間優惠
                                        vSubsidy_Highway = (Int32.TryParse(vRowTemp_H.Cells[11].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //國道及台鐵轉乘優惠
                                        vSubsidy_Others = (Int32.TryParse(vRowTemp_H.Cells[12].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //其他補貼

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, Subsidy_Social, BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, @RealIncome, @Subsidy_8KM, @Subsidy_Limited, @Subsidy_Student, " + Environment.NewLine +
                                                              "       @Subsidy_MRT, @Subsidy_Area, @Subsidy_Highway, @Subsidy_Others, 0, 0, @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       RealIncome = RealIncome + @RealIncome, Subsidy_8KM = Subsidy_8KM + @Subsidy_8KM, " + Environment.NewLine +
                                                              "       Subsidy_Limited = Subsidy_Limited + @Subsidy_Limited, Subsidy_Student = Subsidy_Student + @Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_MRT = Subsidy_MRT + @Subsidy_MRT, Subsidy_Area = Subsidy_Area + @Subsidy_Area, " + Environment.NewLine +
                                                              "       Subsidy_Highway = Subsidy_Highway + @Subsidy_Highway, Subsidy_Others = Subsidy_Others + @Subsidy_Others, " + Environment.NewLine +
                                                              "       ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2010 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        SheetCount = wbExcel_X.NumberOfSheets;
                        for (int vSC = 0; vSC < SheetCount; vSC++) //逐一取回工作表
                        {
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(vSC);
                            for (int i = sheetExcel_X.FirstRowNum + 2; i <= sheetExcel_X.LastRowNum; i++)
                            {
                                XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                if ((vRowTemp_X != null) && (vRowTemp_X.Cells.Count >= 15))
                                {
                                    vTicketLineNo = (Int32.TryParse(vRowTemp_X.Cells[0].ToString().Trim(), out vTempINT)) ? vTempINT.ToString("D6") : ""; //驗票機路線代碼
                                    if (vTicketLineNo != "") //如果對應不到驗票機路線代碼就不轉入
                                    {
                                        vSQLStrTemp = "select (DepNo + ',' + GOVLinesNo) LineData from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                        vaTempArray = vTempStr.Split(',');
                                        vDepNo = vaTempArray[0].ToString().Trim();
                                        vIndexNo = vCalYM.Trim() + vDepNo.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        vLinesGovNo = vaTempArray[1].ToString().Trim();
                                        //vIndexNo = vCalYM.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        //vLinesGovNo = vRowTemp_X.Cells[1].ToString().Trim(); //政府路線編碼
                                        vTicketPCount = (Int32.TryParse(vRowTemp_X.Cells[3].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //刷卡量
                                        vTotalAmount = (Double.TryParse(vRowTemp_X.Cells[4].ToString().Trim(), out vTempFloat)) ? vTempFloat.ToString() : "0.0"; //應收總金額
                                        vRealIncome = (Int32.TryParse(vRowTemp_X.Cells[5].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //實際營收金額
                                        vSubsidy_8KM = (Int32.TryParse(vRowTemp_X.Cells[6].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //基本里程補貼金額
                                        vSubsidy_Limited = (Int32.TryParse(vRowTemp_X.Cells[7].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //票價上限補貼金額
                                        vSubsidy_Student = (Int32.TryParse(vRowTemp_X.Cells[8].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //學生卡25折優惠
                                        vSubsidy_MRT = (Int32.TryParse(vRowTemp_X.Cells[10].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //捷運轉乘優惠補助
                                        vSubsidy_Area = (Int32.TryParse(vRowTemp_X.Cells[11].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //A21-A23區間優惠
                                        vSubsidy_Highway = (Int32.TryParse(vRowTemp_X.Cells[12].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //國道及台鐵轉乘優惠
                                        vSubsidy_Others = (Int32.TryParse(vRowTemp_X.Cells[13].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0"; //其他補貼

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                                              "       BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, @RealIncome, @Subsidy_8KM, @Subsidy_Limited, @Subsidy_Student, " + Environment.NewLine +
                                                              "       0, @Subsidy_MRT, @Subsidy_Area, @Subsidy_Highway, @Subsidy_Others, 0, " + Environment.NewLine +
                                                              "       @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       RealIncome = RealIncome + @RealIncome, Subsidy_8KM = Subsidy_8KM + @Subsidy_8KM, " + Environment.NewLine +
                                                              "       Subsidy_Limited = Subsidy_Limited + @Subsidy_Limited, Subsidy_Student = Subsidy_Student + @Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_MRT = Subsidy_MRT + @Subsidy_MRT, Subsidy_Area = Subsidy_Area + @Subsidy_Area, " + Environment.NewLine +
                                                              "       Subsidy_Highway = Subsidy_Highway + @Subsidy_Highway, Subsidy_Others = Subsidy_Others + @Subsidy_Others, " + Environment.NewLine +
                                                              "       ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_8KM", DbType.Int32, vSubsidy_8KM));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Limited", DbType.Int32, vSubsidy_Limited));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Student", DbType.Int32, vSubsidy_Student));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_MRT", DbType.Int32, vSubsidy_MRT));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Area", DbType.Int32, vSubsidy_Area));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Others", DbType.Int32, vSubsidy_Others));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
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
        }

        /// <summary>
        /// 轉入敬老愛心卡資料
        /// </summary>
        private void ImportFromExcel_Older()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";
                int SheetCount = 0;
                int vTempINT = 0;

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vTicketLineNo = "";
                string vLinesGovNo = "";
                string vTicketType = "";
                string vTicketPCount = "";
                string vTotalAmount = "";
                string vSubsidy_Social = "";
                string vDepNo = "";
                string vTempStr = "";
                string[] vaTempArray;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        SheetCount = wbExcel_H.NumberOfSheets;
                        for (int vSC = 0; vSC < SheetCount; vSC++) //逐一取回工作表
                        {
                            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(vSC);
                            for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                            {
                                HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                if ((vRowTemp_H != null) && (vRowTemp_H.Cells.Count >= 9))
                                {
                                    vTicketLineNo = (Int32.TryParse(vRowTemp_H.Cells[0].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : ""; //驗票機路線代碼
                                    if (vTicketLineNo != "") //如果對應不到驗票機路線代碼就不轉入
                                    {
                                        vSQLStrTemp = "select (DepNo + ',' + GOVLinesNo) LineData from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                        vaTempArray = vTempStr.Split(',');
                                        vDepNo = vaTempArray[0].ToString().Trim();
                                        vIndexNo = vCalYM.Trim() + vDepNo.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        vLinesGovNo = (vaTempArray.Length > 1) ? vaTempArray[1].ToString().Trim() : "";
                                        //先轉老人卡資料
                                        vTicketType = "04";
                                        vTicketPCount = vRowTemp_H.Cells[1].NumericCellValue.ToString();
                                        vTotalAmount = vRowTemp_H.Cells[7].NumericCellValue.ToString();
                                        vSubsidy_Social = vRowTemp_H.Cells[7].NumericCellValue.ToString();

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                                              "       BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, 0, 0, 0, 0, @Subsidy_Social, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                              "       @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, Int32.Parse(vTicketLineNo).ToString("D6")));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    string vErrorMessage = eMessage.Message + Environment.NewLine +
                                                                           "IndexNo：" + vIndexNo + Environment.NewLine +
                                                                           "CalYM：" + vCalYM + Environment.NewLine +
                                                                           "TicketType：" + vTicketType + Environment.NewLine +
                                                                           "TicketLineNo：" + vTicketLineNo + Environment.NewLine +
                                                                           "LinesGovNo：" + vLinesGovNo + Environment.NewLine +
                                                                           "DepNo：" + vDepNo + Environment.NewLine +
                                                                           "TicketPCount：" + vTicketPCount + Environment.NewLine +
                                                                           "TotalAmount：" + vTotalAmount + Environment.NewLine +
                                                                           "Subsidy_Social：" + vSubsidy_Social;

                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + vErrorMessage + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       Subsidy_Social = Subsidy_Social + @Subsidy_Social, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    string vErrorMessage = eMessage.Message + Environment.NewLine +
                                                                           "TicketPCount：" + vTicketPCount + Environment.NewLine +
                                                                           "TotalAmount：" + vTotalAmount + Environment.NewLine +
                                                                           "Subsidy_Social：" + vSubsidy_Social + Environment.NewLine +
                                                                           "IndexNo：" + vIndexNo;
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + vErrorMessage + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
                                        }

                                        //再轉愛心卡資料
                                        vTicketType = "03";
                                        vTicketPCount = vRowTemp_H.Cells[2].NumericCellValue.ToString();
                                        vTotalAmount = vRowTemp_H.Cells[8].NumericCellValue.ToString();
                                        vSubsidy_Social = vRowTemp_H.Cells[8].NumericCellValue.ToString();

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                                              "       BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, 0, 0, 0, 0, @Subsidy_Social, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                              "       @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    string vErrorMessage = eMessage.Message + Environment.NewLine +
                                                                           "IndexNo：" + vIndexNo + Environment.NewLine +
                                                                           "CalYM：" + vCalYM + Environment.NewLine +
                                                                           "TicketType：" + vTicketType + Environment.NewLine +
                                                                           "TicketLineNo：" + vTicketLineNo + Environment.NewLine +
                                                                           "LinesGovNo：" + vLinesGovNo + Environment.NewLine +
                                                                           "DepNo：" + vDepNo + Environment.NewLine +
                                                                           "TicketPCount：" + vTicketPCount + Environment.NewLine +
                                                                           "TotalAmount：" + vTotalAmount + Environment.NewLine +
                                                                           "Subsidy_Social：" + vSubsidy_Social;

                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + vErrorMessage + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       Subsidy_Social = Subsidy_Social + @Subsidy_Social, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    string vErrorMessage = eMessage.Message + Environment.NewLine +
                                                                           "TicketPCount：" + vTicketPCount + Environment.NewLine +
                                                                           "TotalAmount：" + vTotalAmount + Environment.NewLine +
                                                                           "Subsidy_Social：" + vSubsidy_Social + Environment.NewLine +
                                                                           "IndexNo：" + vIndexNo;
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + vErrorMessage + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2010 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        SheetCount = wbExcel_X.NumberOfSheets;
                        for (int vSC = 0; vSC < SheetCount; vSC++) //逐一取回工作表
                        {
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(vSC);
                            for (int i = sheetExcel_X.FirstRowNum + 7; i <= sheetExcel_X.LastRowNum; i++)
                            {
                                XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                if ((vRowTemp_X != null) && (vRowTemp_X.Cells.Count >= 9))
                                {
                                    vTicketLineNo = (Int32.TryParse(vRowTemp_X.Cells[0].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : ""; //驗票機路線代碼
                                    if (vTicketLineNo != "") //如果對應不到驗票機路線代碼就不轉入
                                    {
                                        vSQLStrTemp = "select (DepNo + ',' + GOVLinesNo) LineData from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                        vaTempArray = vTempStr.Split(',');
                                        vDepNo = vaTempArray[0].ToString().Trim();
                                        vIndexNo = vCalYM.Trim() + vDepNo.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        vLinesGovNo = (vaTempArray.Length > 1) ? vaTempArray[1].ToString().Trim() : "";
                                        //先轉老人卡資料
                                        vTicketType = "04";
                                        vTicketPCount = vRowTemp_X.Cells[1].ToString().Trim();
                                        vTotalAmount = vRowTemp_X.Cells[7].ToString().Trim();
                                        vSubsidy_Social = vRowTemp_X.Cells[7].ToString().Trim();

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                                              "       BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, 0, 0, 0, 0, @Subsidy_Social, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                              "       @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, Int32.Parse(vTicketLineNo).ToString("D6")));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       Subsidy_Social = Subsidy_Social + @Subsidy_Social, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
                                        }

                                        //再轉愛心卡資料
                                        vTicketType = "03";
                                        vTicketPCount = vRowTemp_X.Cells[2].ToString().Trim();
                                        vTotalAmount = vRowTemp_X.Cells[8].ToString().Trim();
                                        vSubsidy_Social = vRowTemp_X.Cells[8].ToString().Trim();

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                                              "       BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, 0, 0, 0, 0, @Subsidy_Social, 0, 0, 0, 0, 0, " + Environment.NewLine +
                                                              "       @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       Subsidy_Social = Subsidy_Social + @Subsidy_Social, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Social", DbType.Int32, vSubsidy_Social));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
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
        }

        /// <summary>
        /// 轉入國道台鐵補貼
        /// </summary>
        private void ImportFromExcel_Highway()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";
                int SheetCount = 0;
                int vTempINT = 0;

                string vPrice_City = (eTicketPrice_City.Text.Trim() != "") ? eTicketPrice_City.Text.Trim() : "0";
                string vPrice_Highway = (eTicketPrice_Highway.Text.Trim() != "") ? eTicketPrice_Highway.Text.Trim() : "0";

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vTicketLineNo = "";
                string vLinesGovNo = "";
                string vTicketType = "";
                string vTicketPCount = "";
                string vTotalAmount = "";
                string vRealIncome = "";
                string vSubsidy_Highway = "";
                string vDepNo = "";
                string vTempStr = "";
                string[] vaTempArray;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        SheetCount = wbExcel_H.NumberOfSheets;
                        for (int vSC = 0; vSC < SheetCount; vSC++) //逐一取回工作表
                        {
                            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(vSC);
                            for (int i = sheetExcel_H.FirstRowNum + 3; i <= sheetExcel_H.LastRowNum; i++)
                            {
                                HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                if ((vRowTemp_H != null) && (vRowTemp_H.Cells.Count >= 13))
                                {
                                    vTicketLineNo = (Int32.TryParse(vRowTemp_H.Cells[4].ToString().Trim(), out vTempINT)) ? vTempINT.ToString("D6") : ""; //驗票機路線代碼
                                    if (vTicketLineNo != "") //如果對應不到驗票機路線代碼就不轉入
                                    {
                                        vSubsidy_Highway = vRowTemp_H.Cells[12].ToString().Trim();
                                        switch (vRowTemp_H.Cells[3].ToString().Trim())
                                        {
                                            default:
                                                vTicketType = "";
                                                break;
                                            case "一般卡(含員工卡)":
                                            case "一般卡":
                                            case "1":
                                                vTicketType = "01";
                                                break;
                                            case "學生卡":
                                            case "6":
                                                vTicketType = ((ddlSourceTyoe.SelectedValue == "04") && (Int32.Parse(vSubsidy_Highway) == Int32.Parse(vPrice_Highway))) ? "01" :
                                                              ((ddlSourceTyoe.SelectedValue == "05") && (Int32.Parse(vSubsidy_Highway) == Int32.Parse(vPrice_City))) ? "01" :
                                                              "02";
                                                break;
                                            case "愛心卡":
                                            case "4":
                                                vTicketType = "03";
                                                break;
                                            case "敬老卡":
                                            case "2":
                                            case "3":
                                                vTicketType = "04";
                                                break;
                                            case "愛陪卡(有陪)":
                                                vTicketType = "05";
                                                break;
                                            case "愛陪卡(未陪)":
                                            case "愛陪卡(沒陪)":
                                                vTicketType = "06";
                                                break;
                                            case "5":
                                                vTicketType = ((ddlSourceTyoe.SelectedValue == "04") && (Int32.Parse(vSubsidy_Highway) == Int32.Parse(vPrice_Highway))) ? "06" :
                                                              ((ddlSourceTyoe.SelectedValue == "05") && (Int32.Parse(vSubsidy_Highway) == Int32.Parse(vPrice_City))) ? "06" :
                                                              "05";
                                                break;
                                        }
                                        //vIndexNo = vCalYM.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        //vSQLStrTemp = "select top 1 LinesGovNo from Lines where TicketLineNo = '" + vTicketLineNo + "' order by LinesNo ";
                                        //vLinesGovNo = PF.GetValue(vConnStr, vSQLStrTemp, "LinesGovNo");
                                        vSQLStrTemp = "select (DepNo + ',' + GOVLinesNo) LineData from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                        vaTempArray = vTempStr.Split(',');
                                        vDepNo = vaTempArray[0].ToString().Trim();
                                        vIndexNo = vCalYM.Trim() + vDepNo.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        vLinesGovNo = vaTempArray[1].ToString().Trim();

                                        vTicketPCount = "1";
                                        vTotalAmount = vRowTemp_H.Cells[10].ToString().Trim();
                                        vRealIncome = vRowTemp_H.Cells[11].ToString().Trim();

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                                              "       BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, @RealIncome, 0, 0, 0, 0, 0, 0, @Subsidy_Highway, 0, 0, " + Environment.NewLine +
                                                              "       @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       RealIncome = RealIncome + @RealIncome, Subsidy_Highway = Subsidy_Highway + @Subsidy_Highway, " + Environment.NewLine +
                                                              "       ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2007 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        SheetCount = wbExcel_X.NumberOfSheets;
                        for (int vSC = 0; vSC < SheetCount; vSC++) //逐一取回工作表
                        {
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(vSC);
                            for (int i = sheetExcel_X.FirstRowNum + 3; i <= sheetExcel_X.LastRowNum; i++)
                            {
                                XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                if ((vRowTemp_X != null) && (vRowTemp_X.Cells.Count >= 13))
                                {
                                    vTicketLineNo = (Int32.TryParse(vRowTemp_X.Cells[4].ToString().Trim(), out vTempINT)) ? vTempINT.ToString("D6") : ""; //驗票機路線代碼
                                    if (vTicketLineNo != "") //如果對應不到驗票機路線代碼就不轉入
                                    {
                                        vSubsidy_Highway = vRowTemp_X.Cells[12].ToString().Trim();
                                        switch (vRowTemp_X.Cells[3].ToString().Trim())
                                        {
                                            default:
                                                vTicketType = "";
                                                break;
                                            case "一般卡(含員工卡)":
                                            case "一般卡":
                                            case "1":
                                                vTicketType = "01";
                                                break;
                                            case "學生卡":
                                            case "6":
                                                vTicketType = ((ddlSourceTyoe.SelectedValue == "04") && (Int32.Parse(vSubsidy_Highway) == Int32.Parse(vPrice_Highway))) ? "01" :
                                                              ((ddlSourceTyoe.SelectedValue == "05") && (Int32.Parse(vSubsidy_Highway) == Int32.Parse(vPrice_City))) ? "01" :
                                                              "02";
                                                break;
                                            case "愛心卡":
                                            case "4":
                                                vTicketType = "03";
                                                break;
                                            case "敬老卡":
                                            case "2":
                                            case "3":
                                                vTicketType = "04";
                                                break;
                                            case "愛陪卡(有陪)":
                                                vTicketType = "05";
                                                break;
                                            case "愛陪卡(未陪)":
                                            case "愛陪卡(沒陪)":
                                                vTicketType = "06";
                                                break;
                                            case "5":
                                                vTicketType = ((ddlSourceTyoe.SelectedValue == "04") && (Int32.Parse(vSubsidy_Highway) == Int32.Parse(vPrice_Highway))) ? "06" :
                                                              ((ddlSourceTyoe.SelectedValue == "05") && (Int32.Parse(vSubsidy_Highway) == Int32.Parse(vPrice_City))) ? "06" :
                                                              "05";
                                                break;
                                        }
                                        //vIndexNo = vCalYM.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        //vSQLStrTemp = "select top 1 LinesGovNo from Lines where TicketLineNo = '" + vTicketLineNo + "' order by LinesNo ";
                                        //vLinesGovNo = PF.GetValue(vConnStr, vSQLStrTemp, "LinesGovNo");
                                        vSQLStrTemp = "select (DepNo + ',' + GOVLinesNo) LineData from LinesNoChart where TicketLinesNo = '" + vTicketLineNo + "' ";
                                        vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                        vaTempArray = vTempStr.Split(',');
                                        vDepNo = vaTempArray[0].ToString().Trim();
                                        vIndexNo = vCalYM.Trim() + vDepNo.Trim() + vTicketType.Trim() + vTicketLineNo.Trim(); //序號
                                        vLinesGovNo = vaTempArray[1].ToString().Trim();

                                        vTicketPCount = "1";
                                        vTotalAmount = vRowTemp_X.Cells[10].ToString().Trim();
                                        vRealIncome = vRowTemp_X.Cells[11].ToString().Trim();

                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            vSQLStrTemp = "select count(IndexNo) as RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                                            if (vRCount == 0)
                                            {
                                                vSQLStrTemp = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, DepNo, " + Environment.NewLine +
                                                              "       TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, " + Environment.NewLine +
                                                              "       Subsidy_Social, Subsidy_MRT, Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, " + Environment.NewLine +
                                                              "       BuMan, BuDate)" + Environment.NewLine +
                                                              "values (@IndexNo, @CalYM, @TicketType, @TicketLineNo, @LinesGovNo, @DepNo, " + Environment.NewLine +
                                                              "       @TicketPCount, @TotalAmount, @RealIncome, 0, 0, 0, 0, 0, 0, @Subsidy_Highway, 0, 0, " + Environment.NewLine +
                                                              "       @BuMan, GetDate())";
                                                dsTemp.InsertCommand = vSQLStrTemp;
                                                dsTemp.InsertParameters.Clear();
                                                dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketType", DbType.String, vTicketType));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                                dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGovNo));
                                                dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                                dsTemp.InsertParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.InsertParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.InsertParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                                try
                                                {
                                                    dsTemp.Insert();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else if (vRCount == 1)
                                            {
                                                vSQLStrTemp = "update LinesBounds set TicketPCount = TicketPCount + @TicketPCount, TotalAmount = TotalAmount + @TotalAmount, " + Environment.NewLine +
                                                              "       RealIncome = RealIncome + @RealIncome, Subsidy_Highway = Subsidy_Highway + @Subsidy_Highway, " + Environment.NewLine +
                                                              "       ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                              " where IndexNo = @IndexNo ";
                                                dsTemp.UpdateCommand = vSQLStrTemp;
                                                dsTemp.UpdateParameters.Clear();
                                                dsTemp.UpdateParameters.Add(new Parameter("TicketPCount", DbType.Int32, vTicketPCount));
                                                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                                                dsTemp.UpdateParameters.Add(new Parameter("RealIncome", DbType.Int32, vRealIncome));
                                                dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Highway", DbType.Int32, vSubsidy_Highway));
                                                dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                try
                                                {
                                                    dsTemp.Update();
                                                }
                                                catch (Exception eMessage)
                                                {
                                                    Response.Write("<Script language='Javascript'>");
                                                    Response.Write("alert('" + eMessage.Message + "')");
                                                    Response.Write("</" + "Script>");
                                                }
                                            }
                                            else
                                            {
                                                vSQLStrTemp = "";
                                            }
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
        }

        /// <summary>
        /// 國道準運價
        /// </summary>
        private void ImportFromExcel_Highway2()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";
                string vTempStr = "";
                string[] vaTempArray;
                string vLinesName_Old = "";
                string vLinesName_New = "";
                string vTitle_Temp = "";

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vDepNo_Temp = "";
                string vLinesNo = "";
                string vTicketLinesNo = "";
                string vLinesGOVNo = "";
                string vSubsidy_Highway2 = "";
                string vTicketType = "01"; //國道準運價先都轉到一般卡
                int vSubsidy_Temp = 0;

                switch (vExtName)
                {
                    case ".xls":
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        int vSheetRowCount_H = sheetExcel_H.LastRowNum;
                        HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(0);
                        int vColCount_H = vRowExcel_H.LastCellNum;
                        for (int i = 0; i <= vSheetRowCount_H; i++)
                        {
                            vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            vColCount_H = (vColCount_H < vRowExcel_H.Cells.Count) ? vRowExcel_H.Cells.Count : vColCount_H;
                        }
                        for (int i = 0; i < vSheetRowCount_H; i++)
                        {
                            vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            //DataRow drTemp = dtTemp.NewRow();
                            vLinesName_New = (vRowExcel_H.Cells.Count > 1) ? vRowExcel_H.Cells[1].ToString().Trim() : "";
                            if ((vLinesName_New != vLinesName_Old) && (vLinesName_New != "") && (vLinesName_New.IndexOf("【") >= 0) && (vLinesName_New.IndexOf("(範本)") < 0) ||
                                (vRowExcel_H.Cells[0].ToString().Trim() == "電子票證申請補貼張數"))
                            {
                                vLinesGOVNo = (vLinesName_Old != "") ? vLinesName_Old.Substring(vLinesName_Old.IndexOf("【") + 1, vLinesName_Old.IndexOf("】") - vLinesName_Old.IndexOf("【") - 1) : "";
                                if ((vTitle_Temp == "路線小計") && (vLinesGOVNo != ""))
                                {
                                    vSQLStrTemp = "select top 1 (DepNo + ',' + ERPLinesNo + ',' + TicketLinesNo) LinesData " + Environment.NewLine +
                                                  "  from LinesNoChart " + Environment.NewLine +
                                                  " where GOVLinesNo = '" + vLinesGOVNo + "' " + Environment.NewLine +
                                                  " order by TicketLinesNo";
                                    vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LinesData");
                                    vaTempArray = vTempStr.Split(',');
                                    vDepNo_Temp = vaTempArray[0].Trim();
                                    vLinesNo = vaTempArray[1].Trim();
                                    vTicketLinesNo = vaTempArray[2].Trim();
                                    vIndexNo = vCalYM + vDepNo_Temp + vTicketType + Int32.Parse(vTicketLinesNo).ToString("D6");
                                    vSubsidy_Temp = (Int32)Math.Round((double.Parse(vSubsidy_Highway2) / 1.05), 0, MidpointRounding.AwayFromZero);
                                    vSQLStrTemp = "select count(IndexNo) RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount")) == 1)
                                    {
                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                   "   set Subsidy_Highway2 = SubsidyHighway2 + @Subsidy_Highway2, ModifyMan = @Modifyman, Modifydate = GetDate() " + Environment.NewLine +
                                                                   " where IndexNo = @IndexNo";
                                            dsTemp.UpdateParameters.Clear();
                                            dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Highway2", DbType.Double, vSubsidy_Temp.ToString()));
                                            dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                            dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.Update();
                                        }
                                    }
                                }
                                vLinesName_Old = vLinesName_New.Trim();
                                vLinesName_New = "";
                            }
                            vSubsidy_Highway2 = ((vRowExcel_H.Cells.Count >= 32) && (vRowExcel_H.Cells[31].CellType == NPOI.SS.UserModel.CellType.Numeric)) ? vRowExcel_H.Cells[31].NumericCellValue.ToString() :
                                                ((vRowExcel_H.Cells.Count >= 32) && (vRowExcel_H.Cells[31].CellType == NPOI.SS.UserModel.CellType.Formula)) ? vRowExcel_H.Cells[31].NumericCellValue.ToString() :
                                                "0";
                            vTitle_Temp = (vRowExcel_H.Cells.Count > 4) ? vRowExcel_H.Cells[4].ToString().Trim() : "";
                        }
                        break;

                    case ".xlsx":
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        int vSheetRowCount_X = sheetExcel_X.LastRowNum;
                        XSSFRow vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(0);
                        int vColCount_X = vRowExcel_X.LastCellNum;
                        for (int i = 0; i <= vSheetRowCount_X; i++)
                        {
                            vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            vColCount_X = (vColCount_X < vRowExcel_X.Cells.Count) ? vRowExcel_X.Cells.Count : vColCount_X;
                        }
                        for (int i = 0; i < vSheetRowCount_X; i++)
                        {
                            vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            vLinesName_New = (vRowExcel_X.Cells.Count > 1) ? vRowExcel_X.Cells[1].ToString().Trim() : "";
                            if ((vLinesName_New != vLinesName_Old) && (vLinesName_New != "") && (vLinesName_New.IndexOf("【") >= 0) && (vLinesName_New.IndexOf("(範本)") < 0) ||
                                (vRowExcel_X.Cells[0].ToString().Trim() == "電子票證申請補貼張數"))
                            {
                                vLinesGOVNo = (vLinesName_Old != "") ? vLinesName_Old.Substring(vLinesName_Old.IndexOf("【") + 1, vLinesName_Old.IndexOf("】") - vLinesName_Old.IndexOf("【") - 1) : "";
                                if ((vTitle_Temp == "路線小計") && (vLinesGOVNo != ""))
                                {
                                    vSQLStrTemp = "select top 1 (DepNo + ',' + ERPLinesNo + ',' + TicketLinesNo) LinesData " + Environment.NewLine +
                                                  "  from LinesNoChart " + Environment.NewLine +
                                                  " where GOVLinesNo = '" + vLinesGOVNo + "' " + Environment.NewLine +
                                                  " order by TicketLinesNo";
                                    vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LinesData");
                                    vaTempArray = vTempStr.Split(',');
                                    vDepNo_Temp = vaTempArray[0].Trim();
                                    vLinesNo = vaTempArray[1].Trim();
                                    vTicketLinesNo = vaTempArray[2].Trim();
                                    vIndexNo = vCalYM + vDepNo_Temp + vTicketType + Int32.Parse(vTicketLinesNo).ToString("D6");
                                    vSubsidy_Temp = (Int32)Math.Round((double.Parse(vSubsidy_Highway2) / 1.05), 0, MidpointRounding.AwayFromZero);
                                    vSQLStrTemp = "select count(IndexNo) RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount")) == 1)
                                    {
                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                   "   set Subsidy_Xighway2 = SubsidyHighway2 + @Subsidy_Xighway2, ModifyMan = @Modifyman, Modifydate = GetDate() " + Environment.NewLine +
                                                                   " where IndexNo = @IndexNo";
                                            dsTemp.UpdateParameters.Clear();
                                            dsTemp.UpdateParameters.Add(new Parameter("Subsidy_Xighway2", DbType.Double, vSubsidy_Temp.ToString()));
                                            dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                            dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.Update();
                                        }
                                    }
                                }
                                vLinesName_Old = vLinesName_New.Trim();
                                vLinesName_New = "";
                            }
                            vSubsidy_Highway2 = ((vRowExcel_X.Cells.Count >= 32) && (vRowExcel_X.Cells[31].CellType == NPOI.SS.UserModel.CellType.Numeric)) ? vRowExcel_X.Cells[31].NumericCellValue.ToString() :
                                                ((vRowExcel_X.Cells.Count >= 32) && (vRowExcel_X.Cells[31].CellType == NPOI.SS.UserModel.CellType.Formula)) ? vRowExcel_X.Cells[31].NumericCellValue.ToString() :
                                                "0";
                            vTitle_Temp = (vRowExcel_X.Cells.Count > 4) ? vRowExcel_X.Cells[4].ToString().Trim() : "";
                        }
                        break;

                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;
                }
            }
        }

        /// <summary>
        /// 國小升國中補貼
        /// </summary>
        private void ImportFromExcel_NewHighschool()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";
                int vTempINT = 0;
                string vTempStr = "";
                string[] vaTempArray;

                string vIndexNo = "";
                //string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vDepNo = "";
                string vLinesNo = "";
                string vTicketLinesNo = "";
                string vLinesGOVNo = "";
                string vSubsidy_Student2 = "";
                string vTicketType = "02"; //國小升國中補貼一定是學生卡
                int vSubsidy_Temp = 0;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        for (int i = sheetExcel_H.FirstRowNum + 3; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if ((vRowTemp_H != null) && (vRowTemp_H.Cells.Count >= 8))
                            {
                                vTicketLinesNo = vRowTemp_H.Cells[0].ToString().Trim();
                                if ((vTicketLinesNo != "") && (Int32.TryParse(vTicketLinesNo, out vTempINT)))
                                {
                                    vSQLStrTemp = "select (DepNo + ',' + ERPLinesNo + ',' + GOVLinesNo) LineData from LinesNoChart where TicketLinesNo = '" + vTicketLinesNo + "' ";
                                    vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                    vaTempArray = vTempStr.Split(',');
                                    vDepNo = vaTempArray[0].ToString().Trim();
                                    vLinesNo = vaTempArray[1].ToString().Trim();
                                    vLinesGOVNo = vaTempArray[2].ToString().Trim();
                                    vSubsidy_Temp = (Int32)Math.Round((Int32.Parse(vRowTemp_H.Cells[7].ToString()) / 2.0), 0, MidpointRounding.AwayFromZero);
                                    vSubsidy_Student2 = vSubsidy_Temp.ToString();
                                    vIndexNo = eImportYear.Text.Trim() + "09" + vDepNo + vTicketType + Int32.Parse(vTicketLinesNo).ToString("D6"); //9 月的資料
                                    vSQLStrTemp = "select Subsidy_Student2 from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (PF.GetValue(vConnStr, vSQLStrTemp, "Subsidy_Student2").Trim() != "")
                                    {
                                        vSQLStrTemp = "update LinesBounds set Subsidy_Student2 = " + vSubsidy_Temp.ToString() + Environment.NewLine +
                                                      " where IndexNo = '" + vIndexNo + "' ";
                                        PF.ExecSQL(vConnStr, vSQLStrTemp);
                                    }
                                    vIndexNo = eImportYear.Text.Trim() + "10" + vDepNo + vTicketType + Int32.Parse(vTicketLinesNo).ToString("D6"); //10 月的資料
                                    vSQLStrTemp = "select Subsidy_Student2 from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (PF.GetValue(vConnStr, vSQLStrTemp, "Subsidy_Student2").Trim() != "")
                                    {
                                        vSQLStrTemp = "update LinesBounds set Subsidy_Student2 = " + vSubsidy_Temp.ToString() + Environment.NewLine +
                                                      " where IndexNo = '" + vIndexNo + "' ";
                                        PF.ExecSQL(vConnStr, vSQLStrTemp);
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2007 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        for (int i = sheetExcel_X.FirstRowNum + 3; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if ((vRowTemp_X != null) && (vRowTemp_X.Cells.Count >= 8))
                            {
                                vTicketLinesNo = vRowTemp_X.Cells[0].ToString().Trim();
                                if ((vTicketLinesNo != "") && (Int32.TryParse(vTicketLinesNo, out vTempINT)))
                                {
                                    vSQLStrTemp = "select (DepNo + ',' + ERPLinesNo + ',' + GOVLinesNo) LineData from LinesNoChart where TicketLinesNo = '" + vTicketLinesNo + "' ";
                                    vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                    vaTempArray = vTempStr.Split(',');
                                    vDepNo = vaTempArray[0].ToString().Trim();
                                    vLinesNo = vaTempArray[1].ToString().Trim();
                                    vLinesGOVNo = vaTempArray[2].ToString().Trim();
                                    vSubsidy_Temp = (Int32)Math.Round((Int32.Parse(vRowTemp_X.Cells[7].ToString()) / 2.0), 0, MidpointRounding.AwayFromZero);
                                    vSubsidy_Student2 = vSubsidy_Temp.ToString();
                                    vIndexNo = eImportYear.Text.Trim() + "09" + vDepNo + vTicketType + Int32.Parse(vTicketLinesNo).ToString("D6"); //9 月的資料
                                    vSQLStrTemp = "select Subsidy_Student2 from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (PF.GetValue(vConnStr, vSQLStrTemp, "Subsidy_Student2").Trim() != "")
                                    {
                                        vSQLStrTemp = "update LinesBounds set Subsidy_Student2 = " + vSubsidy_Temp.ToString() + Environment.NewLine +
                                                      " where IndexNo = '" + vIndexNo + "' ";
                                        PF.ExecSQL(vConnStr, vSQLStrTemp);
                                    }
                                    vIndexNo = eImportYear.Text.Trim() + "10" + vDepNo + vTicketType + Int32.Parse(vTicketLinesNo).ToString("D6"); //10 月的資料
                                    vSQLStrTemp = "select Subsidy_Student2 from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (PF.GetValue(vConnStr, vSQLStrTemp, "Subsidy_Student2").Trim() != "")
                                    {
                                        vSQLStrTemp = "update LinesBounds set Subsidy_Student2 = " + vSubsidy_Temp.ToString() + Environment.NewLine +
                                                      " where IndexNo = '" + vIndexNo + "' ";
                                        PF.ExecSQL(vConnStr, vSQLStrTemp);
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
        }

        /// <summary>
        /// 新北市公車間轉乘
        /// </summary>
        private void ImportFromExcel_NewTaipeiBusTrans()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";
                string vTempStr = "";
                string[] vaTempStr;
                int vTempINT = 0;

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vDepNo = "";
                string vLinesNo = "";
                string vTicketLineNo = "";
                string vLinesGOVNo = "";
                string vSubsidy_BusTran = "";

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if ((vRowExcel_H != null) &&
                                (vRowExcel_H.Cells[0].ToString().Trim() == "桃園客運") &&
                                (vRowExcel_H.Cells[1].ToString().Trim() != "") &&
                                (Int32.TryParse(vRowExcel_H.Cells[1].ToString().Trim(), out vTempINT)))
                            {
                                vTicketLineNo = vRowExcel_H.Cells[1].ToString().Trim();
                                vSQLStrTemp = "select (ERPLinesNo + ',' + GOVLinesNo + ',' + DepNo) as DataStr " + Environment.NewLine +
                                              "  from LinesNoChart " + Environment.NewLine +
                                              " where TicketLinesNo = '" + vTicketLineNo + "' ";
                                vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "DataStr");
                                if (vTempStr.Replace(",", "") != "")
                                {
                                    vaTempStr = vTempStr.Split(',');
                                    vLinesNo = vaTempStr[0].Trim();
                                    vLinesGOVNo = vaTempStr[1].Trim();
                                    vDepNo = vaTempStr[2].Trim();
                                    vIndexNo = vCalYM + vDepNo + "01" + Int32.Parse(vTicketLineNo).ToString("D6");
                                    vSubsidy_BusTran = vRowExcel_H.Cells[22].NumericCellValue.ToString();
                                    vTempINT = (Int32.Parse(vSubsidy_BusTran) != 0) ? (Int32)Math.Round((Double.Parse(vSubsidy_BusTran) / 1.05), 0, MidpointRounding.AwayFromZero) : 0;
                                    vSubsidy_BusTran = vTempINT.ToString();
                                    using (SqlDataSource dsTemp = new SqlDataSource())
                                    {
                                        dsTemp.ConnectionString = vConnStr;
                                        vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                        if (PF.GetValue(vConnStr, vSQLStrTemp, "IndexNo") != "")
                                        {
                                            dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                   "    set Subsidy_BusTran = @Subsidy_BusTran, ModifyDate = GetDate(), ModifyMan = @ModifyMan " + Environment.NewLine +
                                                                   " where IndexNo = @IndexNo";
                                            dsTemp.UpdateParameters.Clear();
                                            dsTemp.UpdateParameters.Add(new Parameter("Subsidy_BusTran", DbType.Double, vSubsidy_BusTran));
                                            dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                            dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.Update();
                                        }
                                        else
                                        {
                                            dsTemp.InsertCommand = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, BuMan, BuDate, DepNo, Subsidy_BusTran)" + Environment.NewLine +
                                                                   "values (@IndexNo, @CalYM, '01', @TicketLineNo, @LinesGovNo, @BuMan, GetDate(), @DepNo, @Subsidy_BusTran)";
                                            dsTemp.InsertParameters.Clear();
                                            dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                            dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, vTicketLineNo));
                                            dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGOVNo));
                                            dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                            dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                            dsTemp.InsertParameters.Add(new Parameter("Subsidy_BusTran", DbType.Double, vSubsidy_BusTran));
                                            dsTemp.Insert();
                                        }
                                    }
                                }
                                else
                                {
                                    Response.Write("<Script language='Javascript'>");
                                    Response.Write("alert('找不到驗票機路線代碼 [" + vTicketLineNo + "] 所對應的 ERP 路線代號及單位')");
                                    Response.Write("</" + "Script>");
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2007 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if ((vRowExcel_X != null) &&
                                (vRowExcel_X.Cells[0].ToString().Trim() == "桃園客運") &&
                                (vRowExcel_X.Cells[1].ToString().Trim() != "") &&
                                (Int32.TryParse(vRowExcel_X.Cells[1].ToString().Trim(), out vTempINT)))
                            {
                                vTicketLineNo = vRowExcel_X.Cells[1].ToString().Trim();
                                vSQLStrTemp = "select (ERPLinesNo + ',' + GOVLinesNo + ',' + DepNo) as DataStr " + Environment.NewLine +
                                              "  from LinesNoChart " + Environment.NewLine +
                                              " where TicketLinesNo = '" + vTicketLineNo + "' ";
                                vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "DataStr");
                                if (vTempStr.Replace(",", "") != "")
                                {
                                    vaTempStr = vTempStr.Split(',');
                                    vLinesNo = vaTempStr[0].Trim();
                                    vLinesGOVNo = vaTempStr[1].Trim();
                                    vDepNo = vaTempStr[2].Trim();
                                    vIndexNo = vCalYM + vDepNo + "01" + Int32.Parse(vTicketLineNo).ToString("D6");
                                    vSubsidy_BusTran = vRowExcel_X.Cells[22].NumericCellValue.ToString();
                                    vTempINT = (Int32.Parse(vSubsidy_BusTran) != 0) ? (Int32)Math.Round((Double.Parse(vSubsidy_BusTran) / 1.05), 0, MidpointRounding.AwayFromZero) : 0;
                                    vSubsidy_BusTran = vTempINT.ToString();
                                    using (SqlDataSource dsTemp = new SqlDataSource())
                                    {
                                        dsTemp.ConnectionString = vConnStr;
                                        vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                        if (PF.GetValue(vConnStr, vSQLStrTemp, "IndexNo") != "")
                                        {
                                            dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                   "    set Subsidy_BusTran = @Subsidy_BusTran, ModifyDate = GetDate(), ModifyMan = @ModifyMan " + Environment.NewLine +
                                                                   " where IndexNo = @IndexNo";
                                            dsTemp.UpdateParameters.Clear();
                                            dsTemp.UpdateParameters.Add(new Parameter("Subsidy_BusTran", DbType.Double, vSubsidy_BusTran));
                                            dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                            dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.Update();
                                        }
                                        else
                                        {
                                            dsTemp.InsertCommand = "insert into LinesBounds (IndexNo, CalYM, TicketType, TicketLineNo, LinesGovNo, BuMan, BuDate, DepNo, Subsidy_BusTran)" + Environment.NewLine +
                                                                   "values (@IndexNo, @CalYM, '01', @TicketLineNo, @LinesGovNo, @BuMan, GetDate(), @DepNo, @Subsidy_BusTran)";
                                            dsTemp.InsertParameters.Clear();
                                            dsTemp.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM));
                                            dsTemp.InsertParameters.Add(new Parameter("TicketLineNo", DbType.String, Int32.Parse(vTicketLineNo).ToString("D6")));
                                            dsTemp.InsertParameters.Add(new Parameter("LinesGovNo", DbType.String, vLinesGOVNo));
                                            dsTemp.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                            dsTemp.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo));
                                            dsTemp.InsertParameters.Add(new Parameter("Subsidy_BusTran", DbType.Double, vSubsidy_BusTran));
                                            dsTemp.Insert();
                                        }
                                    }
                                }
                                else
                                {
                                    Response.Write("<Script language='Javascript'>");
                                    Response.Write("alert('找不到驗票機路線代碼 [" + vTicketLineNo + "] 所對應的 ERP 路線代號及單位')");
                                    Response.Write("</" + "Script>");
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
        }

        /// <summary>
        /// 
        /// </summary>
        private void ImportFromExcel_MRT()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vDepNo_Temp = "22"; //952 固定是桃公站
                string vTicketLinesNo = "952"; //北捷轉乘固定都是952
                string vSubsidy_MRT = "";
                string vTicketType = "01"; //北捷轉乘先都轉到一般卡
                int vSubsidy_Temp = 0;

                switch (vExtName)
                {
                    case ".xls": // 97~2003 舊版格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if ((vRowExcel_H != null) && (vRowExcel_H.Cells[0].ToString().Trim() == "合計") && (vRowExcel_H.Cells.Count > 12))
                            {
                                vSubsidy_Temp = ((vRowExcel_H.Cells[12].CellType == NPOI.SS.UserModel.CellType.Numeric) || (vRowExcel_H.Cells[12].CellType == NPOI.SS.UserModel.CellType.Formula)) ?
                                                (Int32)vRowExcel_H.Cells[12].NumericCellValue : 0;
                                if (vSubsidy_Temp > 0)
                                {
                                    vIndexNo = vCalYM + vDepNo_Temp + vTicketType + Int32.Parse(vTicketLinesNo).ToString("D6");
                                    vSubsidy_MRT = Math.Round(((double)vSubsidy_Temp / 1.05), 0, MidpointRounding.AwayFromZero).ToString();
                                    vSQLStrTemp = "select count(IndexNo) RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount")) == 1)
                                    {
                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                   "   set Subsidy_MRT = Subsidy_MRT + @Subsidy_MRT, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                                   " where IndesNo = @IndexNo";
                                            dsTemp.UpdateParameters.Clear();
                                            dsTemp.UpdateParameters.Add(new Parameter("Subsidy_MRT", DbType.Double, vSubsidy_MRT));
                                            dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                            dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.Update();
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //2007 之後新格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if ((vRowExcel_X != null) && (vRowExcel_X.Cells[0].ToString().Trim() == "合計") && (vRowExcel_X.Cells.Count > 12))
                            {
                                vSubsidy_Temp = ((vRowExcel_X.Cells[12].CellType == NPOI.SS.UserModel.CellType.Numeric) || (vRowExcel_X.Cells[12].CellType == NPOI.SS.UserModel.CellType.Formula)) ?
                                                (Int32)vRowExcel_X.Cells[12].NumericCellValue : 0;
                                if (vSubsidy_Temp > 0)
                                {
                                    vIndexNo = vCalYM + vDepNo_Temp + vTicketType + Int32.Parse(vTicketLinesNo).ToString("D6");
                                    vSubsidy_MRT = Math.Round(((double)vSubsidy_Temp / 1.05), 0, MidpointRounding.AwayFromZero).ToString();
                                    vSQLStrTemp = "select count(IndexNo) RCount from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount")) == 1)
                                    {
                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                   "   set Subsidy_MRT = Subsidy_MRT + @Subsidy_MRT, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                                   " where IndesNo = @IndexNo";
                                            dsTemp.UpdateParameters.Clear();
                                            dsTemp.UpdateParameters.Add(new Parameter("Subsidy_MRT", DbType.Double, vSubsidy_MRT));
                                            dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                            dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.Update();
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
        }

        /// <summary>
        /// 新北市公車票差補貼
        /// </summary>
        private void ImportFromExcel_NewTaipeiPriceDiff()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vDepNo = "";
                string vTempStr = "";
                string vTicketLinesNo = "952"; //新北票差補貼的驗票機代碼固定為 952
                string vTicketType = "01";
                string vSubsidy_NTBus = "";
                int vTempINT = 0;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if ((vRowExcel_H != null) && (vRowExcel_H.Cells[3].ToString() != "") && (vRowExcel_H.Cells[4].ToString().Trim() != ""))
                            {
                                vTempStr = (vRowExcel_H.Cells[3].ToString().Trim() == "桃公站") ? "桃園公車站" :
                                           (vRowExcel_H.Cells[3].ToString().Trim() == "中公站") ? "中壢公車站" :
                                           (vRowExcel_H.Cells[3].ToString().Trim() == "中壢站") ? "中壢站公車" :
                                           vRowExcel_H.Cells[3].ToString().Trim();
                                vSQLStrTemp = "select DepNo from Department where [Name] = '" + vTempStr + "' ";
                                vDepNo = PF.GetValue(vConnStr, vSQLStrTemp, "DepNo").Trim();
                                vSubsidy_NTBus = (Int32.TryParse(vRowExcel_H.Cells[4].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0";
                                if ((vDepNo != "") && (Int32.Parse(vSubsidy_NTBus) > 0))
                                {
                                    vIndexNo = vCalYM + Int32.Parse(vDepNo).ToString("D2") + Int32.Parse(vTicketType).ToString("D2") + Int32.Parse(vTicketLinesNo).ToString("D6");
                                    vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (PF.GetValue(vConnStr, vSQLStrTemp, "IndexNo").Trim() != "")
                                    {
                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                   "   set Subsidy_NTBus = isnull(Subsidy_NTBus, 0) + @Subsidy_NTBus, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                                   " where IndexNo = @IndexNo";
                                            dsTemp.UpdateParameters.Clear();
                                            dsTemp.UpdateParameters.Add(new Parameter("Subsidy_NTBus", DbType.Double, vSubsidy_NTBus));
                                            dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                            dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.Update();
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2007 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if ((vRowExcel_X != null) && (vRowExcel_X.Cells[3].ToString() != "") && (vRowExcel_X.Cells[4].ToString().Trim() != ""))
                            {
                                vTempStr = (vRowExcel_X.Cells[3].ToString().Trim() == "桃公站") ? "桃園公車站" :
                                           (vRowExcel_X.Cells[3].ToString().Trim() == "中公站") ? "中壢公車站" :
                                           (vRowExcel_X.Cells[3].ToString().Trim() == "中壢站") ? "中壢站公車" :
                                           vRowExcel_X.Cells[3].ToString().Trim();
                                vSQLStrTemp = "select DepNo from Department where [Name] = '" + vTempStr + "' ";
                                vDepNo = PF.GetValue(vConnStr, vSQLStrTemp, "DepNo").Trim();
                                vSubsidy_NTBus = (Int32.TryParse(vRowExcel_X.Cells[4].ToString().Trim(), out vTempINT)) ? vTempINT.ToString() : "0";
                                if ((vDepNo != "") && (Int32.Parse(vSubsidy_NTBus) > 0))
                                {
                                    vIndexNo = vCalYM + Int32.Parse(vDepNo).ToString("D2") + Int32.Parse(vTicketType).ToString("D2") + Int32.Parse(vTicketLinesNo).ToString("D6");
                                    vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    if (PF.GetValue(vConnStr, vSQLStrTemp, "IndexNo").Trim() != "")
                                    {
                                        using (SqlDataSource dsTemp = new SqlDataSource())
                                        {
                                            dsTemp.ConnectionString = vConnStr;
                                            dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                   "   set Subsidy_NTBus = isnull(Subsidy_NTBus, 0) + @Subsidy_NTBus, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                                   " where IndexNo = @IndexNo";
                                            dsTemp.UpdateParameters.Clear();
                                            dsTemp.UpdateParameters.Add(new Parameter("Subsidy_NTBus", DbType.Double, vSubsidy_NTBus));
                                            dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                            dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                            dsTemp.Update();
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
        }

        /// <summary>
        /// 1280 暫撥款
        /// </summary>
        private void ImportFromExcel_TPMPass()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vDepNo = "";
                string vTempStr = "";
                string vTicketLinesNo = "952"; //雙北定期票補貼的驗票機代碼固定為 952
                string vTicketType = "01";
                string vSubsidy_TPMPass = "";
                double vTempFloat = 0.0;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if ((vRowExcel_H != null) &&
                                (vRowExcel_H.Cells[3].ToString().Trim() != "") &&
                                (vRowExcel_H.Cells[5].ToString().Trim() != ""))
                            {
                                vTempStr = (vRowExcel_H.Cells[3].ToString().Trim() == "桃公站") ? "桃園公車站" :
                                           (vRowExcel_H.Cells[3].ToString().Trim() == "中公站") ? "中壢公車站" :
                                           (vRowExcel_H.Cells[3].ToString().Trim() == "中壢站") ? "中壢站公車" :
                                           vRowExcel_H.Cells[3].ToString().Trim();
                                vSQLStrTemp = "select DepNo from Department where [Name] = '" + vTempStr + "' ";
                                vDepNo = PF.GetValue(vConnStr, vSQLStrTemp, "DepNo").Trim();
                                vSubsidy_TPMPass = ((vRowExcel_H.Cells[5].CellType == CellType.Numeric) && (vRowExcel_H.Cells[5].ToString().Trim() != "")) ?
                                                   vRowExcel_H.Cells[5].NumericCellValue.ToString() :
                                                   (Double.TryParse(vRowExcel_H.Cells[5].ToString().Trim(), out vTempFloat)) ?
                                                   vTempFloat.ToString() : "0.0";
                                if ((vDepNo != "") && (Int32.Parse(vSubsidy_TPMPass) > 0))
                                {
                                    vIndexNo = vCalYM + Int32.Parse(vDepNo).ToString("D2") + Int32.Parse(vTicketType).ToString("D2") + Int32.Parse(vTicketLinesNo).ToString("D6");
                                    vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    using (SqlDataSource dsTemp = new SqlDataSource())
                                    {
                                        dsTemp.ConnectionString = vConnStr;
                                        dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                               "   set Subsidy_TPMPass = isnull(Subsidy_TPMPass, 0) + @Subsidy_TPMPass, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                               " where IndexNo = @IndexNo";
                                        dsTemp.UpdateParameters.Clear();
                                        dsTemp.UpdateParameters.Add(new Parameter("Subsidy_TPMPass", DbType.Double, vSubsidy_TPMPass));
                                        dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                        dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                        dsTemp.Update();
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2007 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if ((vRowExcel_X != null) &&
                                (vRowExcel_X.Cells[3].ToString().Trim() != "") &&
                                (vRowExcel_X.Cells[5].ToString().Trim() != ""))
                            {
                                vTempStr = (vRowExcel_X.Cells[3].ToString().Trim() == "桃公站") ? "桃園公車站" :
                                           (vRowExcel_X.Cells[3].ToString().Trim() == "中公站") ? "中壢公車站" :
                                           (vRowExcel_X.Cells[3].ToString().Trim() == "中壢站") ? "中壢站公車" :
                                           vRowExcel_X.Cells[3].ToString().Trim();
                                vSQLStrTemp = "select DepNo from Department where [Name] = '" + vTempStr + "' ";
                                vDepNo = PF.GetValue(vConnStr, vSQLStrTemp, "DepNo").Trim();
                                vSubsidy_TPMPass = ((vRowExcel_X.Cells[5].CellType == CellType.Numeric) && (vRowExcel_X.Cells[5].ToString().Trim() != "")) ?
                                                   vRowExcel_X.Cells[5].NumericCellValue.ToString() :
                                                   (Double.TryParse(vRowExcel_X.Cells[5].ToString().Trim(), out vTempFloat)) ?
                                                   vTempFloat.ToString() : "0.0";
                                if ((vDepNo != "") && (Int32.Parse(vSubsidy_TPMPass) > 0))
                                {
                                    vIndexNo = vCalYM + Int32.Parse(vDepNo).ToString("D2") + Int32.Parse(vTicketType).ToString("D2") + Int32.Parse(vTicketLinesNo).ToString("D6");
                                    vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                    using (SqlDataSource dsTemp = new SqlDataSource())
                                    {
                                        dsTemp.ConnectionString = vConnStr;
                                        dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                               "   set Subsidy_TPMPass = isnull(Subsidy_TPMPass, 0) + @Subsidy_TPMPass, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                               " where IndexNo = @IndexNo";
                                        dsTemp.UpdateParameters.Clear();
                                        dsTemp.UpdateParameters.Add(new Parameter("Subsidy_TPMPass", DbType.Double, vSubsidy_TPMPass));
                                        dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                        dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                        dsTemp.Update();
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
        }

        /// <summary>
        /// 北捷定期票預撥款
        /// </summary>
        private void ImportFromExcel_TPMPass2()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vDepNo = "";
                string vTempStr = "";
                string vTicketLinesNo = "952"; //雙北定期票補貼的驗票機代碼固定為 952
                string vTicketType = "01";
                string vSubsidy_TPMPass = "";
                double vTempFloat = 0.0;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        break;

                    case ".xlsx": //新版 EXCEL (2007 以後) 格式
                        break;

                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;
                }
            }
        }

        /// <summary>
        /// 新北老人補貼
        /// </summary>
        private void ImportFromExcel_NTOld()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vDepNo = "";
                string vTempStr = "";
                string vTicketLinesNo = "";
                string vERPLinesNo = "";
                string vGOVLinesNo = "";
                string vTicketType = "";
                string vSubsidy_NTOld = "";
                double vTempFloat = 0.0;
                string[] vaTempArray;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        for (int sheetCount = 0; sheetCount < wbExcel_H.NumberOfSheets; sheetCount++)
                        {
                            HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(sheetCount);
                            if ((sheetExcel_H.SheetName != "日樞") &&
                                (sheetExcel_H.SheetName != "非樞") &&
                                (sheetExcel_H.SheetName != "新車表"))
                            {
                                for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                                {
                                    HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                    if ((vRowExcel_H != null) &&
                                        (vRowExcel_H.Cells.Count >= 11))
                                    {
                                        vTempStr = vRowExcel_H.Cells[0].StringCellValue.Trim();
                                        switch (vTempStr)
                                        {
                                            default:
                                                vTicketType = "";
                                                break;

                                            case "愛心票":
                                                vTicketType = "03";
                                                break;

                                            case "敬老票":
                                                vTicketType = "04";
                                                break;
                                        }
                                        if (vTicketType != "")
                                        {
                                            vTicketLinesNo = vRowExcel_H.Cells[2].ToString().Trim().Substring(2);
                                            vSQLStrTemp = "select (ERPLinesNo + ',' + GOVLinesNo + ',' + DepNo) LineData " + Environment.NewLine +
                                                          "  from LinesNoChart " + Environment.NewLine +
                                                          " where TicketLinesNo = '" + vTicketLinesNo + "' ";
                                            vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                            if (vTempStr != "")
                                            {
                                                vaTempArray = vTempStr.Split(',');
                                                vERPLinesNo = (vaTempArray.Length > 0) ? vaTempArray[0].Trim() : "";
                                                vGOVLinesNo = (vaTempArray.Length > 1) ? vaTempArray[1].Trim() : "";
                                                vDepNo = (vaTempArray.Length > 2) ? vaTempArray[2].Trim() : "";
                                            }
                                            if (double.TryParse(vRowExcel_H.Cells[11].ToString().Trim(), out vTempFloat))
                                            {
                                                vSubsidy_NTOld = vTempFloat.ToString();
                                            }
                                            else
                                            {
                                                vSubsidy_NTOld = "0.0";
                                            }
                                            if ((vDepNo != "") &&
                                                (vTicketLinesNo != "") &&
                                                (vTicketType != "") &&
                                                (Double.Parse(vSubsidy_NTOld) > 0.0))
                                            {
                                                vIndexNo = vCalYM + Int32.Parse(vDepNo).ToString("D2") + Int32.Parse(vTicketType).ToString("D2") + Int32.Parse(vTicketLinesNo).ToString("D6");
                                                vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                                using (SqlDataSource dsTemp = new SqlDataSource())
                                                {
                                                    dsTemp.ConnectionString = vConnStr;
                                                    dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                           "   set Subsidy_NTOld = isnull(Subsidy_NTOld, 0) + @Subsidy_NTOld, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                                           " where IndexNo = @IndexNo";
                                                    dsTemp.UpdateParameters.Clear();
                                                    dsTemp.UpdateParameters.Add(new Parameter("Subsidy_NTOld", DbType.Double, vSubsidy_NTOld));
                                                    dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                    dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                    dsTemp.Update();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2007 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        for (int sheetCount = 0; sheetCount < wbExcel_X.NumberOfSheets; sheetCount++)
                        {
                            XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(sheetCount);
                            if ((sheetExcel_X.SheetName != "日樞") &&
                                (sheetExcel_X.SheetName != "非樞") &&
                                (sheetExcel_X.SheetName != "新車表"))
                            {
                                for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                                {
                                    XSSFRow vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                    if ((vRowExcel_X != null) &&
                                        (vRowExcel_X.Cells.Count >= 11))
                                    {
                                        vTempStr = vRowExcel_X.Cells[0].StringCellValue.Trim();
                                        switch (vTempStr)
                                        {
                                            default:
                                                vTicketType = "";
                                                break;

                                            case "愛心票":
                                                vTicketType = "03";
                                                break;

                                            case "敬老票":
                                                vTicketType = "04";
                                                break;
                                        }
                                        if (vTicketType != "")
                                        {
                                            vTicketLinesNo = vRowExcel_X.Cells[2].ToString().Trim().Substring(2);
                                            vSQLStrTemp = "select (ERPLinesNo + ',' + GOVLinesNo + ',' + DepNo) LineData " + Environment.NewLine +
                                                          "  from LinesNoChart " + Environment.NewLine +
                                                          " where TicketLinesNo = '" + vTicketLinesNo + "' ";
                                            vTempStr = PF.GetValue(vConnStr, vSQLStrTemp, "LineData");
                                            if (vTempStr != "")
                                            {
                                                vaTempArray = vTempStr.Split(',');
                                                vERPLinesNo = (vaTempArray.Length > 0) ? vaTempArray[0].Trim() : "";
                                                vGOVLinesNo = (vaTempArray.Length > 1) ? vaTempArray[1].Trim() : "";
                                                vDepNo = (vaTempArray.Length > 2) ? vaTempArray[2].Trim() : "";
                                            }
                                            if (double.TryParse(vRowExcel_X.Cells[11].ToString().Trim(), out vTempFloat))
                                            {
                                                vSubsidy_NTOld = vTempFloat.ToString();
                                            }
                                            else
                                            {
                                                vSubsidy_NTOld = "0.0";
                                            }
                                            if ((vDepNo != "") &&
                                                (vTicketLinesNo != "") &&
                                                (vTicketType != "") &&
                                                (Double.Parse(vSubsidy_NTOld) > 0.0))
                                            {
                                                vIndexNo = vCalYM + Int32.Parse(vDepNo).ToString("D2") + Int32.Parse(vTicketType).ToString("D2") + Int32.Parse(vTicketLinesNo).ToString("D6");
                                                vSQLStrTemp = "select IndexNo from LinesBounds where IndexNo = '" + vIndexNo + "' ";
                                                using (SqlDataSource dsTemp = new SqlDataSource())
                                                {
                                                    dsTemp.ConnectionString = vConnStr;
                                                    dsTemp.UpdateCommand = "update LinesBounds " + Environment.NewLine +
                                                                           "   set Subsidy_NTOld = isnull(Subsidy_NTOld, 0) + @Subsidy_NTOld, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                                           " where IndexNo = @IndexNo";
                                                    dsTemp.UpdateParameters.Clear();
                                                    dsTemp.UpdateParameters.Add(new Parameter("Subsidy_NTOld", DbType.Double, vSubsidy_NTOld));
                                                    dsTemp.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                                                    dsTemp.UpdateParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo));
                                                    dsTemp.Update();
                                                }
                                            }
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
        }

        /// <summary>
        /// 台灣好行
        /// </summary>
        private void ImportFromExcel_TWTraip()
        {
            if (fuExcel.FileName == "")
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇來源檔案並指定要轉入的目標')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
                string vExtName = Path.GetExtension(fuExcel.FileName); //取回副檔名
                string vSQLStrTemp = "";

                string vIndexNo = "";
                string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
                string vDepNo = "";
                string vTempStr = "";
                string vTicketLinesNo = "";
                string vERPLinesNo = "";
                string vGOVLinesNo = "";
                string vTicketType = "";
                string vSubsidy_Trip501 = "";
                string vSubsidy_Trip502 = "";
                double vTempFloat = 0.0;
                string[] vaTempArray;

                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式

                        break;

                    case ".xlsx": //新版 EXCEL (2007 以後) 格式

                        break;
                }
            }
        }

        protected void bbImportData_Click(object sender, EventArgs e)
        {
            switch (ddlSourceTyoe.SelectedValue)
            {
                case "01": //各路線現金收入
                    OpenData();
                    break;
                case "02": //各路線刷卡收入
                    ImportFromExcel_CardIn();
                    OpenData();
                    break;
                case "03": //表一資料
                    ImportFromExcel_List1();
                    OpenData();
                    break;
                case "04": //表三資料
                    //ImportFromExcel_List1(); 
                    ImportFromExcel_List3();
                    OpenData();
                    break;
                case "05": //學生票資料
                    ImportFromExcel_Student();
                    OpenData();
                    break;
                case "06": //敬老 / 愛心卡資料
                    ImportFromExcel_Older();
                    OpenData();
                    break;
                case "07": //國道 / 台鐵轉乘
                case "08":
                    ImportFromExcel_Highway();
                    OpenData();
                    break;
                case "09": //國小升國中補貼
                    ImportFromExcel_NewHighschool();
                    OpenData();
                    break;
                case "10": //新北公車間轉乘補貼
                    ImportFromExcel_NewTaipeiBusTrans();
                    OpenData();
                    break;
                case "11": //新北市公車票差補貼
                    ImportFromExcel_NewTaipeiPriceDiff();
                    OpenData();
                    break;
                case "12": //國道準運價
                    ImportFromExcel_Highway2();
                    OpenData();
                    break;
                case "13": //北捷轉乘補貼
                    ImportFromExcel_MRT();
                    OpenData();
                    break;
                case "14": //1280 暫撥款
                    ImportFromExcel_TPMPass();
                    OpenData();
                    break;
                case "15": //北捷定期票預撥款
                    ImportFromExcel_TPMPass2();
                    OpenData();
                    break;
                case "16": //新北老人補貼
                    ImportFromExcel_NTOld();
                    OpenData();
                    break;
                case "17": //台灣好行
                    ImportFromExcel_TWTraip();
                    OpenData();
                    break;

                default:
                    break;
            }
        }

        protected void bbSearchData_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        private void OpenData()
        {
            plShowData.Visible = false;
            string vCalYM = eImportYear.Text.Trim() + ddlImportMonth.SelectedValue.Trim();
            dsImportData.SelectCommand = "SELECT IndexNo, CalYM, TicketType, " + Environment.NewLine +
                                         "       (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '路線優惠補助    fmLinesBounds   TicketType') AND (CLASSNO = a.TicketType)) AS TicketType_C, " + Environment.NewLine +
                                         "       TicketLineNo, LinesGovNo, TicketPCount, TotalAmount, RealIncome, Subsidy_8KM, Subsidy_Limited, Subsidy_Student, Subsidy_Social, Subsidy_MRT, " + Environment.NewLine +
                                         "       Subsidy_Area, Subsidy_Highway, Subsidy_Others, Subsidy_Dift, BuMan, BuDate, ModifyMan, ModifyDate " + Environment.NewLine +
                                         "  FROM LinesBounds AS a " + Environment.NewLine +
                                         " WHERE (ISNULL(IndexNo, '') <> '') " + Environment.NewLine +
                                         "   AND a.CalYM = '" + vCalYM + "' " + Environment.NewLine +
                                         " ORDER BY a.TicketLineNo, a.TicketType ";
            gridShowData.DataBind();
            plShowData.Visible = true;
        }

        /// <summary>
        /// 匯出 EXCEL 檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExportExcel_Click(object sender, EventArgs e)
        {
            string vTempStr = "";
            double vTempFloat = 0.0;
            if (gridShowData.Rows.Count > 0)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCalYM = eImportYear.Text.Trim() + Int32.Parse(ddlImportMonth.SelectedValue).ToString("D2");
                string vCalYM_Last = (Int32.Parse(eImportYear.Text.Trim()) - 1).ToString("D3") + Int32.Parse(ddlImportMonth.SelectedValue).ToString("D2");
                DateTime vDate_S = (Int32.Parse(eImportYear.Text.Trim()) < 1911) ?
                                   new DateTime(Int32.Parse(eImportYear.Text.Trim()) + 1911, Int32.Parse(ddlImportMonth.SelectedValue.Trim()), 1) :
                                   new DateTime(Int32.Parse(eImportYear.Text.Trim()), Int32.Parse(ddlImportMonth.SelectedValue.Trim()), 1);
                DateTime vDate_E = vDate_S.AddMonths(1).AddDays(-1);
                string vSQLStr_Temp = "select la.TicketLineNo, la.LinesGovNo, lb.LineName, (select LicenseUnit from LicenseA where LicenseNo = lb.LicenseNo) LicenseUnit, " + Environment.NewLine +
                                      "       case when lb.IsExtra = 'V' then '是' else '否' end as IsExtra, (select EndDate from LicenseA where LicenseNo = lb.LicenseNo) EndDate, " + Environment.NewLine +
                                      "       la.DepNo, (select[Name] from Department where DepNo = la.DepNo) as MainDepName, la.CarCount, lb.LicenseRun, lb.LicenseRunSun, " + Environment.NewLine +
                                      "       la.DailyCostByCar, la.DailyIncomeByCar, sum(isnull(la.TotalAmount, 0)) as TotalAmount, cast(null as Float) as TotalCost, " + Environment.NewLine +
                                      "       cast(null as float) as MonthProfit, cast(null as float) as IncomePerKM, cast(null as float) as DistanceCostByCar, " + Environment.NewLine +
                                      "       cast(null as float) as ProfitByDistance, cast(null as float) as LastYMIncomeByKM, cast(null as float) as DiffIncomeByKM, sum(la.TotalKM) as TotalKM, " + Environment.NewLine +
                                      "       sum(isnull(la.RealIncome, 0)) as TotalRealIncome, sum(isnull(la.Subsidy_Dift, 0)) as Subsidy_Dift, sum(isnull(la.Subsidy_Others, 0)) as Subsidy_Others, " + Environment.NewLine +
                                      "       sum(isnull(la.TicketIncome, 0)) as TicketIncome, sum(isnull(la.CashIncome, 0)) as CashIncome, sum(isnull(la.DeductAmount, 0)) as DeductAmount, " + Environment.NewLine +
                                      "       sum(isnull(la.Subsidy_8KM, 0)) as Subsidy_8KM, sum(isnull(la.Subsidy_Student, 0)) as Subsidy_Student, " + Environment.NewLine +
                                      "       sum(isnull(la.Subsidy_Social, 0)) as Subsidy_Social, sum(isnull(la.Subsidy_Ticket_D, 0)) as Subsidy_TicketDift, " + Environment.NewLine +
                                      "       sum(isnull(la.Subsidy_TWTrip, 0)) as Subsidy_TWTrip, sum(isnull(la.Subsidy_NTOld, 0)) as Subsidy_NTOld, " + Environment.NewLine +
                                      "       sum(isnull(la.Subsidy_Highway, 0)) as Subsidy_Highway, sum(isnull(la.Subsidy_TPMPass, 0)) as Subsidy_TPMPass, " + Environment.NewLine +
                                      "       sum(isnull(la.Subsidy_NTBus, 0)) as Subsidy_NTBus, sum(isnull(la.Subsidy_MRT, 0)) as Subsidy_MRT, " + Environment.NewLine +
                                      "       sum(Isnull(la.Subsidy_Student, 0)) as Subsidy_Student2, sum(isnull(la.Subsidy_BusTran, 0)) as Subsidy_BusTran, " + Environment.NewLine +
                                      "       sum(isnull(la.Subsidy_Trial, 0)) as Subsidy_Trial, sum(isnull(la.Subsidy_Trip501, 0)) as Subsidy_Trip501, " + Environment.NewLine +
                                      "       sum(isnull(la.Subsidy_Trip502, 0)) as Subsidy_Trip502, sum(isnull(la.Subsidy_TrialPre, 0)) as Subsidy_TrialPre, " + Environment.NewLine +
                                      "       sum(isnull(la.Subsidy_Highway2, 0)) as Subsidy_Highway2, sum(isnull(la.Subsidy_Old2, 0)) as Subsidy_Old2  " + Environment.NewLine +
                                      "  from LinesBounds as la left join Lines as lb on cast(lb.TicketLineNo as INT) = cast(la.TicketLineNo as INT) " + Environment.NewLine +
                                      " where la.CalYM = '" + vCalYM + "' " + Environment.NewLine +
                                      " group by la.TicketLineNo, lb.LinesNo, lb.LineName, lb.MainGovOffice, lb.IsExtra, la.CarCount, la.LinesGovNo, " + Environment.NewLine +
                                      "          lb.LicenseRun, lb.LicenseRunSun, lb.LinesNo, la.DepNo, la.DailyCostByCar, la.DailyIncomeByCar, lb.LicenseNo " + Environment.NewLine +
                                      " order by la.TicketLineNo, la.DepNo";
                using (SqlConnection connExcel = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdExcel = new SqlCommand(vSQLStr_Temp, connExcel);
                    connExcel.Open();
                    SqlDataReader drExcel = cmdExcel.ExecuteReader();
                    if (drExcel.HasRows)
                    {
                        //有資料才做下面的動作
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

                        string vFileName = "各站各路線 " + eImportYear.Text.Trim() + " 年 " + ddlImportMonth.SelectedValue.Trim() + " 月營收資料";
                        string vHeaderText = "";
                        int vLinesNo = 0;
                        DateTime vBuDate;
                        //新增一個工作表
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                        //寫入標題列
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            vHeaderText = (drExcel.GetName(i).ToUpper() == "TICKETLINENO") ? "驗票機路線代碼" :
                                          (drExcel.GetName(i).ToUpper() == "DEPNO") ? "站別" :
                                          (drExcel.GetName(i).ToUpper() == "LINESNO") ? "ERP路線代碼" :
                                          (drExcel.GetName(i).ToUpper() == "LINESGOVNO") ? "路線政府編碼" :
                                          (drExcel.GetName(i).ToUpper() == "LINENAME") ? "路線中文說明" :
                                          (drExcel.GetName(i).ToUpper() == "MAINGOVOFFICE") ? "主管機關" :
                                          (drExcel.GetName(i).ToUpper() == "ISEXTRA") ? "是否補貼路線" :
                                          (drExcel.GetName(i).ToUpper() == "ENDDATE") ? "許可證有效日期" :
                                          (drExcel.GetName(i).ToUpper() == "MAINDEPNAME") ? "主責站別" :
                                          (drExcel.GetName(i).ToUpper() == "CARCOUNT") ? "車輛數" :
                                          (drExcel.GetName(i).ToUpper() == "LICENSERUN") ? "核定班次數" :
                                          (drExcel.GetName(i).ToUpper() == "LICENSERUNSUN") ? "假日核定班次數" :
                                          (drExcel.GetName(i).ToUpper() == "DAILYCOSTBYCAR") ? "平均每日每車成本" :
                                          (drExcel.GetName(i).ToUpper() == "DAILYINCOMEBYCAR") ? "平均每日每車收入" :
                                          (drExcel.GetName(i).ToUpper() == "TOTALKM") ? "總營運里程" :
                                          (drExcel.GetName(i).ToUpper() == "TOTALAMOUNT") ? "總營收" :
                                          (drExcel.GetName(i).ToUpper() == "TOTALCOST") ? "路線總成本" :
                                          (drExcel.GetName(i).ToUpper() == "MONTHPROFIT") ? "月盈虧金額" :
                                          (drExcel.GetName(i).ToUpper() == "INCOMEPERKM") ? "車公里營收" :
                                          (drExcel.GetName(i).ToUpper() == "DISTANCECOSTBYCAR") ? "車公里成本" :
                                          (drExcel.GetName(i).ToUpper() == "PROFITBYDISTANCE") ? "公里盈虧" :
                                          (drExcel.GetName(i).ToUpper() == "LASTYMINCOMEBYKM") ? "去年同期收入" :
                                          (drExcel.GetName(i).ToUpper() == "DIFFINCOMEBYKM") ? "車公里收入差異" :
                                          (drExcel.GetName(i).ToUpper() == "ERPTOTALKM") ? "ERP路線里程" :
                                          (drExcel.GetName(i).ToUpper() == "TOTALREALINCOME") ? "客月票收入" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_DIFT") ? "營運虧損補貼" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_OTHERS") ? "其他收入/補貼" :
                                          (drExcel.GetName(i).ToUpper() == "TICKETINCOME") ? "票證收入" :
                                          (drExcel.GetName(i).ToUpper() == "CASHINCOME") ? "現金收入" :
                                          (drExcel.GetName(i).ToUpper() == "DEDUCTAMOUNT") ? "拆分扣除金額" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_8KM") ? "8公里優惠" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_STUDENT") ? "學生卡補貼" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_SOCIAL") ? "老人優待補貼" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_TICKET_D") ? "公路票差補貼" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_TWTRIP") ? "台灣好行補貼" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_NTOLD") ? "新北老人補貼" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_HIGHWAY") ? "國道轉乘" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_TPMPASS") ? "雙北定期票" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_NTBUS") ? "新北市票差補貼" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_MRT") ? "北捷轉乘" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_STUDENT2") ? "國小升國中" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_BUSTRAN") ? "新北市公車間轉乘" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_TRIAL") ? "試辦路線補貼" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_TRIP501") ? "台灣好行501、503" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_TRIP502") ? "台灣好行502" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_TRIALPRE") ? "試辦路線補貼預估" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_HIGHWAY2") ? "國道準運價" :
                                          (drExcel.GetName(i).ToUpper() == "LICENSEUNIT") ? "主管機關" :
                                          (drExcel.GetName(i).ToUpper() == "SUBSIDY_OLD2") ? "老人半票補貼" :
                                          drExcel.GetName(i).Trim();
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
                                vHeaderText = drExcel.GetName(i).ToUpper();
                                if (drExcel[i].ToString() != "")
                                {
                                    if ((vHeaderText == "TICKETLINENO") ||
                                        (vHeaderText == "DEPNO") ||
                                        (vHeaderText == "LINESNO") ||
                                        (vHeaderText == "LINENAME") ||
                                        (vHeaderText == "ISEXTRA") ||
                                        (vHeaderText == "LINESGOVNO") ||
                                        (vHeaderText == "MAINGOVOFFICE") ||
                                        (vHeaderText == "LICENSEUNIT") ||
                                        (vHeaderText == "MAINDEPNAME"))
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                    }
                                    else if (vHeaderText == "ENDDATE")
                                    {
                                        vTempStr = drExcel[i].ToString();
                                        vBuDate = DateTime.Parse(drExcel[i].ToString());
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue((vBuDate.Year - 1911).ToString() + "/" + vBuDate.ToString("MM/dd"));
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                    }
                                    else if ((vHeaderText == "CARCOUNT") ||
                                             (vHeaderText == "LICENSERUN") ||
                                             (vHeaderText == "LICENSERUNSUN"))
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                    }
                                    else
                                    {
                                        vTempStr = drExcel[i].ToString();
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.TryParse(vTempStr, out vTempFloat) ? vTempFloat : 0.0);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Float;
                                    }
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(String.Empty);
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
                                                     "LinesBounds.aspx" + Environment.NewLine +
                                                     "計算年月" + vCalYM;
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
        }
    }
}