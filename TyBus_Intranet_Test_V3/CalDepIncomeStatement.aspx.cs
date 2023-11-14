using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class CalDepIncomeStatement : System.Web.UI.Page
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

                DateTime vCalDate = DateTime.Today.AddMonths(-1);

                if (vLoginID != "")
                {
                    if (!IsPostBack)
                    {
                        if (vConnStr == "")
                        {
                            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                        }
                        eCalYear.Text = (vCalDate.Year - 1911).ToString();
                        eCalMonth.Text = vCalDate.Month.ToString();
                        eBudgetYear.Text = (vCalDate.Year - 1911).ToString();
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

        protected void bbImportBudget_Click(object sender, EventArgs e)
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            string vUploadFileName = "";
            string vExtName = Path.GetExtension(eFilePath.FileName);
            vUploadFileName = vUploadPath + eFilePath.FileName;
            eFilePath.SaveAs(vUploadFileName);

            string vTempStr = "";
            string vBudgetIndex = "";
            string vDepNo_Temp = "";
            string vBudgetYear = eBudgetYear.Text.Trim();
            string vChartBarCode = "";
            string vMonth01 = "0";
            string vMonth02 = "0";
            string vMonth03 = "0";
            string vMonth04 = "0";
            string vMonth05 = "0";
            string vMonth06 = "0";
            string vMonth07 = "0";
            string vMonth08 = "0";
            string vMonth09 = "0";
            string vMonth10 = "0";
            string vMonth11 = "0";
            string vMonth12 = "0";
            string vYearTotal = "0";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vRecordNote = "匯入資料_客運業務損益表" + Environment.NewLine +
                                 "CalDepIncomeStatement.aspx";
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            //=============================================================================================
            if (vExtName == ".xls")
            {
                vTempStr = "delete DepMonthBudget where BudgetYear = '" + vBudgetYear + "' ";
                PF.ExecSQL(vConnStr, vTempStr);
                //來源檔案是 97-2003 版的 EXCEL
                HSSFWorkbook wbExcel_H = new HSSFWorkbook(eFilePath.FileContent);
                HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                {
                    HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);

                    if (vRowTemp_H.Cells.Count > 0)
                    {
                        vDepNo_Temp = (vRowTemp_H.GetCell((Int32)vRowTemp_H.FirstCellNum).CellType != CellType.Blank) ? vRowTemp_H.GetCell((Int32)vRowTemp_H.FirstCellNum).StringCellValue : "";
                        vChartBarCode = (vRowTemp_H.GetCell(2).CellType != CellType.Blank) ? vRowTemp_H.GetCell(2).StringCellValue : "";
                        if ((vDepNo_Temp != "") && (vChartBarCode != ""))
                        {
                            vBudgetIndex = vBudgetYear + vDepNo_Temp + vChartBarCode;
                            vMonth01 = (vRowTemp_H.GetCell(4).CellType != CellType.Blank) ? vRowTemp_H.GetCell(4).NumericCellValue.ToString() : "0";
                            vMonth02 = (vRowTemp_H.GetCell(5).CellType != CellType.Blank) ? vRowTemp_H.GetCell(5).NumericCellValue.ToString() : "0";
                            vMonth03 = (vRowTemp_H.GetCell(6).CellType != CellType.Blank) ? vRowTemp_H.GetCell(6).NumericCellValue.ToString() : "0";
                            vMonth04 = (vRowTemp_H.GetCell(7).CellType != CellType.Blank) ? vRowTemp_H.GetCell(7).NumericCellValue.ToString() : "0";
                            vMonth05 = (vRowTemp_H.GetCell(8).CellType != CellType.Blank) ? vRowTemp_H.GetCell(8).NumericCellValue.ToString() : "0";
                            vMonth06 = (vRowTemp_H.GetCell(9).CellType != CellType.Blank) ? vRowTemp_H.GetCell(9).NumericCellValue.ToString() : "0";
                            vMonth07 = (vRowTemp_H.GetCell(10).CellType != CellType.Blank) ? vRowTemp_H.GetCell(10).NumericCellValue.ToString() : "0";
                            vMonth08 = (vRowTemp_H.GetCell(11).CellType != CellType.Blank) ? vRowTemp_H.GetCell(11).NumericCellValue.ToString() : "0";
                            vMonth09 = (vRowTemp_H.GetCell(12).CellType != CellType.Blank) ? vRowTemp_H.GetCell(12).NumericCellValue.ToString() : "0";
                            vMonth10 = (vRowTemp_H.GetCell(13).CellType != CellType.Blank) ? vRowTemp_H.GetCell(13).NumericCellValue.ToString() : "0";
                            vMonth11 = (vRowTemp_H.GetCell(14).CellType != CellType.Blank) ? vRowTemp_H.GetCell(14).NumericCellValue.ToString() : "0";
                            vMonth12 = (vRowTemp_H.GetCell(15).CellType != CellType.Blank) ? vRowTemp_H.GetCell(15).NumericCellValue.ToString() : "0";
                            vYearTotal = (vRowTemp_H.GetCell(16).CellType != CellType.Blank) ? vRowTemp_H.GetCell(16).NumericCellValue.ToString() : "0";

                            vTempStr = "INSERT INTO DepMonthBudget " + Environment.NewLine +
                                       "            (BudgetIndex, DepNo, BudgetYear, ChartBarCode, Month01, Month02, Month03, Month04, Month05, Month06, Month07, Month08, Month09, Month10, Month11, Month12, YearTotal)" + Environment.NewLine +
                                       "     VALUES ('" + vBudgetIndex + "', '" + vDepNo_Temp + "', '" + vBudgetYear + "', '" + vChartBarCode + "', " + vMonth01 + ", " + vMonth02 + ", " + vMonth03 + ", " + vMonth04 + ", " + vMonth05 + ", " + vMonth06 + ", " + vMonth07 + ", " + vMonth08 + ", " + vMonth09 + ", " + vMonth10 + ", " + vMonth11 + ", " + vMonth12 + ", " + vYearTotal + ")";
                            PF.ExecSQL(vConnStr, vTempStr);
                        }
                    }
                }
            }
            else if (vExtName == ".xlsx")
            {
                vTempStr = "delete DepMonthBudget where BudgetYear = '" + vBudgetYear + "' ";
                PF.ExecSQL(vConnStr, vTempStr);
                //來源檔案是 2010 或之後版本的 EXCEL
                XSSFWorkbook wbExcel_X = new XSSFWorkbook(eFilePath.FileContent);
                XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                {
                    XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);

                    if (vRowTemp_X.Cells.Count > 0)
                    {
                        vDepNo_Temp = (vRowTemp_X.GetCell((Int32)vRowTemp_X.FirstCellNum).CellType != CellType.Blank) ? vRowTemp_X.GetCell((Int32)vRowTemp_X.FirstCellNum).ToString() : "";
                        vChartBarCode = (vRowTemp_X.GetCell(2).CellType != CellType.Blank) ? vRowTemp_X.GetCell(2).StringCellValue : "";
                        if ((vDepNo_Temp != "") && (vChartBarCode != ""))
                        {
                            vBudgetIndex = vBudgetYear + vDepNo_Temp + vChartBarCode;

                            vMonth01 = (vRowTemp_X.GetCell(4).CellType != CellType.Blank) ? vRowTemp_X.GetCell(4).NumericCellValue.ToString() : "0";
                            vMonth02 = (vRowTemp_X.GetCell(5).CellType != CellType.Blank) ? vRowTemp_X.GetCell(5).NumericCellValue.ToString() : "0";
                            vMonth03 = (vRowTemp_X.GetCell(6).CellType != CellType.Blank) ? vRowTemp_X.GetCell(6).NumericCellValue.ToString() : "0";
                            vMonth04 = (vRowTemp_X.GetCell(7).CellType != CellType.Blank) ? vRowTemp_X.GetCell(7).NumericCellValue.ToString() : "0";
                            vMonth05 = (vRowTemp_X.GetCell(8).CellType != CellType.Blank) ? vRowTemp_X.GetCell(8).NumericCellValue.ToString() : "0";
                            vMonth06 = (vRowTemp_X.GetCell(9).CellType != CellType.Blank) ? vRowTemp_X.GetCell(9).NumericCellValue.ToString() : "0";
                            vMonth07 = (vRowTemp_X.GetCell(10).CellType != CellType.Blank) ? vRowTemp_X.GetCell(10).NumericCellValue.ToString() : "0";
                            vMonth08 = (vRowTemp_X.GetCell(11).CellType != CellType.Blank) ? vRowTemp_X.GetCell(11).NumericCellValue.ToString() : "0";
                            vMonth09 = (vRowTemp_X.GetCell(12).CellType != CellType.Blank) ? vRowTemp_X.GetCell(12).NumericCellValue.ToString() : "0";
                            vMonth10 = (vRowTemp_X.GetCell(13).CellType != CellType.Blank) ? vRowTemp_X.GetCell(13).NumericCellValue.ToString() : "0";
                            vMonth11 = (vRowTemp_X.GetCell(14).CellType != CellType.Blank) ? vRowTemp_X.GetCell(14).NumericCellValue.ToString() : "0";
                            vMonth12 = (vRowTemp_X.GetCell(15).CellType != CellType.Blank) ? vRowTemp_X.GetCell(15).NumericCellValue.ToString() : "0";
                            vYearTotal = (vRowTemp_X.GetCell(16).CellType != CellType.Blank) ? vRowTemp_X.GetCell(16).NumericCellValue.ToString() : "0";

                            vTempStr = "INSERT INTO DepMonthBudget " + Environment.NewLine +
                                       "            (BudgetIndex, DepNo, BudgetYear, ChartBarCode, Month01, Month02, Month03, Month04, Month05, Month06, Month07, Month08, Month09, Month10, Month11, Month12, YearTotal)" + Environment.NewLine +
                                       "     VALUES ('" + vBudgetIndex + "', '" + vDepNo_Temp + "', '" + vBudgetYear + "', '" + vChartBarCode + "', " + vMonth01 + ", " + vMonth02 + ", " + vMonth03 + ", " + vMonth04 + ", " + vMonth05 + ", " + vMonth06 + ", " + vMonth07 + ", " + vMonth08 + ", " + vMonth09 + ", " + vMonth10 + ", " + vMonth11 + ", " + vMonth12 + ", " + vYearTotal + ")";

                            PF.ExecSQL(vConnStr, vTempStr);
                        }
                    }
                }
            }
            Response.Write("<Script language='Javascript'>");
            Response.Write("alert('資料轉入完畢！')");
            Response.Write("</" + "Script>");
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCalYear = (Int32.Parse(eCalYear.Text.Trim()) > 1911) ? (Int32.Parse(eCalYear.Text.Trim()) - 1911).ToString("D3") : eCalYear.Text.Trim();
            string vLastYear = (Int32.Parse(vCalYear.Trim()) - 1).ToString();
            string vCalYM = vCalYear + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2");
            string vLastYM = vLastYear + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2");
            string vCalMonth = "Month" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2");
            string vCalYM_Range = "";
            string vLastYM_Range = "";
            string vCalMonth_Range = "";

            //定義幾個項目的索引
            int vINTZZ_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'INC ZZ'", "IndexNo_IS")) + 3;
            int vCOS301_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'COS301'", "IndexNo_IS")) + 3;
            int vCOS302_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'COS302'", "IndexNo_IS")) + 3;
            int vKMS01_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'KMS 01'", "IndexNo_IS")) + 3;
            int vKMS02_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'KMS 02'", "IndexNo_IS")) + 3;
            int vKMS03_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'KMS 03'", "IndexNo_IS")) + 3;
            int vKMS04_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'KMS 04'", "IndexNo_IS")) + 3;
            int vKMS05_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'KMS 05'", "IndexNo_IS")) + 3;
            int vKMS100_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'KMS100'", "IndexNo_IS")) + 3;
            int vCars01_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'CARS01'", "IndexNo_IS")) + 3;
            int vAVG01_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'AVG 01'", "IndexNo_IS")) + 3;
            int vAVG02_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'AVG 02'", "IndexNo_IS")) + 3;
            int vAVG03_Index = Int32.Parse(PF.GetValue(vConnStr, "select IndexNo_IS from SHR1202Lines where ChartBarCode = 'AVG 03'", "IndexNo_IS")) + 3;

            for (int i = 1; i <= Int32.Parse(eCalMonth.Text.Trim()); i++)
            {
                vCalYM_Range = (vCalYM_Range == "") ? "'" + vCalYear + i.ToString("D2") + "'" :
                               vCalYM_Range + ", '" + vCalYear + i.ToString("D2") + "'";
                vLastYM_Range = (vLastYM_Range == "") ? "'" + vLastYear + i.ToString("D2") + "'" :
                                vLastYM_Range + ", '" + vLastYear + i.ToString("D2") + "'";
                vCalMonth_Range = (vCalMonth_Range == "") ? "Month" + i.ToString("D2") :
                                  vCalMonth_Range + " + Month" + i.ToString("D2");
            }
            string vDepNoSelStr = "select distinct DepNo from DepMonthIncome where IncomeYM = '" + vCalYM + "' and DepNo not in ('12' ,'21') order by DepNo";
            string vSelectStr = "";
            string vCalDepNo = "";
            string vCalDepName = "";
            //2023/10/26 改用 MARS 多重結果作用集
            //using (SqlConnection connDepNo = new SqlConnection(vConnStr))
            SqlConnectionStringBuilder vConnStr_Temp = new SqlConnectionStringBuilder(vConnStr);
            SqlConnectionStringBuilder vConnStr_Multi = new SqlConnectionStringBuilder();
            using (SqlConnection connMulti = new SqlConnection())
            {
                vConnStr_Multi.InitialCatalog = vConnStr_Temp.InitialCatalog;
                vConnStr_Multi.DataSource = vConnStr_Temp.DataSource;
                vConnStr_Multi.UserID = vConnStr_Temp.UserID;
                vConnStr_Multi.Password = vConnStr_Temp.Password;
                vConnStr_Multi.MultipleActiveResultSets = true;
                connMulti.ConnectionString = vConnStr_Multi.ConnectionString;
                //SqlCommand cmdDepNo = new SqlCommand(vDepNoSelStr, connDepNo);
                //connDepNo.Open();
                SqlCommand cmdDepNo = new SqlCommand(vDepNoSelStr, connMulti);
                connMulti.Open();
                SqlDataReader drDepNo = cmdDepNo.ExecuteReader();
                if (drDepNo.HasRows)
                {
                    string vFileName = eCalYear.Text.Trim() + "年" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "月客運業務損益表";
                    string vSheetName = "";
                    string vHeaderText = "";
                    int vLinesNo = 0;
                    double vCellValue = 0.0;
                    string vTempStr = "";
                    string vDepNoYMStr = "";
                    string vCellStr = "";

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

                    XSSFCellStyle csDoubleData = (XSSFCellStyle)wbExcel.CreateCellStyle();
                    csDoubleData.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csDoubleData.Alignment = HorizontalAlignment.Right;

                    XSSFDataFormat formatDoubleData = (XSSFDataFormat)wbExcel.CreateDataFormat();
                    csDoubleData.DataFormat = formatDoubleData.GetFormat("###,##0.00");

                    XSSFFont fontData = (XSSFFont)wbExcel.CreateFont();
                    fontData.FontHeightInPoints = 12;
                    csData.SetFont(fontData);
                    csDoubleData.SetFont(fontData);

                    while (drDepNo.Read())
                    {
                        //取得部門編號
                        vCalDepNo = drDepNo["DepnO"].ToString().Trim();
                        //取得部門名稱
                        vCalDepName = (vCalDepNo != "00") ? PF.GetValue(vConnStr, "select [Name] from Department where DepNo = '" + vCalDepNo + "' ", "Name") : "業務部";
                        //決定工作表標題
                        vSheetName = vFileName + "(" + vCalDepName + ")";
                        //開始取資料
                        vSelectStr = "select ChartBarCode, LinesName_IS, " + Environment.NewLine +
                                     "       (select sum(SubjectAMT + OtherAMT + ShareAMT) from DepMonthIncome where IncomeYM = '" + vCalYM + "' and DepNo = '" + vCalDepNo + "' and ChartBarCode = z.ChartBarCode) as MonthAMT,cast(0 as float) as MonthAMT_KM, " + Environment.NewLine +
                                     "       case when ChartBarCode = 'COS 28' " + Environment.NewLine +
                                     "            then (select sum(" + vCalMonth + ") from DepMonthBudget where BudgetYear = '" + vCalYear + "' and ChartBarCode in ('COS 28', 'COS 56') and DepNo = '" + vCalDepNo + "') " + Environment.NewLine +
                                     "            else (select " + vCalMonth + " from DepMonthBudget where BudgetYear = '" + vCalYear + "' and ChartBarCode = z.ChartBarCode and DepNo = '" + vCalDepNo + "') end as MonthBudgetAMT, " + Environment.NewLine +
                                     "       cast(0 as float) as MonthBudgetAMT_KM, " + Environment.NewLine +
                                     "       (select sum(SubjectAMT + OtherAMT + ShareAMT) from DepMonthIncome where IncomeYM = '" + vLastYM + "' and DepNo = '" + vCalDepNo + "' and ChartBarCode = z.ChartBarCode) as LastYearMonthAMT,cast(0 as float) as LastYearMonthAMT_KM, " + Environment.NewLine +
                                     "       cast(0 as float) as GrowthRatio, " + Environment.NewLine +
                                     "       (select sum(SubjectAMT + OtherAMT + ShareAMT) from DepMonthIncome where IncomeYM in (" + vCalYM_Range + ") and DepNo = '" + vCalDepNo + "' and ChartBarCode = z.ChartBarCode) as MonthAMT_T,cast(0 as float) as MonthAMT_KM_T, " + Environment.NewLine +
                                     "       case when ChartBarCode = 'COS 28' " + Environment.NewLine +
                                     "            then (select sum(" + vCalMonth_Range + ") from DepMonthBudget where BudgetYear = '" + vCalYear + "' and ChartBarCode in ('COS 28', 'COS 56') and DepNo = '" + vCalDepNo + "') " + Environment.NewLine +
                                     "            else (select " + vCalMonth_Range + " from DepMonthBudget where BudgetYear = '" + vCalYear + "' and ChartBarCode = z.ChartBarCode and DepNo = '" + vCalDepNo + "') end as MonthBudgetAMT_T, " + Environment.NewLine +
                                     "       cast(0 as float) as MonthBudgetAMT_KM_T, " + Environment.NewLine +
                                     "       (select sum(SubjectAMT + OtherAMT + ShareAMT) from DepMonthIncome where IncomeYM in (" + vLastYM_Range + ") and DepNo = '" + vCalDepNo + "' and ChartBarCode = z.ChartBarCode) as LastYearMonthAMT_T,cast(0 as float) as LastYearMonthAMT_KM_T, " + Environment.NewLine +
                                     "       cast(0 as float) as GrowthRatio_T " + Environment.NewLine +
                                     "  from SHR1202Lines z " + Environment.NewLine +
                                     " where isnull(IndexNo_IS, '') <> '' " + Environment.NewLine +
                                     " order by z.IndexNo_IS";
                        //2023/10/26 改用 MARS 多重結果作用集
                        //using (SqlConnection connGetData = new SqlConnection(vConnStr))
                        //{
                        //SqlCommand cmdGetData = new SqlCommand(vSelectStr, connGetData);
                        //connGetData.Open();
                        SqlCommand cmdGetData = new SqlCommand(vSelectStr, connMulti);
                        SqlDataReader drGetData = cmdGetData.ExecuteReader();
                        if (drGetData.HasRows)
                        {
                            //新增一個工作表
                            wsExcel = (XSSFSheet)wbExcel.CreateSheet(vSheetName);
                            //寫入標題列
                            vLinesNo = 0;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(0, CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(0).SetCellValue(vSheetName);
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 15)); //設定合併儲存格, 4個值分別是 (起始行, 結束行, 起始欄, 結束欄)
                            wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csTitle;
                            //標題列 2
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            wsExcel.GetRow(vLinesNo).CreateCell(0);
                            wsExcel.GetRow(vLinesNo).GetCell(0).SetCellValue("");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 1));
                            wsExcel.GetRow(vLinesNo).CreateCell(2);
                            wsExcel.GetRow(vLinesNo).GetCell(2).SetCellValue(vCalYear + "/" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "實際");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 2, 3));
                            wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csTitle;
                            wsExcel.GetRow(vLinesNo).CreateCell(4);
                            wsExcel.GetRow(vLinesNo).GetCell(4).SetCellValue(vCalYear + "/" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "預算");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 4, 5));
                            wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csTitle;
                            wsExcel.GetRow(vLinesNo).CreateCell(6);
                            wsExcel.GetRow(vLinesNo).GetCell(6).SetCellValue(vLastYear + "/" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "實際");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 6, 7));
                            wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csTitle;
                            wsExcel.GetRow(vLinesNo).CreateCell(8);
                            wsExcel.GetRow(vLinesNo).GetCell(8).SetCellValue("成長率%");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 8, 8));
                            wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csTitle;
                            wsExcel.GetRow(vLinesNo).CreateCell(9);
                            wsExcel.GetRow(vLinesNo).GetCell(9).SetCellValue(vCalYear + "/01-" + vCalYear + "/" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "實際");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 9, 10));
                            wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csTitle;
                            wsExcel.GetRow(vLinesNo).CreateCell(11);
                            wsExcel.GetRow(vLinesNo).GetCell(11).SetCellValue(vCalYear + "/01-" + vCalYear + "/" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "預算");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 11, 12));
                            wsExcel.GetRow(vLinesNo).GetCell(11).CellStyle = csTitle;
                            wsExcel.GetRow(vLinesNo).CreateCell(13);
                            wsExcel.GetRow(vLinesNo).GetCell(13).SetCellValue(vLastYear + "/01-" + vLastYear + "/" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "實際");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 13, 14));
                            wsExcel.GetRow(vLinesNo).GetCell(13).CellStyle = csTitle;
                            wsExcel.GetRow(vLinesNo).CreateCell(15);
                            wsExcel.GetRow(vLinesNo).GetCell(15).SetCellValue("成長率%");
                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo + 1, 15, 15));
                            wsExcel.GetRow(vLinesNo).GetCell(15).CellStyle = csTitle;
                            //標題列 3
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            for (int i = 0; i < drGetData.FieldCount; i++)
                            {
                                vHeaderText = (drGetData.GetName(i) == "ChartBarCode") ? "科目" :
                                              (drGetData.GetName(i) == "LinesName_IS") ? "項目" :
                                              ((drGetData.GetName(i) == "MonthAMT") ||
                                               (drGetData.GetName(i) == "MonthBudgetAMT") ||
                                               (drGetData.GetName(i) == "LastYearMonthAMT") ||
                                               (drGetData.GetName(i) == "MonthAMT_T") ||
                                               (drGetData.GetName(i) == "MonthBudgetAMT_T") ||
                                               (drGetData.GetName(i) == "LastYearMonthAMT_T")) ? "金　　額" :
                                              ((drGetData.GetName(i) == "MonthAMT_KM") ||
                                               (drGetData.GetName(i) == "LastYearMonthAMT_KM") ||
                                               (drGetData.GetName(i) == "MonthAMT_KM_T") ||
                                               (drGetData.GetName(i) == "LastYearMonthAMT_KM_T")) ? "每公里收支" :
                                              ((drGetData.GetName(i) == "MonthBudgetAMT_KM") ||
                                               (drGetData.GetName(i) == "MonthBudgetAMT_KM_T")) ? "達成率%" : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(i);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                            }

                            while (drGetData.Read())
                            {
                                vLinesNo++;
                                wsExcel.CreateRow(vLinesNo);
                                for (int i = 0; i < drGetData.FieldCount; i++)
                                {
                                    wsExcel.GetRow(vLinesNo).CreateCell(i);
                                    if (drGetData.GetName(i) == "LinesName_IS") //項目名稱
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drGetData.GetString(i));
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                    }
                                    else if (drGetData.GetName(i) == "MonthBudgetAMT_KM") //當月預算達成率
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("IF(E" + (vLinesNo + 1).ToString() + " <> 0, ROUND((C" + (vLinesNo + 1).ToString() + " / ABS(E" + (vLinesNo + 1).ToString() + ")) * 100, 2), IF(C" + (vLinesNo + 1).ToString() + " <> 0, 100, 0))");
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "MonthBudgetAMT_KM_T") //累計預算達成率
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("IF(L" + (vLinesNo + 1).ToString() + " <> 0, ROUND((J" + (vLinesNo + 1).ToString() + " / ABS(L" + (vLinesNo + 1).ToString() + ")) * 100, 2), IF(J" + (vLinesNo + 1).ToString() + " <> 0, 100, 0))");
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "GrowthRatio") //當月成長率
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("IF(G" + (vLinesNo + 1).ToString() + " <> 0, ROUND((C" + (vLinesNo + 1).ToString() + " - G" + (vLinesNo + 1).ToString() + ") / ABS(G" + (vLinesNo + 1).ToString() + ") * 100, 2), IF(C" + (vLinesNo + 1).ToString() + " <> 0, 100, 0))");
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "GrowthRatio_T") //累計成長率
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("IF(N" + (vLinesNo + 1).ToString() + " <> 0, ROUND((J" + (vLinesNo + 1).ToString() + " - N" + (vLinesNo + 1).ToString() + ") / ABS(N" + (vLinesNo + 1).ToString() + ") * 100, 2), IF(J" + (vLinesNo + 1).ToString() + " <> 0, 100, 0))");
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "ChartBarCode") //第一欄 (空欄)
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("");
                                        //wsExcel.AutoSizeColumn(i);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                    }
                                    else if (drGetData.GetName(i) == "MonthAMT") //當月實際
                                    {
                                        if (drGetData.GetString(0) == "COS302")
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("C" + vINTZZ_Index.ToString() + "-C" + vCOS301_Index.ToString());
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 01")
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("ROUND(C" + vINTZZ_Index.ToString() + "/C" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 02")
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("ROUND(C" + vCOS301_Index.ToString() + "/C" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 03")
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("ROUND(C" + vCOS302_Index.ToString() + "/C" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "COS 28")
                                        {
                                            vTempStr = "select Sum(SubjectAMT + OtherAMT + ShareAMT) MonthAMT " + Environment.NewLine +
                                                       "  from DepMonthIncome " + Environment.NewLine +
                                                       " where IncomeYM = '" + vCalYM + "' " + Environment.NewLine +
                                                       "   and DepNo = '" + vCalDepNo + "' " + Environment.NewLine +
                                                       "   and ChartBarCode in ('COS 28', 'COS 56')";
                                            vCellValue = Double.Parse(PF.GetValue(vConnStr, vTempStr, "MonthAMT"));
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                        }
                                        else
                                        {
                                            if (vCalDepNo == "00")
                                            {
                                                vTempStr = (drGetData.GetString(0) == "KMS 01") ? "select sum(KMS_Bus + KMS_Rule + KMS_NoneBusi) TotalKM from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vCalYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS 02") ? "select sum(KMS_Tour )KMS_Tour from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vCalYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS 03") ? "select sum(KMS_Rent) KMS_Rent from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vCalYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS 04") ? "select sum(KMS_Trans) KMS_Trans from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vCalYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS 05") ? "select sum(KMS_Spec) KMS_Spec from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vCalYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS100") ? "select sum(DriveRange) DriveRange from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vCalYM + "%' " :
                                                           (drGetData.GetString(0) == "CARS01") ? "select sum(CarCount) CarCount from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vCalYM + "%' " : "";
                                            }
                                            else
                                            {
                                                vTempStr = (drGetData.GetString(0) == "KMS 01") ? "select (KMS_Bus + KMS_Rule + KMS_NoneBusi) TotalKM from CarCount where DepNoYM = '" + vCalYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS 02") ? "select KMS_Tour from CarCount where DepNoYM = '" + vCalYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS 03") ? "select KMS_Rent from CarCount where DepNoYM = '" + vCalYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS 04") ? "select KMS_Trans from CarCount where DepNoYM = '" + vCalYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS 05") ? "select KMS_Spec from CarCount where DepNoYM = '" + vCalYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS100") ? "select DriveRange from CarCount where DepNoYM = '" + vCalYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "CARS01") ? "select CarCount from CarCount where DepNoYM = '" + vCalYM + vCalDepNo + "' " : "";
                                            }
                                            vCellValue = (drGetData.GetString(0) == "KMS 01") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "TotalKM")) :
                                                         (drGetData.GetString(0) == "KMS 02") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Tour")) :
                                                         (drGetData.GetString(0) == "KMS 03") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Rent")) :
                                                         (drGetData.GetString(0) == "KMS 04") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Trans")) :
                                                         (drGetData.GetString(0) == "KMS 05") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Spec")) :
                                                         (drGetData.GetString(0) == "KMS100") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "DriveRange")) :
                                                         (drGetData.GetString(0) == "CARS01") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "CarCount")) :
                                                         (drGetData.IsDBNull(drGetData.GetOrdinal(drGetData.GetName(i)))) ? 0.0 : drGetData.GetDouble(i);
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                        }
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "LastYearMonthAMT") //去年同期實際
                                    {
                                        if (drGetData.GetString(0) == "COS302")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("G" + vINTZZ_Index.ToString() + "-G" + vCOS301_Index.ToString());
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 01")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(G" + vINTZZ_Index.ToString() + "/G" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 02")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(G" + vCOS301_Index.ToString() + "/G" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 03")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(G" + vCOS302_Index.ToString() + "/G" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "COS 28")
                                        {
                                            vTempStr = "select Sum(SubjectAMT + OtherAMT + ShareAMT) MonthAMT " + Environment.NewLine +
                                                       "  from DepMonthIncome " + Environment.NewLine +
                                                       " where IncomeYM = '" + vLastYM + "' " + Environment.NewLine +
                                                       "   and DepNo = '" + vCalDepNo + "' " + Environment.NewLine +
                                                       "   and ChartBarCode in ('COS 28', 'COS 56')";
                                            vCellValue = Double.Parse(PF.GetValue(vConnStr, vTempStr, "MonthAMT"));
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                        }
                                        else
                                        {
                                            if (vCalDepNo == "00")
                                            {
                                                vTempStr = (drGetData.GetString(0) == "KMS 01") ? "select sum(KMS_Bus + KMS_Rule + KMS_NoneBusi) TotalKM from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vLastYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS 02") ? "select sum(KMS_Tour )KMS_Tour from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vLastYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS 03") ? "select sum(KMS_Rent) KMS_Rent from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vLastYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS 04") ? "select sum(KMS_Trans) KMS_Trans from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vLastYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS 05") ? "select sum(KMS_Spec) KMS_Spec from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vLastYM + "%' " :
                                                           (drGetData.GetString(0) == "KMS100") ? "select sum(DriveRange) DriveRange from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vLastYM + "%' " :
                                                           (drGetData.GetString(0) == "CARS01") ? "select sum(CarCount) CarCount from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM like '" + vLastYM + "%' " : "";
                                            }
                                            else
                                            {
                                                vTempStr = (drGetData.GetString(0) == "KMS 01") ? "select (KMS_Bus + KMS_Rule + KMS_NoneBusi) TotalKM from CarCount where DepNoYM = '" + vLastYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS 02") ? "select KMS_Tour from CarCount where DepNoYM = '" + vLastYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS 03") ? "select KMS_Rent from CarCount where DepNoYM = '" + vLastYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS 04") ? "select KMS_Trans from CarCount where DepNoYM = '" + vLastYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS 05") ? "select KMS_Spec from CarCount where DepNoYM = '" + vLastYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "KMS100") ? "select DriveRange from CarCount where DepNoYM = '" + vLastYM + vCalDepNo + "' " :
                                                           (drGetData.GetString(0) == "CARS01") ? "select CarCount from CarCount where DepNoYM = '" + vLastYM + vCalDepNo + "' " : "";
                                            }
                                            vCellValue = (drGetData.GetString(0) == "KMS 01") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "TotalKM")) :
                                                         (drGetData.GetString(0) == "KMS 02") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Tour")) :
                                                         (drGetData.GetString(0) == "KMS 03") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Rent")) :
                                                         (drGetData.GetString(0) == "KMS 04") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Trans")) :
                                                         (drGetData.GetString(0) == "KMS 05") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Spec")) :
                                                         (drGetData.GetString(0) == "KMS100") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "DriveRange")) :
                                                         (drGetData.GetString(0) == "CARS01") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "CarCount")) :
                                                         (drGetData.IsDBNull(drGetData.GetOrdinal(drGetData.GetName(i)))) ? 0.0 : drGetData.GetDouble(i);
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                        }
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "MonthAMT_T") //當年度累計
                                    {
                                        vDepNoYMStr = "";
                                        if (drGetData.GetString(0) == "COS302")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("J" + vINTZZ_Index.ToString() + "-J" + vCOS301_Index.ToString());
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 01")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(J" + vINTZZ_Index.ToString() + "/J" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 02")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(J" + vCOS301_Index.ToString() + "/J" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 03")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(J" + vCOS302_Index.ToString() + "/J" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "COS 28")
                                        {
                                            vTempStr = "select Sum(SubjectAMT + OtherAMT + ShareAMT) MonthAMT " + Environment.NewLine +
                                                       "  from DepMonthIncome " + Environment.NewLine +
                                                       " where IncomeYM between '" + vCalYear + "01' and '" + vCalYM + "' " + Environment.NewLine +
                                                       "   and DepNo = '" + vCalDepNo + "' " + Environment.NewLine +
                                                       "   and ChartBarCode in ('COS 28', 'COS 56')";
                                            vCellValue = Double.Parse(PF.GetValue(vConnStr, vTempStr, "MonthAMT"));
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                        }
                                        else
                                        {
                                            if (vCalDepNo == "00")
                                            {
                                                string vDepNoSelStr_Sub = "select distinct DepNo from DepMonthIncome where IncomeYM = '" + vCalYM + "' and DepNo not in ('00', '12' ,'21') order by DepNo";

                                                for (int vMonth = 1; vMonth <= Int32.Parse(eCalMonth.Text.Trim()); vMonth++)
                                                {
                                                    using (SqlConnection connTempDepNo = new SqlConnection(vConnStr))
                                                    {
                                                        SqlCommand cmdTempDepNo = new SqlCommand(vDepNoSelStr_Sub, connTempDepNo);
                                                        connTempDepNo.Open();
                                                        SqlDataReader drTempDepNo = cmdTempDepNo.ExecuteReader();
                                                        if (drTempDepNo.HasRows)
                                                        {
                                                            while (drTempDepNo.Read())
                                                            {
                                                                vDepNoYMStr = (vDepNoYMStr == "") ?
                                                                              "'" + vCalYear + vMonth.ToString("D2") + drTempDepNo[0].ToString().Trim() + "'" :
                                                                              vDepNoYMStr + ", '" + vCalYear + vMonth.ToString("D2") + drTempDepNo[0].ToString().Trim() + "'";
                                                            }
                                                        }
                                                        drTempDepNo.Close();
                                                        cmdTempDepNo.Cancel();
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                for (int vMonth = 1; vMonth <= Int32.Parse(eCalMonth.Text.Trim()); vMonth++)
                                                {
                                                    vDepNoYMStr = (vDepNoYMStr == "") ?
                                                                  "'" + vCalYear + vMonth.ToString("D2") + vCalDepNo + "'" :
                                                                  vDepNoYMStr + ", '" + vCalYear + vMonth.ToString("D2") + vCalDepNo + "'";
                                                }
                                            }
                                            vTempStr = (drGetData.GetString(0) == "KMS 01") ? "select sum(KMS_Bus + KMS_Rule + KMS_NoneBusi) TotalKM from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS 02") ? "select sum(KMS_Tour) KMS_Tour from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS 03") ? "select sum(KMS_Rent) KMS_Rent from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS 04") ? "select sum(KMS_Trans) KMS_Trans from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS 05") ? "select sum(KMS_Spec) KMS_Spec from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS100") ? "select sum(DriveRange) DriveRange from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "CARS01") ? "select sum(CarCount) CarCount from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " : "";
                                            vCellValue = (drGetData.GetString(0) == "KMS 01") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "TotalKM")) :
                                                         (drGetData.GetString(0) == "KMS 02") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Tour")) :
                                                         (drGetData.GetString(0) == "KMS 03") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Rent")) :
                                                         (drGetData.GetString(0) == "KMS 04") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Trans")) :
                                                         (drGetData.GetString(0) == "KMS 05") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Spec")) :
                                                         (drGetData.GetString(0) == "KMS100") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "DriveRange")) :
                                                         (drGetData.GetString(0) == "CARS01") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "CarCount")) :
                                                         (drGetData.IsDBNull(drGetData.GetOrdinal(drGetData.GetName(i)))) ? 0.0 : drGetData.GetDouble(i);
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                        }
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "LastYearMonthAMT_T") //去年同期累計
                                    {
                                        vDepNoYMStr = "";
                                        if (drGetData.GetString(0) == "COS302")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("N" + vINTZZ_Index.ToString() + "-N" + vCOS301_Index.ToString());
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 01")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(N" + vINTZZ_Index.ToString() + "/N" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 02")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(N" + vCOS301_Index.ToString() + "/N" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 03")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(N" + vCOS302_Index.ToString() + "/N" + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "COS 28")
                                        {
                                            vTempStr = "select Sum(SubjectAMT + OtherAMT + ShareAMT) MonthAMT " + Environment.NewLine +
                                                       "  from DepMonthIncome " + Environment.NewLine +
                                                       " where IncomeYM between '" + vLastYear + "01' and '" + vLastYM + "' " + Environment.NewLine +
                                                       "   and DepNo = '" + vCalDepNo + "' " + Environment.NewLine +
                                                       "   and ChartBarCode in ('COS 28', 'COS 56')";
                                            vCellValue = Double.Parse(PF.GetValue(vConnStr, vTempStr, "MonthAMT"));
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                        }
                                        else
                                        {
                                            if (vCalDepNo == "00")
                                            {
                                                string vDepNoSelStr_Sub = "select distinct DepNo from DepMonthIncome where IncomeYM = '" + vCalYM + "' and DepNo not in ('00', '12' ,'21') order by DepNo";

                                                for (int vMonth = 1; vMonth <= Int32.Parse(eCalMonth.Text.Trim()); vMonth++)
                                                {
                                                    using (SqlConnection connTempDepNo = new SqlConnection(vConnStr))
                                                    {
                                                        SqlCommand cmdTempDepNo = new SqlCommand(vDepNoSelStr_Sub, connTempDepNo);
                                                        connTempDepNo.Open();
                                                        SqlDataReader drTempDepNo = cmdTempDepNo.ExecuteReader();
                                                        if (drTempDepNo.HasRows)
                                                        {
                                                            while (drTempDepNo.Read())
                                                            {
                                                                vDepNoYMStr = (vDepNoYMStr == "") ?
                                                                              "'" + vLastYear + vMonth.ToString("D2") + drTempDepNo[0].ToString().Trim() + "'" :
                                                                              vDepNoYMStr + ", '" + vLastYear + vMonth.ToString("D2") + drTempDepNo[0].ToString().Trim() + "'";
                                                            }
                                                        }
                                                        drTempDepNo.Close();
                                                        cmdTempDepNo.Cancel();
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                for (int vMonth = 1; vMonth <= Int32.Parse(eCalMonth.Text.Trim()); vMonth++)
                                                {
                                                    vDepNoYMStr = (vDepNoYMStr == "") ?
                                                                  "'" + vLastYear + vMonth.ToString("D2") + vCalDepNo + "'" :
                                                                  vDepNoYMStr + ", '" + vLastYear + vMonth.ToString("D2") + vCalDepNo + "'";
                                                }
                                            }
                                            vTempStr = (drGetData.GetString(0) == "KMS 01") ? "select sum(KMS_Bus + KMS_Rule + KMS_NoneBusi) TotalKM from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS 02") ? "select sum(KMS_Tour) KMS_Tour from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS 03") ? "select sum(KMS_Rent) KMS_Rent from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS 04") ? "select sum(KMS_Trans) KMS_Trans from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS 05") ? "select sum(KMS_Spec) KMS_Spec from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "KMS100") ? "select sum(DriveRange) DriveRange from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " :
                                                       (drGetData.GetString(0) == "CARS01") ? "select sum(CarCount) CarCount from CarCount where DepNo in (select DepNo from Department where InSHReport = 'V') and DepNoYM in (" + vDepNoYMStr + ") " : "";
                                            vCellValue = (drGetData.GetString(0) == "KMS 01") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "TotalKM")) :
                                                         (drGetData.GetString(0) == "KMS 02") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Tour")) :
                                                         (drGetData.GetString(0) == "KMS 03") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Rent")) :
                                                         (drGetData.GetString(0) == "KMS 04") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Trans")) :
                                                         (drGetData.GetString(0) == "KMS 05") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "KMS_Spec")) :
                                                         (drGetData.GetString(0) == "KMS100") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "DriveRange")) :
                                                         (drGetData.GetString(0) == "CARS01") ? Double.Parse(PF.GetValue(vConnStr, vTempStr, "CarCount")) :
                                                         (drGetData.IsDBNull(drGetData.GetOrdinal(drGetData.GetName(i)))) ? 0.0 : drGetData.GetDouble(i);
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                        }
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "MonthAMT_KM") //當月每公里收支
                                    {
                                        if (drGetData.GetString(0).Substring(0, 3).Trim().ToUpper() == "COS")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(C" + (vLinesNo + 1).ToString() + " <> 0, ROUND(C" + (vLinesNo + 1).ToString() + "/ C" + vKMS100_Index.ToString() + ", 2), 0)");
                                        }
                                        else if (drGetData.GetString(0).Substring(0, 3).Trim().ToUpper() == "INC")
                                        {
                                            if ((drGetData.GetString(0).ToUpper() == "INC 10") || (drGetData.GetString(0).ToUpper() == "INC 20") || (drGetData.GetString(0).ToUpper() == "INC 55"))
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(C" + (vLinesNo + 1).ToString() + " <> 0, ROUND(C" + (vLinesNo + 1).ToString() + "/ C" + vKMS01_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 30")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(C" + (vLinesNo + 1).ToString() + " <> 0, ROUND(C" + (vLinesNo + 1).ToString() + "/ C" + vKMS02_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 40")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(C" + (vLinesNo + 1).ToString() + " <> 0, ROUND(C" + (vLinesNo + 1).ToString() + "/ C" + vKMS03_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 50")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(C" + (vLinesNo + 1).ToString() + " <> 0, ROUND(C" + (vLinesNo + 1).ToString() + "/ C" + vKMS04_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 58")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(C" + (vLinesNo + 1).ToString() + " <> 0, ROUND(C" + (vLinesNo + 1).ToString() + "/ C" + vKMS05_Index.ToString() + ", 2), 0)");
                                            }
                                            else
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(C" + (vLinesNo + 1).ToString() + " <> 0, ROUND(C" + (vLinesNo + 1).ToString() + "/ C" + vKMS100_Index.ToString() + ", 2), 0)");
                                            }
                                        }
                                        else
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(0);
                                        }
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "LastYearMonthAMT_KM") //去年同期每公里收支
                                    {
                                        if (drGetData.GetString(0).Substring(0, 3).Trim().ToUpper() == "COS")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(G" + (vLinesNo + 1).ToString() + " <> 0, ROUND(G" + (vLinesNo + 1).ToString() + "/ G" + vKMS100_Index.ToString() + ", 2), 0)");
                                        }
                                        else if (drGetData.GetString(0).Substring(0, 3).Trim().ToUpper() == "INC")
                                        {
                                            if ((drGetData.GetString(0).ToUpper() == "INC 10") || (drGetData.GetString(0).ToUpper() == "INC 20") || (drGetData.GetString(0).ToUpper() == "INC 55"))
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(G" + (vLinesNo + 1).ToString() + " <> 0, ROUND(G" + (vLinesNo + 1).ToString() + "/ G" + vKMS01_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 30")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(G" + (vLinesNo + 1).ToString() + " <> 0, ROUND(G" + (vLinesNo + 1).ToString() + "/ G" + vKMS02_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 40")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(G" + (vLinesNo + 1).ToString() + " <> 0, ROUND(G" + (vLinesNo + 1).ToString() + "/ G" + vKMS03_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 50")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(G" + (vLinesNo + 1).ToString() + " <> 0, ROUND(G" + (vLinesNo + 1).ToString() + "/ G" + vKMS04_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 58")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(G" + (vLinesNo + 1).ToString() + " <> 0, ROUND(G" + (vLinesNo + 1).ToString() + "/ G" + vKMS05_Index.ToString() + ", 2), 0)");
                                            }
                                            else
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(G" + (vLinesNo + 1).ToString() + " <> 0, ROUND(G" + (vLinesNo + 1).ToString() + "/ G" + vKMS100_Index.ToString() + ", 2), 0)");
                                            }
                                        }
                                        else
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(0);
                                        }
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "MonthAMT_KM_T") //當年度每公里收支累計
                                    {
                                        if (drGetData.GetString(0).Substring(0, 3).Trim().ToUpper() == "COS")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(J" + (vLinesNo + 1).ToString() + " <> 0, ROUND(J" + (vLinesNo + 1).ToString() + "/ J" + vKMS100_Index.ToString() + ", 2), 0)");
                                        }
                                        else if (drGetData.GetString(0).Substring(0, 3).Trim().ToUpper() == "INC")
                                        {
                                            if ((drGetData.GetString(0).ToUpper() == "INC 10") || (drGetData.GetString(0).ToUpper() == "INC 20") || (drGetData.GetString(0).ToUpper() == "INC 55"))
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(J" + (vLinesNo + 1).ToString() + " <> 0, ROUND(J" + (vLinesNo + 1).ToString() + "/ J" + vKMS01_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 30")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(J" + (vLinesNo + 1).ToString() + " <> 0, ROUND(J" + (vLinesNo + 1).ToString() + "/ J" + vKMS02_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 40")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(J" + (vLinesNo + 1).ToString() + " <> 0, ROUND(J" + (vLinesNo + 1).ToString() + "/ J" + vKMS03_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 50")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(J" + (vLinesNo + 1).ToString() + " <> 0, ROUND(J" + (vLinesNo + 1).ToString() + "/ J" + vKMS04_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 58")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(J" + (vLinesNo + 1).ToString() + " <> 0, ROUND(J" + (vLinesNo + 1).ToString() + "/ J" + vKMS05_Index.ToString() + ", 2), 0)");
                                            }
                                            else
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(J" + (vLinesNo + 1).ToString() + " <> 0, ROUND(J" + (vLinesNo + 1).ToString() + "/ J" + vKMS100_Index.ToString() + ", 2), 0)");
                                            }
                                        }
                                        else
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(0);
                                        }
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if (drGetData.GetName(i) == "LastYearMonthAMT_KM_T") //去年每公里收支累計
                                    {
                                        if (drGetData.GetString(0).Substring(0, 3).Trim().ToUpper() == "COS")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(N" + (vLinesNo + 1).ToString() + " <> 0, ROUND(N" + (vLinesNo + 1).ToString() + "/ N" + vKMS100_Index.ToString() + ", 2), 0)");
                                        }
                                        else if (drGetData.GetString(0).Substring(0, 3).Trim().ToUpper() == "INC")
                                        {
                                            if ((drGetData.GetString(0).ToUpper() == "INC 10") || (drGetData.GetString(0).ToUpper() == "INC 20") || (drGetData.GetString(0).ToUpper() == "INC 55"))
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(N" + (vLinesNo + 1).ToString() + " <> 0, ROUND(N" + (vLinesNo + 1).ToString() + "/ N" + vKMS01_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 30")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(N" + (vLinesNo + 1).ToString() + " <> 0, ROUND(N" + (vLinesNo + 1).ToString() + "/ N" + vKMS02_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 40")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(N" + (vLinesNo + 1).ToString() + " <> 0, ROUND(N" + (vLinesNo + 1).ToString() + "/ N" + vKMS03_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 50")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(N" + (vLinesNo + 1).ToString() + " <> 0, ROUND(N" + (vLinesNo + 1).ToString() + "/ N" + vKMS04_Index.ToString() + ", 2), 0)");
                                            }
                                            else if (drGetData.GetString(0).ToUpper() == "INC 58")
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(N" + (vLinesNo + 1).ToString() + " <> 0, ROUND(N" + (vLinesNo + 1).ToString() + "/ N" + vKMS05_Index.ToString() + ", 2), 0)");
                                            }
                                            else
                                            {
                                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("IF(N" + (vLinesNo + 1).ToString() + " <> 0, ROUND(N" + (vLinesNo + 1).ToString() + "/ N" + vKMS100_Index.ToString() + ", 2), 0)");
                                            }
                                        }
                                        else
                                        {
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(0);
                                        }
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                    }
                                    else if ((drGetData.GetName(i) == "MonthBudgetAMT") || (drGetData.GetName(i) == "MonthBudgetAMT_T")) //預算
                                    {
                                        vCellStr = ((char)(i + 65)).ToString();
                                        if (drGetData.GetString(0) == "AVG 01")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(" + vCellStr + vINTZZ_Index.ToString() + "/" + vCellStr + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 02")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(" + vCellStr + vCOS301_Index.ToString() + "/" + vCellStr + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else if (drGetData.GetString(0) == "AVG 03")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellFormula("ROUND(" + vCellStr + vCOS302_Index.ToString() + "/" + vCellStr + vCars01_Index.ToString() + ", 2)");
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                        else
                                        {
                                            vCellValue = (drGetData.IsDBNull(drGetData.GetOrdinal(drGetData.GetName(i)))) ? 0.0 : drGetData.GetDouble(i);
                                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;
                                        }
                                    }
                                    else
                                    {
                                        vCellValue = (drGetData.IsDBNull(drGetData.GetOrdinal(drGetData.GetName(i)))) ? 0.0 : drGetData.GetDouble(i);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vCellValue);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csDoubleData;

                                    }
                                }
                            }
                        }
                        drGetData.Close();
                        cmdGetData.Cancel();
                        //}
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
                            string vRecordCalYMStr = eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月";
                            string vRecordBudgetYearStr = eBudgetYear.Text.Trim() + "年度";
                            string vRecordNote = "匯出檔案_客運業務損益表" + Environment.NewLine +
                                                 "CalDepIncomeStatement.aspx" + Environment.NewLine +
                                                 "計算年月：" + vRecordCalYMStr + Environment.NewLine +
                                                 "預算年度：" + vRecordBudgetYearStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //===========================================================================================================
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
                drDepNo.Close();
                cmdDepNo.Cancel();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}