using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverWorkHourLevelByMan : System.Web.UI.Page
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
        //private string[] DoctorLevelIndex = new string[12] { "DoctorLevel05", "DoctorLevel06", "DoctorLevel07", "DoctorLevel08", "DoctorLevel09", "DoctorLevel10", "DoctorLevel11", "DoctorLevel12", "DoctorLevel01", "DoctorLevel02", "DoctorLevel03", "DoctorLevel04" };

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
                    if (!IsPostBack)
                    {
                        plPrint.Visible = false;
                        plSearch.Visible = true;
                        eDriveYear_Search.Text = (DateTime.Today.AddMonths(-1).Year - 1911).ToString();
                        eDriveMonth_Search.SelectedIndex = DateTime.Today.AddMonths(-1).Month - 1;
                    }
                    else
                    {
                        //OpenData();
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

        private string GetSelectStr_Print()
        {
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ? DateTime.Today.AddMonths(-1) : DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vWStr_BuDate = "          and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine;
            string vWStr_DepNo = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo between '" + eDepNoS_Search.Text.Trim() + "' and '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? "   and e.DepNo = '" + eDepNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo = '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eDriver_Search.Text.Trim() != "") ? "          and Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "select e.DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                "       t.Driver, e.[Name] EmpName, cast(t.WorkTime as float) as MonthTotalHours, " + Environment.NewLine +
                                "       Cast(case when cast(t.WorkTime as float) / 60 <= 222.0 then 1 else 0 end as int) as Level_01, " + Environment.NewLine +
                                "       Cast(case when cast(t.WorkTime as float) / 60 > 222.0 and cast(t.WorkTime as float) / 60 <= 300.0 then 1 else 0 end as int) as Level_02, " + Environment.NewLine +
                                "       Cast(case when cast(t.WorkTime as float) / 60 > 300.0 and cast(t.WorkTime as float) / 60 <= 350.0 then 1 else 0 end as int) as Level_03, " + Environment.NewLine +
                                "       Cast(case when cast(t.WorkTime as float) / 60 > 350.0 and cast(t.WorkTime as float) / 60 <= 400.0 then 1 else 0 end as int) as Level_04, " + Environment.NewLine +
                                "       Cast(case when cast(t.WorkTime as float) / 60 > 400.0  then 1 else 0 end as int) as Level_05 " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "        select Driver, sum(WorkHR * 60 + WorkMin) as WorkTime " + Environment.NewLine +
                                "          from RunSheetA " + Environment.NewLine +
                                "         where isnull(AssignNo, '') <> '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_EmpNo +
                                "         group by Driver " + Environment.NewLine +
                                ") t left join Employee e on e.EmpNo = t.Driver " + Environment.NewLine +
                                " where isnull(t.Driver, '') <> '' " + Environment.NewLine +
                                vWStr_DepNo +
                                " order by e.DepNo, t.Driver ";
            return vResultStr;
        }

        private string GetSelectStr_Excel()
        {
            DateTime vCalDate = ((eDriveYear_Search.Text.Trim() == "") || (eDriveMonth_Search.SelectedValue == "")) ? DateTime.Today.AddMonths(-1) : DateTime.Parse(eDriveYear_Search.Text.Trim() + "/" + eDriveMonth_Search.SelectedValue + "/01");
            string vBuDate_S = PF.GetMonthFirstDay(vCalDate, "B");
            string vBuDate_E = PF.GetMonthLastDay(vCalDate, "B");
            string vWStr_BuDate = "          and BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine;
            string vWStr_DepNo = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo between '" + eDepNoS_Search.Text.Trim() + "' and '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? "   and e.DepNo = '" + eDepNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and e.DepNo = '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eDriver_Search.Text.Trim() != "") ? "          and Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "select e.DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName, " + Environment.NewLine +
                                "       t.Driver, e.[Name] EmpName, cast(t.WorkTime as float) as MonthTotalHours, " + Environment.NewLine +
                                "       case when cast(t.WorkTime as float) / 60 <= 222.0 then 'V' else '' end as Level_01, " + Environment.NewLine +
                                "       case when cast(t.WorkTime as float) / 60 > 222.0 and cast(t.WorkTime as float) / 60 <= 300.0 then 'V' else '' end as Level_02, " + Environment.NewLine +
                                "       case when cast(t.WorkTime as float) / 60 > 300.0 and cast(t.WorkTime as float) / 60 <= 350.0 then 'V' else '' end as Level_03, " + Environment.NewLine +
                                "       case when cast(t.WorkTime as float) / 60 > 350.0 and cast(t.WorkTime as float) / 60 <= 400.0 then 'V' else '' end as Level_04, " + Environment.NewLine +
                                "       case when cast(t.WorkTime as float) / 60 > 400.0  then 'V' else '' end as Level_05 " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "        select Driver, sum(WorkHR * 60 + WorkMin) as WorkTime " + Environment.NewLine +
                                "          from RunSheetA " + Environment.NewLine +
                                "         where isnull(AssignNo, '') <> '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_EmpNo +
                                "         group by Driver " + Environment.NewLine +
                                ") t left join Employee e on e.EmpNo = t.Driver " + Environment.NewLine +
                                " where isnull(t.Driver, '') <> '' " + Environment.NewLine +
                                vWStr_DepNo +
                                " order by e.DepNo, t.Driver ";
            return vResultStr;
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if ((eDriveYear_Search.Text.Trim() != "") && (eDriveMonth_Search.Text.Trim() != ""))
            {
                string vReportName = eDriveYear_Search.Text.Trim() + " 年 " + eDriveMonth_Search.Text.Trim() + " 月 駕駛員工時級數分析表(人)";
                string vSelectStr = GetSelectStr_Print();
                DataTable dtPrint = new DataTable();
                using (SqlConnection connPrint = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                    connPrint.Open();
                    daPrint.Fill(dtPrint);
                    if (dtPrint.Rows.Count > 0)
                    {
                        ReportDataSource rdsPrint = new ReportDataSource("DriverWorkHourLevelByManP", dtPrint);

                        rvPrint.LocalReport.DataSources.Clear();
                        rvPrint.LocalReport.ReportPath = @"Report\DriverWorkHourLevelByManP.rdlc";
                        rvPrint.LocalReport.DataSources.Add(rdsPrint);
                        rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                        rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                        rvPrint.LocalReport.Refresh();
                        plPrint.Visible = true;
                        plSearch.Visible = false;

                        string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                                 ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                                 ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                        string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                        string vRecordNote = "列印資料_" + vReportName + Environment.NewLine +
                                             "DriverWorkHourLevelByMan.aspx" + Environment.NewLine +
                                             "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                             "站別：" + vRecordDepNoStr + Environment.NewLine +
                                             "駕駛員：" + vRecordDriverStr;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('找不到符合條件的資料！')");
                        Response.Write("</" + "Script>");
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('行車日期年月不可空白！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if ((eDriveYear_Search.Text.Trim() != "") && (eDriveMonth_Search.Text.Trim() != ""))
            {
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

                HSSFCellStyle csData_Right = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Right.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Right.Alignment = HorizontalAlignment.Right;

                HSSFFont fontData_Right = (HSSFFont)wbExcel.CreateFont();
                fontData_Right.FontHeightInPoints = 12;
                csData_Right.SetFont(fontData);

                HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
                fontData_Red.FontHeightInPoints = 12; //字體大小
                fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
                csData_Red.SetFont(fontData_Red);

                HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Int.Alignment = HorizontalAlignment.Right;

                HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Int.DataFormat = format.GetFormat("###,##0");

                int vLinesNo = 0;
                string vFileName = eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月駕駛員工時級數分析表(人)";
                string vHeaderText = "";
                string vRiskDepNo = "";
                string vGroupText = "";
                string vCellData = "";
                string vSelectStr = GetSelectStr_Excel();
                DataTable dtExcel = new DataTable();
                using (SqlConnection connExcel = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daExcel = new SqlDataAdapter(vSelectStr, connExcel);
                    connExcel.Open();
                    daExcel.Fill(dtExcel);
                    if (dtExcel.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtExcel.Rows.Count; i++)
                        {
                            if (dtExcel.Rows[i]["DepNo"].ToString().Trim() != vRiskDepNo)
                            {
                                if (vGroupText != "")
                                {
                                    // vGroupText 變數不是空白，表示前面已經有建立過其他單位的頁面，所以要做一個合計
                                    wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                                    vLinesNo++;
                                    wsExcel.CreateRow(vLinesNo);
                                    for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                                    {
                                        vHeaderText = dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim();
                                        if (vHeaderText == "MONTHTOTALHOURS")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue("");
                                            wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData;
                                        }
                                        else if (vHeaderText == "DEPNO")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue("合計：");
                                            wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 2));
                                            wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Right;
                                        }
                                        else if (vHeaderText == "EMPNAME")
                                        {
                                            wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellFormula("COUNTIF(D2:D" + vLinesNo.ToString() + ", \"*\")");
                                            wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Right;
                                        }
                                        else if (vHeaderText.IndexOf("LEVEL_") >= 0)
                                        {
                                            int vColumnsInt = ColumnsIndex + 65;
                                            string vColumnsChar = ((char)vColumnsInt).ToString();
                                            wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellFormula("COUNTIF(" + vColumnsChar + "2:" + vColumnsChar + vLinesNo.ToString() + ", \"V\")");
                                            wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Right;
                                        }
                                    }
                                }
                                vRiskDepNo = dtExcel.Rows[i]["DepNo"].ToString().Trim();
                                vLinesNo = 0; //列數歸零
                                              //依單位分別寫進不同的工作表
                                vGroupText = dtExcel.Rows[i]["DepNo"].ToString().Trim() + "_" + dtExcel.Rows[i]["DepName"].ToString().Trim();
                                //建立新的工作表
                                wsExcel = (HSSFSheet)wbExcel.CreateSheet(vGroupText);
                                //建立標題列
                                wsExcel.CreateRow(vLinesNo);
                                for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                                {
                                    vHeaderText = (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "DEPNO") ? "" :
                                                  (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "DEPNAME") ? "單位" :
                                                  (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "DRIVER") ? "工號" :
                                                  (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "EMPNAME") ? "姓名" :
                                                  (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "MONTHTOTALHOURS") ? "月工時" :
                                                  (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "LEVEL_01") ? "0-222" :
                                                  (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "LEVEL_02") ? "223-300" :
                                                  (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "LEVEL_03") ? "301-350" :
                                                  (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "LEVEL_04") ? "351-400" :
                                                  (dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim() == "LEVEL_05") ? "401以上" : "";
                                    wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue(vHeaderText);
                                    wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csTitle;
                                }
                                vLinesNo++;
                            }
                            wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                            wsExcel.CreateRow(vLinesNo);
                            for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                            {
                                if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper() == "MONTHTOTALHOURS")
                                {
                                    vCellData = (Int32.Parse(dtExcel.Rows[i][ColumnsIndex].ToString().Trim()) / 60).ToString() + ":" +
                                                (Int32.Parse(dtExcel.Rows[i][ColumnsIndex].ToString().Trim()) - (60 * (Int32.Parse(dtExcel.Rows[i][ColumnsIndex].ToString().Trim()) / 60))).ToString("D2");
                                    wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue(vCellData);
                                    //wsExcel.AutoSizeColumn(ColumnsIndex);
                                    wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Right;
                                }
                                else if (dtExcel.Columns[ColumnsIndex].ColumnName.Trim().ToUpper().IndexOf("LEVEL_") >= 0)
                                {
                                    wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue(dtExcel.Rows[i][ColumnsIndex].ToString().Trim());
                                    //wsExcel.AutoSizeColumn(ColumnsIndex);
                                    wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csTitle;
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue(dtExcel.Rows[i][ColumnsIndex].ToString().Trim());
                                    //wsExcel.AutoSizeColumn(ColumnsIndex);
                                    wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData;
                                }
                            }
                            vLinesNo++;
                        }
                        //在最後一個工作表加上小計
                        wsExcel = (HSSFSheet)wbExcel.GetSheet(vGroupText);
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int ColumnsIndex = 0; ColumnsIndex < dtExcel.Columns.Count; ColumnsIndex++)
                        {
                            vHeaderText = dtExcel.Columns[ColumnsIndex].ColumnName.ToUpper().Trim();
                            if (vHeaderText == "MONTHTOTALHOURS")
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue("");
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData;
                            }
                            else if (vHeaderText == "DEPNO")
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellValue("合計：");
                                wsExcel.AddMergedRegion(new CellRangeAddress(vLinesNo, vLinesNo, 0, 2));
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Right;
                            }
                            else if (vHeaderText == "EMPNAME")
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellFormula("COUNTIF(D2:D" + vLinesNo.ToString() + ", \"*\")");
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Right;
                            }
                            else if (vHeaderText.IndexOf("LEVEL_") >= 0)
                            {
                                int vColumnsInt = ColumnsIndex + 65;
                                string vColumnsChar = ((char)vColumnsInt).ToString();
                                wsExcel.GetRow(vLinesNo).CreateCell(ColumnsIndex).SetCellFormula("COUNTIF(" + vColumnsChar + "2:" + vColumnsChar + vLinesNo.ToString() + ", \"V\")");
                                wsExcel.GetRow(vLinesNo).GetCell(ColumnsIndex).CellStyle = csData_Right;
                            }
                        }
                        try
                        {
                            /*
                            MemoryStream msTarget = new MemoryStream();
                            wbExcel.Write(msTarget);
                            // 設定強制下載標頭
                            Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
                            // 輸出檔案
                            Response.BinaryWrite(msTarget.ToArray());
                            msTarget.Close();

                            Response.End(); //*/
                            var msTarget = new NPOIMemoryStream();
                            msTarget.AllowClose = false;
                            wbExcel.Write(msTarget);
                            msTarget.Flush();
                            msTarget.Seek(0, SeekOrigin.Begin);
                            msTarget.AllowClose = true;

                            if (msTarget.Length > 0)
                            {

                                string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                                         ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                                         ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                                string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                                string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                     "DriverWorkHourLevelByMan.aspx" + Environment.NewLine +
                                                     "行車日期：" + eDriveYear_Search.Text.Trim() + "年" + eDriveMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                     "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                     "駕駛員：" + vRecordDriverStr;
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
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('找不到符合條件的資料！')");
                        Response.Write("</" + "Script>");
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('行車日期年月不可空白！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCLoseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plSearch.Visible = true;
        }
    }
}