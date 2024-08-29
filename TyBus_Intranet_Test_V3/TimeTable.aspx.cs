using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class TimeTable : System.Web.UI.Page
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
                    string vBuDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDate_S.ClientID;
                    string vBuDateScript = "window.open('" + vBuDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_S.Attributes["onClick"] = vBuDateScript;

                    vBuDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDate_E.ClientID;
                    vBuDateScript = "window.open('" + vBuDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_E.Attributes["onClick"] = vBuDateScript;
                    if (!IsPostBack)
                    {
                        if ((Int32.Parse(vLoginDepNo) > 10) || (Int32.Parse(vLoginDepNo) == 7) || (Int32.Parse(vLoginDepNo) == 8))
                        {
                            eBuDate_S.Text = DateTime.Today.ToShortDateString();
                            eBuDate_E.Text = "";
                            eBuDate_E.Visible = false;
                            lbSplit_3.Visible = false;
                        }
                        else if (vLoginDepNo == "06")
                        {
                            eBuDate_S.Text = DateTime.Today.ToShortDateString();
                            eBuDate_E.Text = "";
                            eBuDate_E.Visible = true;
                            lbSplit_3.Visible = true;
                        }
                        else
                        {
                            eBuDate_S.Text = DateTime.Today.AddMonths(-1).ToShortDateString();
                            eBuDate_E.Text = DateTime.Today.ToShortDateString();
                            eBuDate_E.Visible = true;
                            lbSplit_3.Visible = true;
                        }
                        plShowData.Visible = false;
                        plReport.Visible = false;
                        bbReport.Enabled = false;
                        bbExcel.Enabled = false;
                    }
                    else
                    {
                        TimeTableDataBind();
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

        private string GetSelStr()
        {
            DateTime vBuDate_S = (eBuDate_S.Text.Trim() != "") ? DateTime.Parse(eBuDate_S.Text.Trim()) : DateTime.Today;
            DateTime vBuDate_E = (eBuDate_E.Text.Trim() != "") ? DateTime.Parse(eBuDate_E.Text.Trim()) : DateTime.Today.AddMonths(1);
            string vWStr_BuDate = ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() != "")) ? "   and a.BuDate between '" + vBuDate_S.Year.ToString("D4") + "/" + vBuDate_S.ToString("MM/dd") + "' and'" + vBuDate_E.Year.ToString("D4") + "/" + vBuDate_E.ToString("MM/dd") + "' " + Environment.NewLine :
                                  ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() == "")) ? "   and a.BuDate = '" + vBuDate_S.Year.ToString("D4") + "/" + vBuDate_S.ToString("MM/dd") + "' " + Environment.NewLine :
                                  ((eBuDate_S.Text.Trim() == "") && (eBuDate_E.Text.Trim() != "")) ? "   and a.BuDate = '" + vBuDate_E.Year.ToString("D4") + "/" + vBuDate_E.ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and a.DepNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.DepNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_LinesNo = ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() != "")) ? "   and b.LinesNo between '" + eLinesNo_S.Text.Trim() + "' and '" + eLinesNo_E.Text.Trim() + "' " + Environment.NewLine :
                                   ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() == "")) ? "   and b.LinesNo = '" + eLinesNo_S.Text.Trim() + "' " + Environment.NewLine :
                                   ((eLinesNo_S.Text.Trim() == "") && (eLinesNo_E.Text.Trim() != "")) ? "   and b.LinesNo = '" + eLinesNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vSelectStr = "select a.BuDate, a.DepNo, (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                                "       b.LinesNo, (select LineName from Lines where LinesNo = b.LinesNo) LinesName, b.ToTime, b.ToLine, b.BackTime, b.BackLine, " + Environment.NewLine +
                                "       b.Car_ID, (select CLassTxt from DBDICB where FKey = '行車記錄單b     runsheetb       CARTYPE' and ClassNo = b.CarType) CarType_C, " + Environment.NewLine +
                                "       a.Driver, e.[Name] DriverName, e.Cellphone CellPhone, isnull(b.ActualKM, 0) ActualKM " + Environment.NewLine +
                                "  from RunSheetA a left join RunSheetB b on b.AssignNo = a.ASSIGNNO left join Employee e on e.EmpNo = a.Driver" + Environment.NewLine +
                                " where isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo + vWStr_LinesNo +
                                " order by a.BuDate, b.LinesNo, b.ToTime";
            return vSelectStr;
        }

        private string GetSelStr_Excel()
        {
            DateTime vBuDate_S = (eBuDate_S.Text.Trim() != "") ? DateTime.Parse(eBuDate_S.Text.Trim()) : DateTime.Today;
            DateTime vBuDate_E = (eBuDate_E.Text.Trim() != "") ? DateTime.Parse(eBuDate_E.Text.Trim()) : DateTime.Today.AddMonths(1);
            string vWStr_BuDate = ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() != "")) ? "   and a.BuDate between '" + vBuDate_S.Year.ToString("D4") + "/" + vBuDate_S.ToString("MM/dd") + "' and'" + vBuDate_E.Year.ToString("D4") + "/" + vBuDate_E.ToString("MM/dd") + "' " + Environment.NewLine :
                                  ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() == "")) ? "   and a.BuDate = '" + vBuDate_S.Year.ToString("D4") + "/" + vBuDate_S.ToString("MM/dd") + "' " + Environment.NewLine :
                                  ((eBuDate_S.Text.Trim() == "") && (eBuDate_E.Text.Trim() != "")) ? "   and a.BuDate = '" + vBuDate_E.Year.ToString("D4") + "/" + vBuDate_E.ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and a.DepNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.DepNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_LinesNo = ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() != "")) ? "   and b.LinesNo between '" + eLinesNo_S.Text.Trim() + "' and '" + eLinesNo_E.Text.Trim() + "' " + Environment.NewLine :
                                   ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() == "")) ? "   and b.LinesNo = '" + eLinesNo_S.Text.Trim() + "' " + Environment.NewLine :
                                   ((eLinesNo_S.Text.Trim() == "") && (eLinesNo_E.Text.Trim() != "")) ? "   and b.LinesNo = '" + eLinesNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vSelectStr = "select a.BuDate, a.DepNo, (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                                "       b.LinesNo, (select LineName from Lines where LinesNo = b.LinesNo) LinesName, b.ToTime, b.ToLine, b.BackTime, b.BackLine, " + Environment.NewLine +
                                "       b.Car_ID, (select CLassTxt from DBDICB where FKey = '行車記錄單b     runsheetb       CARTYPE' and ClassNo = b.CarType) CarType_C, " + Environment.NewLine +
                                "       a.Driver, e.[Name] DriverName, e.Cellphone CellPhone, isnull(b.ActualKM, 0) ActualKM " + Environment.NewLine +
                                "  from RunSheetA a left join RunSheetB b on b.AssignNo = a.ASSIGNNO left join Employee e on e.EmpNo = a.Driver" + Environment.NewLine +
                                " where isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_DepNo + vWStr_LinesNo +
                                " order by b.LinesNo, a.BuDate, b.ToTime";
            return vSelectStr;
        }

        private int GetRecordCount()
        {
            DateTime vBuDate_S = (eBuDate_S.Text.Trim() != "") ? DateTime.Parse(eBuDate_S.Text.Trim()) : DateTime.Today;
            DateTime vBuDate_E = (eBuDate_E.Text.Trim() != "") ? DateTime.Parse(eBuDate_E.Text.Trim()) : DateTime.Today.AddMonths(1);
            string vWStr_BuDate = ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() != "")) ? "   and a.BuDate between '" + vBuDate_S.Year.ToString("D4") + "/" + vBuDate_S.ToString("MM/dd") + "' and'" + vBuDate_E.Year.ToString("D4") + "/" + vBuDate_E.ToString("MM/dd") + "' " + Environment.NewLine :
                                  ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() == "")) ? "   and a.BuDate = '" + vBuDate_S.Year.ToString("D4") + "/" + vBuDate_S.ToString("MM/dd") + "' " + Environment.NewLine :
                                  ((eBuDate_S.Text.Trim() == "") && (eBuDate_E.Text.Trim() != "")) ? "   and a.BuDate = '" + vBuDate_E.Year.ToString("D4") + "/" + vBuDate_E.ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and a.DepNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and a.DepNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_LinesNo = ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() != "")) ? "   and b.LinesNo between '" + eLinesNo_S.Text.Trim() + "' and '" + eLinesNo_E.Text.Trim() + "' " + Environment.NewLine :
                                   ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() == "")) ? "   and b.LinesNo = '" + eLinesNo_S.Text.Trim() + "' " + Environment.NewLine :
                                   ((eLinesNo_S.Text.Trim() == "") && (eLinesNo_E.Text.Trim() != "")) ? "   and b.LinesNo = '" + eLinesNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vTempStr = "select count(b.AssignNo) RCount " + Environment.NewLine +
                              "  from RunSheetA a left join RunSheetB b on b.AssignNo = a.ASSIGNNO " + Environment.NewLine +
                              " where isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                              vWStr_BuDate + vWStr_DepNo + vWStr_LinesNo;
            int vResultCount = Int32.Parse(PF.GetValue(vConnStr, vTempStr, "RCount"));
            return vResultCount;
        }

        private void TimeTableDataBind()
        {
            string vSelStr = GetSelStr();
            sdsTimeTable.SelectCommand = "";
            sdsTimeTable.SelectCommand = vSelStr;
            gridTimeTable.DataBind();
            if (gridTimeTable.Rows.Count > 0)
            {
                plShowData.Visible = true;
                bbReport.Enabled = true;
                bbExcel.Enabled = true;
            }
            else
            {
                plShowData.Visible = false;
                bbReport.Enabled = false;
                bbExcel.Enabled = false;
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            TimeTableDataBind();
        }

        /// <summary>
        /// Excel 匯出資料依路線編號分頁
        /// </summary>
        private void ExportExcel_PagedByLinesNo()
        {
            DateTime vBuDate;
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

            DateTime vBuDate_S = (eBuDate_S.Text.Trim() != "") ? DateTime.Parse(eBuDate_S.Text.Trim()) : DateTime.Today;
            DateTime vBuDate_E = (eBuDate_E.Text.Trim() != "") ? DateTime.Parse(eBuDate_E.Text.Trim()) : DateTime.Today.AddMonths(1);
            string vHeaderText = "";
            string vFileName = ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() != "")) ?
                               (vBuDate_S.Year - 1911).ToString("D3") + vBuDate_S.Month.ToString("D2") + vBuDate_S.Day.ToString("D2") + "至" +
                               (vBuDate_E.Year - 1911).ToString("D3") + vBuDate_E.Month.ToString("D2") + vBuDate_E.Day.ToString("D2") + "路線時刻表" :
                               ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() == "")) ?
                               (vBuDate_S.Year - 1911).ToString("D3") + vBuDate_S.Month.ToString("D2") + vBuDate_S.Day.ToString("D2") + "路線時刻表" :
                               ((eBuDate_S.Text.Trim() == "") && (eBuDate_E.Text.Trim() != "")) ?
                               (vBuDate_E.Year - 1911).ToString("D3") + vBuDate_E.Month.ToString("D2") + vBuDate_E.Day.ToString("D2") + "路線時刻表" : "";
            int vLinesNo = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelStr_Excel();

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    string vLinesNo_Old = "";
                    string vLinesNo_Current = "";
                    while (drExcel.Read())
                    {
                        //vLinesNo_Current = drExcel["LinesNo"].ToString().Trim();
                        vLinesNo_Current = drExcel[3].ToString().Trim();
                        if (vLinesNo_Current != "")
                        {
                            if (vLinesNo_Current != vLinesNo_Old)
                            {
                                //路線編號有變
                                vLinesNo = 0;
                                vLinesNo_Old = vLinesNo_Current;
                                //新增一個工作表
                                wsExcel = (HSSFSheet)wbExcel.CreateSheet(vLinesNo_Old + "路線時刻表");
                                //寫入標題列
                                wsExcel.CreateRow(vLinesNo);
                                for (int i = 0; i < drExcel.FieldCount; i++)
                                {
                                    vHeaderText = (drExcel.GetName(i) == "BuDate") ? "日期" :
                                                  (drExcel.GetName(i) == "DepNo") ? "部門編號" :
                                                  (drExcel.GetName(i) == "DepName") ? "部門" :
                                                  (drExcel.GetName(i) == "LinesNo") ? "路線代號" :
                                                  (drExcel.GetName(i) == "LinesName") ? "路線名稱" :
                                                  (drExcel.GetName(i) == "ToTime") ? "出車時刻" :
                                                  (drExcel.GetName(i) == "ToLine") ? "去程路線" :
                                                  (drExcel.GetName(i) == "BackTime") ? "回程時刻" :
                                                  (drExcel.GetName(i) == "BackLine") ? "回程路線" :
                                                  (drExcel.GetName(i) == "Car_ID") ? "車號" :
                                                  (drExcel.GetName(i) == "CarType_C") ? "車種" :
                                                  (drExcel.GetName(i) == "Driver") ? "駕駛員工號" :
                                                  (drExcel.GetName(i) == "DriverName") ? "駕駛員姓名" :
                                                  (drExcel.GetName(i) == "CellPhone") ? "電話" :
                                                  (drExcel.GetName(i) == "ActualKM") ? "公里數" : drExcel.GetName(i);
                                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                                }
                            }
                            wsExcel = (HSSFSheet)wbExcel.GetSheet(vLinesNo_Old + "路線時刻表");
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            for (int i = 0; i < drExcel.FieldCount; i++)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(i);
                                if ((drExcel.GetName(i) == "ActualKM"))
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                }
                                else if ((drExcel.GetName(i) == "BuDate") && (drExcel[i].ToString() != ""))
                                {
                                    string vTempStr = drExcel[i].ToString();
                                    vBuDate = DateTime.Parse(drExcel[i].ToString());
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
                            string vBuDateStr = ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() != "")) ? "從 " + eBuDate_S.Text.Trim() + " 起至 " + eBuDate_E.Text.Trim() + " 止" :
                                                ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() == "")) ? eBuDate_S.Text.Trim() :
                                                ((eBuDate_S.Text.Trim() == "") && (eBuDate_E.Text.Trim() != "")) ? eBuDate_E.Text.Trim() : "不分日期";
                            string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "從 " + eDepNo_S.Text.Trim() + " 起至 " + eDepNo_E.Text.Trim() + " 止" :
                                               ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() :
                                               ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() : "全部";
                            string vLinesNOStr = ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() != "")) ? "從 " + eLinesNo_S.Text.Trim() + " 起至 " + eLinesNo_E.Text.Trim() + " 止" :
                                                 ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() == "")) ? eLinesNo_S.Text.Trim() :
                                                 ((eLinesNo_S.Text.Trim() == "") && (eLinesNo_E.Text.Trim() != "")) ? eLinesNo_E.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_路線時刻表" + Environment.NewLine +
                                                 "TimeTable.aspx" + Environment.NewLine +
                                                 "日期：" + vBuDateStr + Environment.NewLine +
                                                 "單位：" + vDepNoStr + Environment.NewLine +
                                                 "路線：" + vLinesNOStr;
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
                    Response.Write("alert('找不到符合條件的資料，請重新設定過濾條件後進行查詢！')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// Excel 匯出資料不分工作表
        /// </summary>
        private void ExportExcel_SinglePage()
        {
            DateTime vBuDate;
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

            DateTime vBuDate_S = (eBuDate_S.Text.Trim() != "") ? DateTime.Parse(eBuDate_S.Text.Trim()) : DateTime.Today;
            DateTime vBuDate_E = (eBuDate_E.Text.Trim() != "") ? DateTime.Parse(eBuDate_E.Text.Trim()) : DateTime.Today.AddMonths(1);
            string vHeaderText = "";
            string vFileName = ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() != "")) ?
                               (vBuDate_S.Year - 1911).ToString("D3") + vBuDate_S.Month.ToString("D2") + vBuDate_S.Day.ToString("D2") + "至" +
                               (vBuDate_E.Year - 1911).ToString("D3") + vBuDate_E.Month.ToString("D2") + vBuDate_E.Day.ToString("D2") + "路線時刻表" :
                               ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() == "")) ?
                               (vBuDate_S.Year - 1911).ToString("D3") + vBuDate_S.Month.ToString("D2") + vBuDate_S.Day.ToString("D2") + "路線時刻表" :
                               ((eBuDate_S.Text.Trim() == "") && (eBuDate_E.Text.Trim() != "")) ?
                               (vBuDate_E.Year - 1911).ToString("D3") + vBuDate_E.Month.ToString("D2") + vBuDate_E.Day.ToString("D2") + "路線時刻表" : "";
            int vLinesNo = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelStr_Excel();

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
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
                        vHeaderText = (drExcel.GetName(i) == "BuDate") ? "日期" :
                                      (drExcel.GetName(i) == "DepNo") ? "部門編號" :
                                      (drExcel.GetName(i) == "DepName") ? "部門" :
                                      (drExcel.GetName(i) == "LinesNo") ? "路線代號" :
                                      (drExcel.GetName(i) == "LinesName") ? "路線名稱" :
                                      (drExcel.GetName(i) == "ToTime") ? "出車時刻" :
                                      (drExcel.GetName(i) == "ToLine") ? "去程路線" :
                                      (drExcel.GetName(i) == "BackTime") ? "回程時刻" :
                                      (drExcel.GetName(i) == "BackLine") ? "回程路線" :
                                      (drExcel.GetName(i) == "Car_ID") ? "車號" :
                                      (drExcel.GetName(i) == "CarType_C") ? "車種" :
                                      (drExcel.GetName(i) == "Driver") ? "駕駛員工號" :
                                      (drExcel.GetName(i) == "DriverName") ? "駕駛員姓名" :
                                      (drExcel.GetName(i) == "CellPhone") ? "電話" :
                                      (drExcel.GetName(i) == "ActualKM") ? "公里數" : drExcel.GetName(i);
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
                            if ((drExcel.GetName(i) == "ActualKM"))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                            }
                            else if ((drExcel.GetName(i) == "BuDate") && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
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
                            string vBuDateStr = ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() != "")) ? "從 " + eBuDate_S.Text.Trim() + " 起至 " + eBuDate_E.Text.Trim() + " 止" :
                                                ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() == "")) ? eBuDate_S.Text.Trim() :
                                                ((eBuDate_S.Text.Trim() == "") && (eBuDate_E.Text.Trim() != "")) ? eBuDate_E.Text.Trim() : "不分日期";
                            string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "從 " + eDepNo_S.Text.Trim() + " 起至 " + eDepNo_E.Text.Trim() + " 止" :
                                               ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() :
                                               ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() : "全部";
                            string vLinesNOStr = ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() != "")) ? "從 " + eLinesNo_S.Text.Trim() + " 起至 " + eLinesNo_E.Text.Trim() + " 止" :
                                                 ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() == "")) ? eLinesNo_S.Text.Trim() :
                                                 ((eLinesNo_S.Text.Trim() == "") && (eLinesNo_E.Text.Trim() != "")) ? eLinesNo_E.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_路線時刻表" + Environment.NewLine +
                                                 "TimeTable.aspx" + Environment.NewLine +
                                                 "日期：" + vBuDateStr + Environment.NewLine +
                                                 "單位：" + vDepNoStr + Environment.NewLine +
                                                 "路線：" + vLinesNOStr;
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

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
        {
            int vRCount = GetRecordCount();
            if (vRCount <= 65530)
            {
                ExportExcel_SinglePage();
            }
            else
            {
                ExportExcel_PagedByLinesNo();
            }
        }

        /// <summary>
        /// 預覽報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbReport_Click(object sender, EventArgs e)
        {
            DateTime vPrintDate_S = (eBuDate_S.Text.Trim() != "") ? DateTime.Parse(eBuDate_S.Text.Trim()) : DateTime.Today;
            DateTime vPrintDate_E = (eBuDate_E.Text.Trim() != "") ? DateTime.Parse(eBuDate_E.Text.Trim()) : DateTime.Today.AddMonths(1);
            string vReportName = ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() != "")) ?
                                 (vPrintDate_S.Year - 1911).ToString("D3") + vPrintDate_S.Month.ToString("D2") + vPrintDate_S.Day.ToString("D2") + "至" +
                                 (vPrintDate_E.Year - 1911).ToString("D3") + vPrintDate_E.Month.ToString("D2") + vPrintDate_E.Day.ToString("D2") + "路線時刻表" :
                                 ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() == "")) ?
                                 (vPrintDate_S.Year - 1911).ToString("D3") + vPrintDate_S.Month.ToString("D2") + vPrintDate_S.Day.ToString("D2") + "路線時刻表" :
                                 ((eBuDate_S.Text.Trim() == "") && (eBuDate_E.Text.Trim() != "")) ?
                                 (vPrintDate_E.Year - 1911).ToString("D3") + vPrintDate_E.Month.ToString("D2") + vPrintDate_E.Day.ToString("D2") + "路線時刻表" : "";
            string vTempStr = "select [Name] from Custom where Types = 'O' and Code = 'A000' ";
            string vCompanyName = PF.GetValue(vConnStr, vTempStr, "Name");
            //統計報表
            string vSelectStr = GetSelStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);

                ReportDataSource rdsPrint = new ReportDataSource("TimeTableP", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\TimeTableP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.Refresh();
                plShowData.Visible = false;
                plReport.Visible = true;

                string vBuDateStr = ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() != "")) ? "從 " + eBuDate_S.Text.Trim() + " 起至 " + eBuDate_E.Text.Trim() + " 止" :
                                    ((eBuDate_S.Text.Trim() != "") && (eBuDate_E.Text.Trim() == "")) ? eBuDate_S.Text.Trim() :
                                    ((eBuDate_S.Text.Trim() == "") && (eBuDate_E.Text.Trim() != "")) ? eBuDate_E.Text.Trim() : "不分日期";
                string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "從 " + eDepNo_S.Text.Trim() + " 起至 " + eDepNo_E.Text.Trim() + " 止" :
                                   ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() :
                                   ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() : "全部";
                string vLinesNOStr = ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() != "")) ? "從 " + eLinesNo_S.Text.Trim() + " 起至 " + eLinesNo_E.Text.Trim() + " 止" :
                                     ((eLinesNo_S.Text.Trim() != "") && (eLinesNo_E.Text.Trim() == "")) ? eLinesNo_S.Text.Trim() :
                                     ((eLinesNo_S.Text.Trim() == "") && (eLinesNo_E.Text.Trim() != "")) ? eLinesNo_E.Text.Trim() : "全部";
                string vRecordNote = "預覽報表_路線時刻表" + Environment.NewLine +
                                     "TimeTable.aspx" + Environment.NewLine +
                                     "日期：" + vBuDateStr + Environment.NewLine +
                                     "單位：" + vDepNoStr + Environment.NewLine +
                                     "路線：" + vLinesNOStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
        }
    }
}