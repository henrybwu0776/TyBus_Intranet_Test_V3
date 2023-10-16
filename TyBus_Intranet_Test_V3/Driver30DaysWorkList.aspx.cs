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
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class Driver30DaysWorkList : System.Web.UI.Page
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
        DateTime vToday = DateTime.Today;

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
                        plShowData.Visible = true;
                        eCalDays.Enabled = true;
                        if (vLoginDepNo != "09")
                        {
                            eCalDays.Enabled = false;
                            lbDataNote.Text = "";
                        }
                        eStopDate.Text = vToday.ToShortDateString();
                        eCalDays.Text = "30"; //回算天數預設30天
                        eLastDays.Text = "7"; //最近天數預設7天
                        if (Int32.Parse(vLoginDepNo) > 10)
                        {
                            sdsDepNoList.SelectCommand = "SELECT e.DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DepNo)) AS DepName " + Environment.NewLine +
                                                         "  FROM EmployeeDepNo e " + Environment.NewLine +
                                                         " WHERE e.EmpNo = '" + vLoginID + "' " + Environment.NewLine +
                                                         "   AND isnull(e.IsUsed, 0) = 1 " + Environment.NewLine +
                                                         " ORDER BY e.IsDefault DESC, e.DepNo";
                            eDepNo.Text = PF.GetValue(vConnStr, "select DepNo from EmployeeDepNo where EmpNo = '" + vLoginID + "' and IsDefault = 1", "DepNo");
                        }
                        else
                        {
                            sdsDepNoList.SelectCommand = "select cast('' as varchar) DepNo, cast('' as varchar) DepName " + Environment.NewLine +
                                                         " union all " + Environment.NewLine +
                                                         "select DepNo, [Name] as DepName " + Environment.NewLine +
                                                         "  from Department " + Environment.NewLine +
                                                         " where InSHReport = 'V'" + Environment.NewLine +
                                                         " order by DepNo ";
                            eDepNo.Text = "";
                        }
                        sdsDepNoList.DataBind();
                        ddlDepNo.SelectedIndex = 0;
                    }

                    string vStopDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eStopDate.ClientID;
                    string vStopDateScript = "window.open('" + vStopDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eStopDate.Attributes["onClick"] = vStopDateScript;
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
        /// 取回資料查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelStr()
        {
            string vSelectStr = "";
            string vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "   and DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver.Text.Trim() != "") ? "   and EmpNo = '" + eDriver.Text.Trim() + "' " + Environment.NewLine : "";
            DateTime vStopDate = DateTime.Parse(eStopDate.Text.Trim());
            DateTime vCurrDate;
            string vStopDate_Str = vStopDate.Year.ToString("D4") + "/" + vStopDate.ToString("MM/dd");
            string vCurrDate_Str = "";
            int vCalDays = (eCalDays.Text.Trim() != "") ? Int32.Parse(eCalDays.Text.Trim()) : 30;

            vSelectStr = "select DepNo 站別代號, (select [Name] from Department where DepNo = e.DepNo) 站別, EmpNo 駕駛員工號, [Name] 駕駛員, " + Environment.NewLine +
                         "       (select count(Driver) from RunSheetA a where a.Driver = e.EmpNo and a.BuDate between DateAdd(day, -5, '" + vStopDate_Str + "') and '" + vStopDate_Str + "') 最近六天出勤天數, " + Environment.NewLine +
                         "       cast(0 as int) 連續未休出勤天數, " + Environment.NewLine;
            for (int i = vCalDays; i > 0; i--)
            {
                vCurrDate = vStopDate.AddDays(i * -1);
                vCurrDate_Str = vCurrDate.Year.ToString("D4") + "/" + vCurrDate.ToString("MM/dd");
                vSelectStr = vSelectStr + "       case when isnull((select Driver from RunSheetA a where a.Driver = e.EmpNo and a.BuDate = '" + vCurrDate_Str + "'), '') = '' then '' else 'O' end as '" + vCurrDate.ToString("MM/dd") + "'," + Environment.NewLine;
            }
            vSelectStr = vSelectStr +
                         "       case when isnull((select Driver from RunSheetA a where a.Driver = e.EmpNo and a.BuDate = '" + vStopDate_Str + "'), '') = '' then '' else 'O' end as '" + vStopDate.ToString("MM/dd") + "' " + Environment.NewLine +
                         "  from Employee e " + Environment.NewLine +
                         " where [Type] = '20' " + Environment.NewLine + vWStr_DepNo + vWStr_Driver +
                         "   and (LeaveDay is null or LeaveDay > '" + vStopDate_Str + "') " + Environment.NewLine +
                         "   and e.DepNo in (select DepNo from Department where InSHReport = 'V') " + Environment.NewLine +
                         " order by e.DepNo, e.EmpNo";

            return vSelectStr;
        }

        /// <summary>
        /// 取回列印用資料查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetPrintSelStr()
        {
            string vSelectStr = "";
            string vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "   and DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver.Text.Trim() != "") ? "   and EmpNo = '" + eDriver.Text.Trim() + "' " + Environment.NewLine : "";
            DateTime vStopDate = DateTime.Parse(eStopDate.Text.Trim());
            DateTime vCurrDate;
            string vStopDate_Str = vStopDate.Year.ToString("D4") + "/" + vStopDate.ToString("MM/dd");
            string vCurrDate_Str = "";
            int vCalDays = (eCalDays.Text.Trim() != "") ? Int32.Parse(eCalDays.Text.Trim()) : 30;

            vSelectStr = "select DepNo 站別代號, (select [Name] from Department where DepNo = e.DepNo) 站別, EmpNo 駕駛員工號, [Name] 駕駛員, " + Environment.NewLine +
                         "       (select count(Driver) from RunSheetA a where a.Driver = e.EmpNo and a.BuDate between DateAdd(day, -5, '" + vStopDate_Str + "') and '" + vStopDate_Str + "') 最近六天出勤天數, " + Environment.NewLine +
                         "       cast(0 as int) 連續天數, " + Environment.NewLine;
            for (int i = vCalDays; i > 0; i--)
            {
                vCurrDate = vStopDate.AddDays(i * -1);
                vCurrDate_Str = vCurrDate.Year.ToString("D4") + "/" + vCurrDate.ToString("MM/dd");
                vSelectStr = vSelectStr + "       case when isnull((select Driver from RunSheetA a where a.Driver = e.EmpNo and a.BuDate = '" + vCurrDate_Str + "'), '') = '' then '' else 'O' end as '" + (vCalDays - i + 1).ToString("D2") + "'," + Environment.NewLine;
            }
            vSelectStr = vSelectStr +
                         "       case when isnull((select Driver from RunSheetA a where a.Driver = e.EmpNo and a.BuDate = '" + vStopDate_Str + "'), '') = '' then '' else 'O' end as '" + (vCalDays + 1).ToString("D2") + "' " + Environment.NewLine +
                         "  from Employee e " + Environment.NewLine +
                         " where [Type] = '20' " + Environment.NewLine + vWStr_DepNo + vWStr_Driver +
                         "   and (LeaveDay is null or LeaveDay > '" + vStopDate_Str + "') " + Environment.NewLine +
                         "   and e.DepNo in (select DepNo from Department where InSHReport = 'V') " + Environment.NewLine +
                         " order by e.DepNo, e.EmpNo";

            return vSelectStr;
        }

        /// <summary>
        /// 超時人員匯出 EXCEL 用字串
        /// </summary>
        /// <returns></returns>
        private string GetOverTimeSelStr()
        {
            string vSelectStr = "";
            string vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "   and DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver.Text.Trim() != "") ? "   and EmpNo = '" + eDriver.Text.Trim() + "' " + Environment.NewLine : "";
            DateTime vStopDate = DateTime.Parse(eStopDate.Text.Trim());
            DateTime vCurrDate;
            string vStopDate_Str = vStopDate.Year.ToString("D4") + "/" + vStopDate.ToString("MM/dd");
            string vCurrDate_Str = "";
            int vCalDays = (eCalDays.Text.Trim() != "") ? Int32.Parse(eCalDays.Text.Trim()) : 30;
            int vLastDays = (eLastDays.Text.Trim() != "") ? Int32.Parse(eLastDays.Text.Trim()) : 7;
            string vBackDay = ((vLastDays - 1) * -1).ToString();

            vSelectStr = "select DepNo 站別代號, (select [Name] from Department where DepNo = e.DepNo) 站別, EmpNo 駕駛員工號, [Name] 駕駛員, " + Environment.NewLine +
                         "       (select count(Driver) from RunSheetA a where a.Driver = e.EmpNo and a.BuDate between DateAdd(day, " + vBackDay + ", '" + vStopDate_Str + "') and '" + vStopDate_Str + "') [最近 " + vLastDays + " 天出勤天數], " + Environment.NewLine +
                         "       cast(0 as int) 連續未休出勤天數, " + Environment.NewLine;
            for (int i = vCalDays; i > 0; i--)
            {
                vCurrDate = vStopDate.AddDays(i * -1);
                vCurrDate_Str = vCurrDate.Year.ToString("D4") + "/" + vCurrDate.ToString("MM/dd");
                vSelectStr = vSelectStr + "       case when isnull((select Driver from RunSheetA a where a.Driver = e.EmpNo and a.BuDate = '" + vCurrDate_Str + "'), '') = '' then '' else 'O' end as '" + vCurrDate.ToString("MM/dd") + "'," + Environment.NewLine;
            }
            vSelectStr = vSelectStr +
                         "       case when isnull((select Driver from RunSheetA a where a.Driver = e.EmpNo and a.BuDate = '" + vStopDate_Str + "'), '') = '' then '' else 'O' end as '" + vStopDate.ToString("MM/dd") + "' " + Environment.NewLine +
                         "  from Employee e " + Environment.NewLine +
                         " where [Type] = '20' " + Environment.NewLine + vWStr_DepNo + vWStr_Driver +
                         "   and (LeaveDay is null or LeaveDay > '" + vStopDate_Str + "') " + Environment.NewLine +
                         "   and e.DepNo in (select DepNo from Department where InSHReport = 'V') " + Environment.NewLine +
                         " order by e.DepNo, e.EmpNo";

            return vSelectStr;
        }

        /// <summary>
        /// 超時人員匯出EXCEL查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetPrintOverTimeSelStr()
        {
            string vSelectStr = "";
            string vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "   and DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver.Text.Trim() != "") ? "   and EmpNo = '" + eDriver.Text.Trim() + "' " + Environment.NewLine : "";
            DateTime vStopDate = DateTime.Parse(eStopDate.Text.Trim());
            DateTime vCurrDate;
            string vStopDate_Str = vStopDate.Year.ToString("D4") + "/" + vStopDate.ToString("MM/dd");
            string vCurrDate_Str = "";
            int vCalDays = (eCalDays.Text.Trim() != "") ? Int32.Parse(eCalDays.Text.Trim()) : 30;
            int vLastDays = (eLastDays.Text.Trim() != "") ? Int32.Parse(eLastDays.Text.Trim()) : 7;
            string vBackDay = ((vLastDays - 1) * -1).ToString();

            vSelectStr = "select DepNo 站別代號, (select [Name] from Department where DepNo = e.DepNo) 站別, EmpNo 駕駛員工號, [Name] 駕駛員, " + Environment.NewLine +
                         "       (select count(Driver) from RunSheetA a where a.Driver = e.EmpNo and a.BuDate between DateAdd(day, " + vBackDay + ", '" + vStopDate_Str + "') and '" + vStopDate_Str + "') [最近 " + vLastDays + " 天出勤天數], " + Environment.NewLine +
                         "       cast(0 as int) 連續天數, " + Environment.NewLine;
            for (int i = vCalDays; i > 0; i--)
            {
                vCurrDate = vStopDate.AddDays(i * -1);
                vCurrDate_Str = vCurrDate.Year.ToString("D4") + "/" + vCurrDate.ToString("MM/dd");
                vSelectStr = vSelectStr + "       case when isnull((select Driver from RunSheetA a where a.Driver = e.EmpNo and a.BuDate = '" + vCurrDate_Str + "'), '') = '' then '' else 'O' end as '" + (vCalDays - i + 1).ToString("D2") + "'," + Environment.NewLine;
            }
            vSelectStr = vSelectStr +
                         "       case when isnull((select Driver from RunSheetA a where a.Driver = e.EmpNo and a.BuDate = '" + vStopDate_Str + "'), '') = '' then '' else 'O' end as '" + (vCalDays + 1).ToString("D2") + "' " + Environment.NewLine +
                         "  from Employee e " + Environment.NewLine +
                         " where [Type] = '20' " + Environment.NewLine + vWStr_DepNo + vWStr_Driver +
                         "   and (LeaveDay is null or LeaveDay > '" + vStopDate_Str + "') " + Environment.NewLine +
                         "   and e.DepNo in (select DepNo from Department where InSHReport = 'V') " + Environment.NewLine +
                         " order by e.DepNo, e.EmpNo";

            return vSelectStr;
        }

        protected void ddlDepNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDepNo.Text = ddlDepNo.SelectedValue.Trim();
        }

        protected void gridShowData_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            DataRowView drvTemp = (DataRowView)e.Row.DataItem;

            if ((e.Row.RowType == DataControlRowType.DataRow) && (drvTemp != null))
            {
                if (Int32.Parse(drvTemp[4].ToString().Trim()) >= 6)
                {
                    e.Row.Attributes.Add("style", "background-color: #FFC1C1");
                }
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            string vSelStr = GetSelStr();
            int vMaxDays = 0;
            int vDaysCount = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelStr, connTemp);
                connTemp.Open();
                DataTable dtTemp = new DataTable("DriverWorkList");
                SqlDataAdapter daTemp = new SqlDataAdapter(cmdTemp);
                daTemp.Fill(dtTemp);
                for (int i = 0; i < dtTemp.Rows.Count; i++)
                {
                    vMaxDays = 0;
                    vDaysCount = 0;
                    string vEmpNo_DataTable = dtTemp.Rows[i][2].ToString();
                    for (int J = 6; J < dtTemp.Columns.Count; J++)
                    {
                        if (dtTemp.Rows[i][J].ToString().Trim() == "O")
                        {
                            vDaysCount = vDaysCount + 1;
                        }
                        else
                        {
                            if (vMaxDays < vDaysCount)
                            {
                                vMaxDays = vDaysCount;
                            }
                            vDaysCount = 0;
                        }
                    }
                    if (vMaxDays < vDaysCount)
                    {
                        vMaxDays = vDaysCount;
                    }
                    vDaysCount = 0;
                    dtTemp.Rows[i][5] = vMaxDays.ToString();
                }
                gridShowData.DataSource = dtTemp;
                gridShowData.DataBind();
                lbDataNote.Text = "當天打 O 表示有出勤 (以行車憑單為準) ，至 " + eStopDate.Text.Trim() + " 為止已經連續上班至少 6 天人員以紅底表示。";
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelStr();
            int vMaxDays = 0;
            int vDaysCount = 0;
            Boolean vIsRed = false;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                DataTable dtExcel = new DataTable();
                SqlDataAdapter daExcel = new SqlDataAdapter(cmdExcel);
                daExcel.Fill(dtExcel);
                if (dtExcel.Rows.Count > 0)
                {
                    //查詢結果有資料的時候才執行
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vMaxDays = 0;
                        vDaysCount = 0;
                        string vEmpNo_DataTable = dtExcel.Rows[i][2].ToString();
                        for (int J = 6; J < dtExcel.Columns.Count; J++)
                        {
                            if (dtExcel.Rows[i][J].ToString().Trim() == "O")
                            {
                                vDaysCount = vDaysCount + 1;
                            }
                            else
                            {
                                if (vMaxDays < vDaysCount)
                                {
                                    vMaxDays = vDaysCount;
                                }
                                vDaysCount = 0;
                            }
                        }
                        if (vMaxDays < vDaysCount)
                        {
                            vMaxDays = vDaysCount;
                        }
                        vDaysCount = 0;
                        dtExcel.Rows[i][5] = vMaxDays.ToString();
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
                    fontData.FontHeightInPoints = 12; //字體大小
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
                    string vFileName = "駕駛員最近30天出勤狀況";
                    int vLinesNo = 0;

                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < dtExcel.Columns.Count; i++)
                    {
                        vHeaderText = dtExcel.Columns[i].ColumnName.Trim();
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vLinesNo++;
                        vIsRed = (Int32.Parse(dtExcel.Rows[i][4].ToString().Trim()) >= 6);
                        wsExcel.CreateRow(vLinesNo);
                        for (int j = 0; j < dtExcel.Columns.Count; j++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(j);
                            if (j == 4)
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(double.Parse(dtExcel.Rows[i][j].ToString().Trim()));
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(dtExcel.Rows[i][j].ToString().Trim());
                            }
                            wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = (vIsRed) ? csData_Red : csData;
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
                            string vDepNoStr = (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : "全部";
                            string vDriverStr = (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : "全部";
                            string vCalDayStr = (eCalDays.Text.Trim() != "") ? eCalDays.Text.Trim() : "30天";
                            string vLastDayStr = (eLastDays.Text.Trim() != "") ? eLastDays.Text.Trim() : "7天";
                            string vStopDateStr = (eStopDate.Text.Trim() != "") ? eStopDate.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_駕駛員最近30天出勤狀況" + Environment.NewLine +
                                                 "CustomServiceDriverList.aspx" + Environment.NewLine +
                                                 "站別：" + vDepNoStr + Environment.NewLine +
                                                 "駕駛員編號：" + vDriverStr + Environment.NewLine +
                                                 "回算天數：" + vCalDayStr + Environment.NewLine +
                                                 "最近天數：" + vLastDayStr + Environment.NewLine +
                                                 "截止日期：" + vStopDateStr;
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

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void eDriver_TextChanged(object sender, EventArgs e)
        {
            string vDriver = eDriver.Text.Trim();
            string vSQLStr = "";
            string vDriverName = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select [Name] from Employee where EmpNo = '" + vDriver + "' ";
            vDriverName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriver;
                vSQLStr = "select top 1 EmpNo from EMployee where [Name] = '" + vDriverName + "' ";
                vDriver = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eDriver.Text = vDriver;
            lbDriverName.Text = vDriverName;
            if ((vDriver == "") && (vDriverName == ""))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('查無駕駛員！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 超時人員列表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_OverTime_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetPrintOverTimeSelStr();
            int vMaxDays = 0;
            int vDaysCount = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                SqlCommand cmdPrint = new SqlCommand(vSelectStr, connPrint);
                connPrint.Open();
                SqlDataAdapter daPrint = new SqlDataAdapter(cmdPrint);
                DataTable dtTemp = new DataTable();
                daPrint.Fill(dtTemp);
                if (dtTemp.Rows.Count > 0)
                {
                    //查詢結果有資料的時候才執行
                    for (int i = 0; i < dtTemp.Rows.Count; i++)
                    {
                        vMaxDays = 0;
                        vDaysCount = 0;
                        string vEmpNo_DataTable = dtTemp.Rows[i][2].ToString();
                        for (int J = 6; J < dtTemp.Columns.Count; J++)
                        {
                            if (dtTemp.Rows[i][J].ToString().Trim() == "O")
                            {
                                vDaysCount = vDaysCount + 1;
                            }
                            else
                            {
                                if (vMaxDays < vDaysCount)
                                {
                                    vMaxDays = vDaysCount;
                                }
                                vDaysCount = 0;
                            }
                        }
                        if (vMaxDays < vDaysCount)
                        {
                            vMaxDays = vDaysCount;
                        }
                        vDaysCount = 0;
                        dtTemp.Rows[i][5] = vMaxDays.ToString();
                    }

                    int vCalDays = (eCalDays.Text.Trim() != "") ? Int32.Parse(eCalDays.Text.Trim()) : 30;
                    DateTime vStopDate = DateTime.Parse(eStopDate.Text.Trim()).AddDays(vCalDays * -1);
                    string vReportName = "駕駛員最近30天出勤不符七休一規定名單";
                    string vDateRange = "日期範圍：自 " + vStopDate.ToShortDateString() + " 起至 " + eStopDate.Text.Trim();

                    DataTable dtPrint = new DataTable();
                    dtPrint.Columns.Add("DepNo", typeof(string));
                    dtPrint.Columns.Add("DepName", typeof(string));
                    dtPrint.Columns.Add("EmpNo", typeof(string));
                    dtPrint.Columns.Add("EmpName", typeof(string));
                    dtPrint.Columns.Add("WorkDaysNear6Day", typeof(int));
                    dtPrint.Columns.Add("OverDays", typeof(int));
                    for (int i = 0; i < dtTemp.Rows.Count; i++)
                    {
                        if (Int32.Parse(dtTemp.Rows[i][5].ToString().Trim()) >= 7)
                        {
                            dtPrint.Rows.Add(dtTemp.Rows[i][0], dtTemp.Rows[i][1], dtTemp.Rows[i][2], dtTemp.Rows[i][3], dtTemp.Rows[i][4], dtTemp.Rows[i][5]);
                        }
                    }

                    ReportDataSource rdsPrint = new ReportDataSource("DriverWorkOver7DaysP", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\DriverWorkOver7DaysP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("DateRange", vDateRange));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("FieldName", "最近 " + eLastDays.Text.Trim() + " 天出勤天數"));
                    rvPrint.LocalReport.Refresh();
                    plShowData.Visible = false;
                    plPrint.Visible = true;
                    string vDepNoStr = (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : "全部";
                    string vDriverStr = (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : "全部";
                    string vCalDayStr = (eCalDays.Text.Trim() != "") ? eCalDays.Text.Trim() : "30天";
                    string vLastDayStr = (eLastDays.Text.Trim() != "") ? eLastDays.Text.Trim() : "7天";
                    string vStopDateStr = (eStopDate.Text.Trim() != "") ? eStopDate.Text.Trim() : "";
                    string vRecordNote = "預覽報表_駕駛員最近30天出勤狀況 (超時人員)" + Environment.NewLine +
                                         "CustomServiceDriverList.aspx" + Environment.NewLine +
                                         "站別：" + vDepNoStr + Environment.NewLine +
                                         "駕駛員編號：" + vDriverStr + Environment.NewLine +
                                         "回算天數：" + vCalDayStr + Environment.NewLine +
                                         "最近天數：" + vLastDayStr + Environment.NewLine +
                                         "截止日期：" + vStopDateStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        /// <summary>
        /// 超時人員資料匯出EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_OverTime_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetOverTimeSelStr();
            int vMaxDays = 0;
            int vDaysCount = 0;
            Boolean vIsRed = false;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                DataTable dtExcel = new DataTable();
                SqlDataAdapter daExcel = new SqlDataAdapter(cmdExcel);
                daExcel.Fill(dtExcel);
                if (dtExcel.Rows.Count > 0)
                {
                    //查詢結果有資料的時候才執行
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vMaxDays = 0;
                        vDaysCount = 0;
                        string vEmpNo_DataTable = dtExcel.Rows[i][2].ToString();
                        for (int J = 6; J < dtExcel.Columns.Count; J++)
                        {
                            if (dtExcel.Rows[i][J].ToString().Trim() == "O")
                            {
                                vDaysCount = vDaysCount + 1;
                            }
                            else
                            {
                                if (vMaxDays < vDaysCount)
                                {
                                    vMaxDays = vDaysCount;
                                }
                                vDaysCount = 0;
                            }
                        }
                        if (vMaxDays < vDaysCount)
                        {
                            vMaxDays = vDaysCount;
                        }
                        vDaysCount = 0;
                        dtExcel.Rows[i][5] = vMaxDays.ToString();
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
                    fontData.FontHeightInPoints = 12; //字體大小
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
                    string vFileName = "駕駛員最近30天出勤不符七休一規定名單";
                    int vLinesNo = 0;

                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < 6; i++)
                    {
                        vHeaderText = dtExcel.Columns[i].ColumnName.Trim();
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vIsRed = (Int32.Parse(dtExcel.Rows[i][5].ToString().Trim()) >= 7);
                        if (vIsRed)
                        {
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            for (int j = 0; j < 6; j++)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(j);
                                if (j >= 4)
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(double.Parse(dtExcel.Rows[i][j].ToString().Trim()));
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(dtExcel.Rows[i][j].ToString().Trim());
                                }
                                wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
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
                            string vDepNoStr = (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : "全部";
                            string vDriverStr = (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : "全部";
                            string vCalDayStr = (eCalDays.Text.Trim() != "") ? eCalDays.Text.Trim() : "30天";
                            string vLastDayStr = (eLastDays.Text.Trim() != "") ? eLastDays.Text.Trim() : "7天";
                            string vStopDateStr = (eStopDate.Text.Trim() != "") ? eStopDate.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_駕駛員最近30天出勤狀況 (超時人員)" + Environment.NewLine +
                                                 "CustomServiceDriverList.aspx" + Environment.NewLine +
                                                 "站別：" + vDepNoStr + Environment.NewLine +
                                                 "駕駛員編號：" + vDriverStr + Environment.NewLine +
                                                 "回算天數：" + vCalDayStr + Environment.NewLine +
                                                 "最近天數：" + vLastDayStr + Environment.NewLine +
                                                 "截止日期：" + vStopDateStr;
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

        protected void bbCLoseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
        }
    }
}