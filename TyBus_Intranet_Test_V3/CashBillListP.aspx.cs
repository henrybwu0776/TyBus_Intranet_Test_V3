using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class CashBillListP : System.Web.UI.Page
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
        private DateTime vToday = DateTime.Today;

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
                        rbListMode.SelectedIndex = 0;
                        eMonthList_Year.Text = (vToday.AddMonths(-1).Year - 1911).ToString();
                        eMonthList_Month.Text = vToday.AddMonths(-1).Month.ToString();
                        plSearch.Visible = true;
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

        private string GetSelectStr(string vSelectMode)
        {
            string vSelectStr = "";
            string vDateSel_S = "";
            string vDateSel_E = "";
            string vWStr_DepNo = "";

            switch (vSelectMode)
            {
                case "1":
                    vDateSel_S = ((eMonthList_Year.Text.Trim() != "") && (eMonthList_Month.Text.Trim() != "")) ?
                                 (Int32.Parse(eMonthList_Year.Text.Trim()) + 1911).ToString("D4") + "/" + eMonthList_Month.Text.Trim() + "/01" :
                                 PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "B");
                    vDateSel_E = PF.GetMonthLastDay(DateTime.Parse(vDateSel_S), "B");
                    vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "           and c.DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";

                    vSelectStr = "select z.DepNo, (select [Name] from Department where DepNo = z.DepNo) DepName, " + Environment.NewLine +
                                 "       z.Driver, (select [Name] from Employee where EmpNo = z.Driver) Driver_C, z.LinesNo, " + Environment.NewLine +
                                 "       sum(z.TotalAmount01) TotalAmount01, sum(z.TotalAmount02) TotalAmount02, " + Environment.NewLine +
                                 "       sum(z.TotalAmount03) TotalAmount03, sum(z.TotalAmount04) TotalAmount04, " + Environment.NewLine +
                                 "       sum(z.TotalAmount05) TotalAmount05, sum(z.TotalAmount06) TotalAmount06, " + Environment.NewLine +
                                 "       sum(z.TotalAmount07) TotalAmount07, sum(z.TotalAmount08) TotalAmount08, " + Environment.NewLine +
                                 "       sum(z.TotalAmount09) TotalAmount09, sum(z.TotalAmount10) TotalAmount10, " + Environment.NewLine +
                                 "       sum(z.TotalAmount11) TotalAmount11, sum(z.TotalAmount12) TotalAmount12, " + Environment.NewLine +
                                 "       sum(z.TotalAmount13) TotalAmount13, sum(z.TotalAmount14) TotalAmount14, " + Environment.NewLine +
                                 "       sum(z.TotalAmount15) TotalAmount15, sum(z.TotalAmount16) TotalAmount16, " + Environment.NewLine +
                                 "       sum(z.TotalAmount17) TotalAmount17, sum(z.TotalAmount18) TotalAmount18, " + Environment.NewLine +
                                 "       sum(z.TotalAmount19) TotalAmount19, sum(z.TotalAmount20) TotalAmount20, " + Environment.NewLine +
                                 "       sum(z.TotalAmount21) TotalAmount21, sum(z.TotalAmount22) TotalAmount22, " + Environment.NewLine +
                                 "       sum(z.TotalAmount23) TotalAmount23, sum(z.TotalAmount24) TotalAmount24, " + Environment.NewLine +
                                 "       sum(z.TotalAmount25) TotalAmount25, sum(z.TotalAmount26) TotalAmount26, " + Environment.NewLine +
                                 "       sum(z.TotalAmount27) TotalAmount27, sum(z.TotalAmount28) TotalAmount28, " + Environment.NewLine +
                                 "       sum(z.TotalAmount29) TotalAmount29, sum(z.TotalAmount30) TotalAmount30, " + Environment.NewLine +
                                 "       sum(z.TotalAmount31) TotalAmount31 " + Environment.NewLine +
                                 "  from ( " + Environment.NewLine +
                                 "        select DepNo, Driver, LinesNo, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 1 then TotalAmt else null end TotalAmount01, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 2 then TotalAmt else null end TotalAmount02, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 3 then TotalAmt else null end TotalAmount03, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 4 then TotalAmt else null end TotalAmount04, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 5 then TotalAmt else null end TotalAmount05, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 6 then TotalAmt else null end TotalAmount06, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 7 then TotalAmt else null end TotalAmount07, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 8 then TotalAmt else null end TotalAmount08, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 9 then TotalAmt else null end TotalAmount09, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 10 then TotalAmt else null end TotalAmount10, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 11 then TotalAmt else null end TotalAmount11, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 12 then TotalAmt else null end TotalAmount12, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 13 then TotalAmt else null end TotalAmount13, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 14 then TotalAmt else null end TotalAmount14, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 15 then TotalAmt else null end TotalAmount15, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 16 then TotalAmt else null end TotalAmount16, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 17 then TotalAmt else null end TotalAmount17, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 18 then TotalAmt else null end TotalAmount18, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 19 then TotalAmt else null end TotalAmount19, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 20 then TotalAmt else null end TotalAmount20, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 21 then TotalAmt else null end TotalAmount21, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 22 then TotalAmt else null end TotalAmount22, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 23 then TotalAmt else null end TotalAmount23, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 24 then TotalAmt else null end TotalAmount24, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 25 then TotalAmt else null end TotalAmount25, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 26 then TotalAmt else null end TotalAmount26, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 27 then TotalAmt else null end TotalAmount27, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 28 then TotalAmt else null end TotalAmount28, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 29 then TotalAmt else null end TotalAmount29, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 30 then TotalAmt else null end TotalAmount30, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 31 then TotalAmt else null end TotalAmount31 " + Environment.NewLine +
                                 "          from CashBill c " + Environment.NewLine +
                                 "         where c.CBDate between '" + vDateSel_S + "' and '" + vDateSel_E + "' " + Environment.NewLine + vWStr_DepNo +
                                 ") z " + Environment.NewLine +
                                 " group by z.DepNo, z.Driver, z.LinesNo " + Environment.NewLine +
                                 " order by z.DepNo, z.Driver, z.LinesNo ";
                    break;
                case "2":
                    vDateSel_S = ((eMonthList_Year.Text.Trim() != "") && (eMonthList_Month.Text.Trim() != "")) ?
                                 (Int32.Parse(eMonthList_Year.Text.Trim()) + 1911).ToString("D4") + "/" + eMonthList_Month.Text.Trim() + "/01" :
                                 PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "B");
                    vDateSel_E = PF.GetMonthLastDay(DateTime.Parse(vDateSel_S), "B");
                    vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "           and c.DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";

                    vSelectStr = "select z.DepNo, (select [Name] from Department where DepNo = z.DepNo) DepName, " + Environment.NewLine +
                                 "       z.Car_ID, z.Car_No, z.LinesNo, " + Environment.NewLine +
                                 "       sum(z.TotalAmount01) TotalAmount01, sum(z.TotalAmount02) TotalAmount02, " + Environment.NewLine +
                                 "       sum(z.TotalAmount03) TotalAmount03, sum(z.TotalAmount04) TotalAmount04, " + Environment.NewLine +
                                 "       sum(z.TotalAmount05) TotalAmount05, sum(z.TotalAmount06) TotalAmount06, " + Environment.NewLine +
                                 "       sum(z.TotalAmount07) TotalAmount07, sum(z.TotalAmount08) TotalAmount08, " + Environment.NewLine +
                                 "       sum(z.TotalAmount09) TotalAmount09, sum(z.TotalAmount10) TotalAmount10, " + Environment.NewLine +
                                 "       sum(z.TotalAmount11) TotalAmount11, sum(z.TotalAmount12) TotalAmount12, " + Environment.NewLine +
                                 "       sum(z.TotalAmount13) TotalAmount13, sum(z.TotalAmount14) TotalAmount14, " + Environment.NewLine +
                                 "       sum(z.TotalAmount15) TotalAmount15, sum(z.TotalAmount16) TotalAmount16, " + Environment.NewLine +
                                 "       sum(z.TotalAmount17) TotalAmount17, sum(z.TotalAmount18) TotalAmount18, " + Environment.NewLine +
                                 "       sum(z.TotalAmount19) TotalAmount19, sum(z.TotalAmount20) TotalAmount20, " + Environment.NewLine +
                                 "       sum(z.TotalAmount21) TotalAmount21, sum(z.TotalAmount22) TotalAmount22, " + Environment.NewLine +
                                 "       sum(z.TotalAmount23) TotalAmount23, sum(z.TotalAmount24) TotalAmount24, " + Environment.NewLine +
                                 "       sum(z.TotalAmount25) TotalAmount25, sum(z.TotalAmount26) TotalAmount26, " + Environment.NewLine +
                                 "       sum(z.TotalAmount27) TotalAmount27, sum(z.TotalAmount28) TotalAmount28, " + Environment.NewLine +
                                 "       sum(z.TotalAmount29) TotalAmount29, sum(z.TotalAmount30) TotalAmount30, " + Environment.NewLine +
                                 "       sum(z.TotalAmount31) TotalAmount31 " + Environment.NewLine +
                                 "  from ( " + Environment.NewLine +
                                 "        select DepNo, Car_ID, Car_No, LinesNo, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 1 then TotalAmt else null end TotalAmount01, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 2 then TotalAmt else null end TotalAmount02, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 3 then TotalAmt else null end TotalAmount03, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 4 then TotalAmt else null end TotalAmount04, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 5 then TotalAmt else null end TotalAmount05, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 6 then TotalAmt else null end TotalAmount06, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 7 then TotalAmt else null end TotalAmount07, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 8 then TotalAmt else null end TotalAmount08, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 9 then TotalAmt else null end TotalAmount09, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 10 then TotalAmt else null end TotalAmount10, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 11 then TotalAmt else null end TotalAmount11, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 12 then TotalAmt else null end TotalAmount12, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 13 then TotalAmt else null end TotalAmount13, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 14 then TotalAmt else null end TotalAmount14, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 15 then TotalAmt else null end TotalAmount15, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 16 then TotalAmt else null end TotalAmount16, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 17 then TotalAmt else null end TotalAmount17, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 18 then TotalAmt else null end TotalAmount18, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 19 then TotalAmt else null end TotalAmount19, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 20 then TotalAmt else null end TotalAmount20, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 21 then TotalAmt else null end TotalAmount21, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 22 then TotalAmt else null end TotalAmount22, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 23 then TotalAmt else null end TotalAmount23, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 24 then TotalAmt else null end TotalAmount24, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 25 then TotalAmt else null end TotalAmount25, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 26 then TotalAmt else null end TotalAmount26, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 27 then TotalAmt else null end TotalAmount27, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 28 then TotalAmt else null end TotalAmount28, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 29 then TotalAmt else null end TotalAmount29, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 30 then TotalAmt else null end TotalAmount30, " + Environment.NewLine +
                                 "               case when day(c.CBDate) = 31 then TotalAmt else null end TotalAmount31 " + Environment.NewLine +
                                 "          from CashBill c " + Environment.NewLine +
                                 "         where c.CBDate between '" + vDateSel_S + "' and '" + vDateSel_E + "' " + Environment.NewLine + vWStr_DepNo +
                                 ") z " + Environment.NewLine +
                                 " group by z.DepNo, z.Car_ID, z.Car_No, z.LinesNo " + Environment.NewLine +
                                 " order by z.DepNo, z.Car_ID, z.Car_No, z.LinesNo ";
                    break;
            }
            return vSelectStr;
        }

        protected void ddlDepNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDepNo.Text = ddlDepNo.SelectedValue.Trim();
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelStr = GetSelectStr(rbListMode.SelectedValue);
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
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
                    //string vFileName = eMonthList_Year.Text.Trim() + "年" + eMonthList_Month.Text.Trim() + "月份" + rbListMode.SelectedItem.Text.Trim();
                    string vFileName = "";
                    vFileName = string.Format("{0}年{1}月份{2}", eMonthList_Year.Text.Trim(), eMonthList_Month.Text.Trim(), rbListMode.SelectedItem.Text.Trim());
                    string vCalMonth = string.Format("{0}/{1}/01", eMonthList_Year.Text.Trim(), eMonthList_Month.Text.Trim());
                    //int vMonthDay = PF.GetMonthDays(DateTime.Parse((Int32.Parse(eMonthList_Year.Text.Trim()) + 1911).ToString("D4") + "/" + eMonthList_Month.Text.Trim() + "/01"));
                    int vMonthDay = PF.GetMonthDays(DateTime.Parse(vCalMonth));
                    int vCalIndex = 0;
                    int vColIndex = 0;
                    int vFirstIntColIndex = 0;
                    string vExcelCellName_1 = "";
                    string vExcelCellName_2 = "";
                    int vLinesNo = 0;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        if (drExcel.GetName(i).ToUpper().IndexOf("TOTALAMOUNT") >= 0)
                        {
                            vCalIndex = Int32.Parse(drExcel.GetName(i).Replace("TotalAmount", ""));
                            if (vCalIndex <= vMonthDay)
                            {
                                vColIndex++;
                                vHeaderText = drExcel.GetName(i).Replace("TotalAmount", Int32.Parse(eMonthList_Month.Text.Trim()).ToString("D2") + "/");
                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                            }
                        }
                        else
                        {
                            vColIndex++;
                            vHeaderText = (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員工號" :
                                          (drExcel.GetName(i).ToUpper() == "DRIVER_C") ? "駕駛員姓名" :
                                          (drExcel.GetName(i).ToUpper() == "DEPNO") ? "部門代號" :
                                          (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "部門" :
                                          (drExcel.GetName(i).ToUpper() == "LINESNO") ? "路線代號" :
                                          (drExcel.GetName(i).ToUpper() == "CAR_NO") ? "車輛代碼" :
                                          (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "牌照號碼" :
                                          (drExcel.GetName(i).ToUpper() == "TOTALCOUNT") ? "納金次數" : "";
                            wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                        }
                    }
                    wsExcel.GetRow(vLinesNo).CreateCell(vColIndex).SetCellValue("納金次數");
                    wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csTitle;
                    vColIndex++;
                    wsExcel.GetRow(vLinesNo).CreateCell(vColIndex).SetCellValue("小計");
                    wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csTitle;
                    vLinesNo++;
                    while (drExcel.Read())
                    {
                        wsExcel.CreateRow(vLinesNo);
                        vColIndex = 0;
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            vHeaderText = drExcel.GetName(i).Trim().ToUpper();
                            if (vHeaderText.IndexOf("TOTALAMOUNT") >= 0)
                            {
                                vCalIndex = Int32.Parse(drExcel.GetName(i).Replace("TotalAmount", ""));
                                vFirstIntColIndex = (vCalIndex == 1) ? i : vFirstIntColIndex;
                                if (vCalIndex <= vMonthDay)
                                {
                                    wsExcel.GetRow(vLinesNo).CreateCell(i);
                                    vColIndex++;
                                    // 2020.07.13 應龍潭站要求，金額為 0 的一樣顯示出 0
                                    //if (drExcel[i].ToString().Trim() == "0")
                                    //{
                                    //    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    //    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("");
                                    //    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                    //}
                                    //else
                                    //{
                                    //    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    //    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                    //    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                    //}

                                    if (drExcel[i].ToString().Trim() == "0")
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(0);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                    }
                                    else if (drExcel[i].ToString().Trim() == "")
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("");
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                    }
                                    else
                                    {
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                        wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                    }
                                }
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(i);
                                vColIndex++;
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                        }
                        //計算起迄格數
                        vExcelCellName_1 = ((vFirstIntColIndex / 26) != 0) ?
                                           ((char)((vFirstIntColIndex / 26) + 64)).ToString() + ((char)((vFirstIntColIndex % 26) + 65)).ToString() + (vLinesNo + 1).ToString() :
                                           ((char)((vFirstIntColIndex % 26) + 65)).ToString() + (vLinesNo + 1).ToString();
                        vExcelCellName_2 = ((vColIndex / 26) != 0) ?
                                           ((char)((vColIndex / 26) + 64)).ToString() + ((char)((vColIndex % 26) + 64)).ToString() + (vLinesNo + 1).ToString() :
                                           ((char)((vColIndex % 26) + 64)).ToString() + (vLinesNo + 1).ToString();
                        //加入納金次數
                        wsExcel.GetRow(vLinesNo).CreateCell(vColIndex);
                        wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellType(CellType.Formula);
                        wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellFormula("COUNTIF(" + vExcelCellName_1 + ":" + vExcelCellName_2 + ",\">=0\")");
                        wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csData_Int;
                        //加入行末小計
                        vColIndex++;
                        wsExcel.GetRow(vLinesNo).CreateCell(vColIndex);
                        wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellType(CellType.Formula);
                        wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellFormula("SUM(" + vExcelCellName_1 + ":" + vExcelCellName_2 + ")");
                        wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csData_Int;
                        vLinesNo++;
                    }
                    //加入每日小計
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < vColIndex + 1; i++)
                    {
                        if (i < vFirstIntColIndex - 1)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("");
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                        }
                        else if (i == vFirstIntColIndex - 1)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue("每日小計");
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                        }
                        else if (wsExcel.GetRow(0).GetCell(i).ToString() == "納金次數")
                        {
                            //計算起迄格數
                            vExcelCellName_1 = ((vFirstIntColIndex / 26) != 0) ?
                                               ((char)((vFirstIntColIndex / 26) + 64)).ToString() + ((char)((vFirstIntColIndex % 26) + 65)).ToString() + (vLinesNo + 1).ToString() :
                                               ((char)((vFirstIntColIndex % 26) + 65)).ToString() + (vLinesNo + 1).ToString();
                            vExcelCellName_2 = (((i - 1) / 26) != 0) ?
                                               ((char)(((i - 1) / 26) + 64)).ToString() + ((char)(((i - 1) % 26) + 65)).ToString() + (vLinesNo + 1).ToString() :
                                               ((char)(((i - 1) % 26) + 65)).ToString() + (vLinesNo + 1).ToString();
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Formula);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("COUNTIF(" + vExcelCellName_1 + ":" + vExcelCellName_2 + ",\">0\")");
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                        }
                        else
                        {
                            vExcelCellName_1 = ((i / 26) != 0) ?
                                               ((char)((i / 26) + 64)).ToString() + ((char)((i % 26) + 65)).ToString() :
                                               ((char)((i % 26) + 65)).ToString();
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Formula);
                            wsExcel.GetRow(vLinesNo).GetCell(i).SetCellFormula("SUM(" + vExcelCellName_1 + "2:" + vExcelCellName_1 + vLinesNo.ToString() + ")");
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
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
                            string vExportMode = (rbListMode.SelectedValue == "1") ? "月繳銷清冊_依駕駛員統計" : "月繳銷清冊_依車輛別統計";
                            string vCalYMStr = ((eMonthList_Year.Text.Trim() != "") && (eMonthList_Month.Text.Trim() != "")) ?
                                               eMonthList_Year.Text.Trim() + " 年 " + eMonthList_Month.Text.Trim() + " 月" :
                                               DateTime.Now.Year.ToString() + " 年 " + DateTime.Now.Month.ToString("D2") + " 月";
                            string vCalDepNo = (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_"+vExportMode + Environment.NewLine +
                                                 "CashBillListP.aspx" + Environment.NewLine +
                                                 "計算年月：" + vCalYMStr +
                                                 "車輛站別：" + vCalDepNo;
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
    }
}