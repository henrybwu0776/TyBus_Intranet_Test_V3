using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class CarFixWorkList : System.Web.UI.Page
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
        DateTime vCalDate = DateTime.Today.AddMonths(-1);

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                Thread.CurrentThread.CurrentCulture = Cal;

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
                        eFixYear_Search.Text = (vCalDate.Year - 1911).ToString();
                        eFixMonth_Search.Text = vCalDate.Month.ToString();
                    }
                    else
                    {

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

        protected void ddlDepNo_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDepNo_Search.Text = ddlDepNo_Search.SelectedValue.Trim();
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vRecordDepNoStr = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "全部";
            string vRecordDateStr = eFixYear_Search.Text.Trim() + "年" + eFixMonth_Search.Text.Trim() + "月";
            string vRecordOutputTypeStr = rbOutputMode.SelectedItem.Text.Trim();
            string vRecordNote = "匯出檔案_修車廠工作單資料匯出" + Environment.NewLine +
                                 "CarFixWorkList.aspx" + Environment.NewLine +
                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                 "進廠日期：" + vRecordDateStr + Environment.NewLine +
                                 "資料類別：" + vRecordOutputTypeStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

            switch (rbOutputMode.SelectedValue)
            {
                case "0": //全部資料
                    ExportAllData();
                    break;
                case "1": //工作項目與標準工時
                    ExportWorkItemData();
                    break;
                case "2": //工作人時
                    ExportPersonTimeData();
                    break;
            }
        }

        /// <summary>
        /// 同時取回工時和人時資料
        /// </summary>
        /// <returns></returns>
        private string getAllWorkSelectStr()
        {
            string vCalYear = (eFixYear_Search.Text.Trim() != "") ? (Int32.Parse(eFixYear_Search.Text.Trim()) + 1911).ToString() : DateTime.Today.AddMonths(-1).Year.ToString();
            string vCalMonth = (eFixMonth_Search.Text.Trim() != "") ? eFixMonth_Search.Text.Trim() : DateTime.Today.AddMonths(-1).Month.ToString();
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vReturnStr = "select a.WorkNo, a.BuDate, a.Car_ID, a.DepNo, (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                                "       a.MainTain1, (select ClassTxt from DBDICB where FKey = '工作單A         FixworkA        MAINTAIN1' and ClassNo = a.MainTain1) MainTain1_C, " + Environment.NewLine +
                                "       a.MainTain, (select ClassTxt from DBDICB where FKey = '工作單A         FixworkA        MAINTAIN' and ClassNo = a.MainTain) MainTain_C, " + Environment.NewLine +
                                "       a.[Service], (select ClassTxt from DBDICB where FKey = '工作單A         FixworkA        SERVICE' and ClassNo = a.[Service]) Service_C, " + Environment.NewLine +
                                "       d.Car_Class, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and CLassNo = d.Car_Class) Car_Class_C, " + Environment.NewLine +
                                "       d.ProdDate, d.Car_TypeID, d.Canon, a.FixDateIn, a.TotTime, a.FixDateOut, a.EndTime, " + Environment.NewLine +
                                "       b.RealNo, (select SevName from FixSev where SevNo = b.RealNo) SevName, (select SevHour from FixSev where SevNo = b.RealNo) SevHour, " + Environment.NewLine +
                                "       a.Driver, (select[Name] from Employee where EmpNo = a.Driver) DriverName, " + Environment.NewLine +
                                "       a.BuMan, (select[Name] from Employee where EmpNo = a.BuMan) BuManName, " + Environment.NewLine +
                                "       a.Driver, (select[Name] from Employee where EmpNo = a.Driver) DriverName, " + Environment.NewLine +
                                "       a.FixMan, (select[Name] from EMployee where EmpNo = a.FixMan) FixManName, " + Environment.NewLine +
                                "       c.FixDate, c.FixMan FixMan_C, (select[Name] from Employee where EmpNo = c.FixMan) FixManName_C, c.[Hours], c.[Minute], a.Remark " + Environment.NewLine +
                                "  from FixWorkA a left join FixWorkB b on b.WorkNo = a.WorkNo " + Environment.NewLine +
                                "                  left join FixWorkC c on c.WorkNo = a.WorkNo " + Environment.NewLine +
                                "                  left join Car_InfoA d on d.Car_ID = a.Car_ID " + Environment.NewLine +
                                " where year(a.FixDateIn) = " + vCalYear + " and month(a.FixDateIn) = " + vCalMonth + Environment.NewLine +
                                vWStr_DepNo +
                                " order by a.WorkNo, b.Item";
            return vReturnStr;
        }

        /// <summary>
        /// 取回工作項目與標準工時的查詢字串
        /// </summary>
        /// <returns></returns>
        private string getWorkItemSelectStr()
        {
            string vCalYear = (eFixYear_Search.Text.Trim() != "") ? (Int32.Parse(eFixYear_Search.Text.Trim()) + 1911).ToString() : DateTime.Today.AddMonths(-1).Year.ToString();
            string vCalMonth = (eFixMonth_Search.Text.Trim() != "") ? eFixMonth_Search.Text.Trim() : DateTime.Today.AddMonths(-1).Month.ToString();
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vReturnStr = "select a.WorkNo, a.BuDate, a.Car_ID, a.DepNo, (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                                "       a.MainTain1, (select ClassTxt from DBDICB where FKey = '工作單A         FixworkA        MAINTAIN1' and ClassNo = a.MainTain1) MainTain1_C, " + Environment.NewLine +
                                "       a.MainTain, (select ClassTxt from DBDICB where FKey = '工作單A         FixworkA        MAINTAIN' and ClassNo = a.MainTain) MainTain_C, " + Environment.NewLine +
                                "       a.[Service], (select ClassTxt from DBDICB where FKey = '工作單A         FixworkA        SERVICE' and ClassNo = a.[Service]) Service_C, " + Environment.NewLine +
                                "       c.Car_Class, (select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and CLassNo = c.Car_Class) Car_Class_C, " + Environment.NewLine +
                                "       c.ProdDate, c.Canon, a.FixDateIn, a.TotTime, a.FixDateOut, a.EndTime, " + Environment.NewLine +
                                "       b.RealNo, (select SevName from FixSev where SevNo = b.RealNo) SevName, (select SevHour from FixSev where SevNo = b.RealNo) SevHour, " + Environment.NewLine +
                                "       a.Driver, (select[Name] from Employee where EmpNo = a.Driver) DriverName, " + Environment.NewLine +
                                "       a.BuMan, (select[Name] from Employee where EmpNo = a.BuMan) BuManName, a.Remark " + Environment.NewLine +
                                "  from FixWorkA a left join FixWorkB b on b.WorkNo = a.WorkNo " + Environment.NewLine +
                                "                  left join Car_InfoA c on c.Car_ID = a.Car_ID " + Environment.NewLine +
                                " where year(a.FixDateIn) = " + vCalYear + " and month(a.FixDateIn) = " + vCalMonth + Environment.NewLine +
                                vWStr_DepNo +
                                " order by a.WorkNo, b.Item";
            return vReturnStr;
        }

        /// <summary>
        /// 取回工作人時的查詢字串
        /// </summary>
        /// <returns></returns>
        private string getPersonTimeSelectStr()
        {
            string vCalYear = (eFixYear_Search.Text.Trim() != "") ? (Int32.Parse(eFixYear_Search.Text.Trim()) + 1911).ToString() : DateTime.Today.AddMonths(-1).Year.ToString();
            string vCalMonth = (eFixMonth_Search.Text.Trim() != "") ? eFixMonth_Search.Text.Trim() : DateTime.Today.AddMonths(-1).Month.ToString();
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vReturnStr = "select a.WorkNo, a.BuDate, a.Car_ID, a.DepNo, (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                                "       a.Driver, (select[Name] from Employee where EmpNo = a.Driver) DriverName, " + Environment.NewLine +
                                "       a.BuMan, (select[Name] from Employee where EmpNo = a.BuMan) BuManName, " + Environment.NewLine +
                                "       a.FixMan, (select[Name] from EMployee where EmpNo = a.FixMan) FixManName, " + Environment.NewLine +
                                "       c.FixDate, c.FixMan FixMan_C, (select[Name] from Employee where EmpNo = c.FixMan) FixManName_C, " + Environment.NewLine +
                                "       c.[Hours], c.[Minute] " + Environment.NewLine +
                                "  from FixWorkA a left join FixWorkC c on c.WorkNo = a.WorkNo " + Environment.NewLine +
                                " where year(a.FixDateIN) = " + vCalYear + " and Month(a.FixDateIn) = " + vCalMonth + Environment.NewLine +
                                vWStr_DepNo +
                                " order by a.WorkNo, c.FixMan";
            return vReturnStr;
        }

        /// <summary>
        /// 匯出包括工作項目、標準工時和工作人時的資料
        /// </summary>
        private void ExportAllData()
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
            int vLinesNo = 0;

            //取回查詢字串
            string vSelectStr = getAllWorkSelectStr();
            string vFileName = "桃園客運_" + eFixYear_Search.Text.Trim() + "_年_" + eFixMonth_Search.Text.Trim() + "_月_修車廠工作單資料清冊";
            string vSheetName = "工作項目、標準工時及工作人時";
            DateTime vBuDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "WORKNO") ? "工作單號" :
                                      (drExcel.GetName(i).ToUpper() == "BUDATE") ? "開單日期" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "車號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "場站編號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN1") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN1_C") ? "小修類別" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN_C") ? "維修項目" :
                                      (drExcel.GetName(i).ToUpper() == "SERVICE") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "SERVICE_C") ? "保養級別" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_CLASS") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_CLASS_C") ? "廠牌" :
                                      (drExcel.GetName(i).ToUpper() == "PRODDATE") ? "年份" :
                                      (drExcel.GetName(i).ToUpper() == "CANON") ? "污染排放標準值" :
                                      (drExcel.GetName(i).ToUpper() == "FIXDATEIN") ? "進廠日期" :
                                      (drExcel.GetName(i).ToUpper() == "TOTTIME") ? "進廠時間" :
                                      (drExcel.GetName(i).ToUpper() == "FIXDATEOUT") ? "出廠日期" :
                                      (drExcel.GetName(i).ToUpper() == "ENDTIME") ? "出廠時間" :
                                      (drExcel.GetName(i).ToUpper() == "REALNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "SEVNAME") ? "工作項目" :
                                      (drExcel.GetName(i).ToUpper() == "SEVHOUR") ? "標準工時" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員工號" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "BUMAN") ? "開單人工號" :
                                      (drExcel.GetName(i).ToUpper() == "BUMANNAME") ? "開單人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK") ? "備註" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員工號" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "BUMAN") ? "開單人工號" :
                                      (drExcel.GetName(i).ToUpper() == "BUMANNAME") ? "開單人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMAN") ? "承修人工號" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMANNAME") ? "承修人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "FIXDATE") ? "承修日期" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMAN_C") ? "維修人工號" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMANNAME_C") ? "維修人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "HOURS") ? "工作小時" :
                                      (drExcel.GetName(i).ToUpper() == "MINUTE") ? "工作分鐘" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_TYPEID") ? "型別" : drExcel.GetName(i);
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
                                 (drExcel.GetName(i).ToUpper() == "FIXDATEIN") ||
                                 (drExcel.GetName(i).ToUpper() == "FIXDATEOUT") ||
                                 (drExcel.GetName(i).ToUpper() == "FIXDATE")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i).ToUpper() == "CANON") ||
                                 (drExcel.GetName(i).ToUpper() == "SEVHOUR") ||
                                 (drExcel.GetName(i).ToUpper() == "HOURS") ||
                                 (drExcel.GetName(i).ToUpper() == "MINUTE")) && (drExcel[i].ToString() != ""))
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
                }
                else
                {

                }
            }

            //取回查詢字串
            vSelectStr = getWorkItemSelectStr();
            vSheetName = "工作項目與標準工時";

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "WORKNO") ? "工作單號" :
                                      (drExcel.GetName(i).ToUpper() == "BUDATE") ? "開單日期" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "車號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "場站編號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN1") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN1_C") ? "小修類別" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN_C") ? "維修項目" :
                                      (drExcel.GetName(i).ToUpper() == "SERVICE") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "SERVICE_C") ? "保養級別" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_CLASS") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_CLASS_C") ? "廠牌" :
                                      (drExcel.GetName(i).ToUpper() == "PRODDATE") ? "年份" :
                                      (drExcel.GetName(i).ToUpper() == "CANON") ? "污染排放標準值" :
                                      (drExcel.GetName(i).ToUpper() == "FIXDATEIN") ? "進廠日期" :
                                      (drExcel.GetName(i).ToUpper() == "TOTTIME") ? "進廠時間" :
                                      (drExcel.GetName(i).ToUpper() == "FIXDATEOUT") ? "出廠日期" :
                                      (drExcel.GetName(i).ToUpper() == "ENDTIME") ? "出廠時間" :
                                      (drExcel.GetName(i).ToUpper() == "REALNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "SEVNAME") ? "工作項目" :
                                      (drExcel.GetName(i).ToUpper() == "SEVHOUR") ? "標準工時" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員工號" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "BUMAN") ? "開單人工號" :
                                      (drExcel.GetName(i).ToUpper() == "BUMANNAME") ? "開單人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK") ? "備註" : drExcel.GetName(i);
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
                                 (drExcel.GetName(i).ToUpper() == "FIXDATEIN") ||
                                 (drExcel.GetName(i).ToUpper() == "FIXDATEOUT")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i).ToUpper() == "CANON") ||
                                 (drExcel.GetName(i).ToUpper() == "SEVHOUR")) && (drExcel[i].ToString() != ""))
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
                }
                else
                {

                }
            }

            //取回查詢字串
            vSelectStr = getPersonTimeSelectStr();
            vSheetName = "工作人時";

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "WORKNO") ? "工作單號" :
                                      (drExcel.GetName(i).ToUpper() == "BUDATE") ? "開單日期" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "車號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "場站編號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員工號" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "BUMAN") ? "開單人工號" :
                                      (drExcel.GetName(i).ToUpper() == "BUMANNAME") ? "開單人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMAN") ? "承修人工號" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMANNAME") ? "承修人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "FIXDATE") ? "承修日期" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMAN_C") ? "維修人工號" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMANNAME_C") ? "維修人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "HOURS") ? "工作小時" :
                                      (drExcel.GetName(i).ToUpper() == "MINUTE") ? "工作分鐘" : drExcel.GetName(i);
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
                                 (drExcel.GetName(i).ToUpper() == "FIXDATE")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i).ToUpper() == "HOURS") ||
                                 (drExcel.GetName(i).ToUpper() == "MINUTE")) && (drExcel[i].ToString() != ""))
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
                }
                else
                {

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

        /// <summary>
        /// 匯出工作項目與標準工時
        /// </summary>
        private void ExportWorkItemData()
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
            int vLinesNo = 0;

            //取回查詢字串
            string vSelectStr = getWorkItemSelectStr();
            string vFileName = "桃園客運_" + eFixYear_Search.Text.Trim() + "_年_" + eFixMonth_Search.Text.Trim() + "_月_修車廠工作單工作項目與標準工時資料清冊";
            string vSheetName = "工作項目與標準工時";
            DateTime vBuDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "WORKNO") ? "工作單號" :
                                      (drExcel.GetName(i).ToUpper() == "BUDATE") ? "開單日期" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "車號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "場站編號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN1") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN1_C") ? "小修類別" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "MAINTAIN_C") ? "維修項目" :
                                      (drExcel.GetName(i).ToUpper() == "SERVICE") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "SERVICE_C") ? "保養級別" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_CLASS") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_CLASS_C") ? "廠牌" :
                                      (drExcel.GetName(i).ToUpper() == "PRODDATE") ? "年份" :
                                      (drExcel.GetName(i).ToUpper() == "CANON") ? "污染排放標準值" :
                                      (drExcel.GetName(i).ToUpper() == "FIXDATEIN") ? "進廠日期" :
                                      (drExcel.GetName(i).ToUpper() == "TOTTIME") ? "進廠時間" :
                                      (drExcel.GetName(i).ToUpper() == "FIXDATEOUT") ? "出廠日期" :
                                      (drExcel.GetName(i).ToUpper() == "ENDTIME") ? "出廠時間" :
                                      (drExcel.GetName(i).ToUpper() == "REALNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "REALNO_C") ? "工作項目" :
                                      (drExcel.GetName(i).ToUpper() == "SEVHOUR") ? "標準工時" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員工號" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "BUMAN") ? "開單人工號" :
                                      (drExcel.GetName(i).ToUpper() == "BUMANNAME") ? "開單人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK") ? "備註" : drExcel.GetName(i);
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
                                 (drExcel.GetName(i).ToUpper() == "FIXDATEIN") ||
                                 (drExcel.GetName(i).ToUpper() == "FIXDATEOUT")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i).ToUpper() == "CANON") ||
                                 (drExcel.GetName(i).ToUpper() == "SEVHOUR")) && (drExcel[i].ToString() != ""))
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

                }
            }
        }

        /// <summary>
        /// 匯出工作人時
        /// </summary>
        private void ExportPersonTimeData()
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
            int vLinesNo = 0;

            //取回查詢字串
            string vSelectStr = getPersonTimeSelectStr();
            string vFileName = "桃園客運_" + eFixYear_Search.Text.Trim() + "_年_" + eFixMonth_Search.Text.Trim() + "_月_修車廠工作單工作人時資料清冊";
            string vSheetName = "工作人時";
            DateTime vBuDate;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "WORKNO") ? "工作單號" :
                                      (drExcel.GetName(i).ToUpper() == "BUDATE") ? "開單日期" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "車號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNO") ? "場站編號" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVER") ? "駕駛員工號" :
                                      (drExcel.GetName(i).ToUpper() == "DRIVERNAME") ? "駕駛員姓名" :
                                      (drExcel.GetName(i).ToUpper() == "BUMAN") ? "開單人工號" :
                                      (drExcel.GetName(i).ToUpper() == "BUMANNAME") ? "開單人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMAN") ? "承修人工號" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMANNAME") ? "承修人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "FIXDATE") ? "承修日期" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMAN_C") ? "維修人工號" :
                                      (drExcel.GetName(i).ToUpper() == "FIXMANNAME_C") ? "維修人姓名" :
                                      (drExcel.GetName(i).ToUpper() == "HOURS") ? "工作小時" :
                                      (drExcel.GetName(i).ToUpper() == "MINUTE") ? "工作分鐘" : drExcel.GetName(i);
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
                                 (drExcel.GetName(i).ToUpper() == "FIXDATE")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i).ToUpper() == "HOURS") ||
                                 (drExcel.GetName(i).ToUpper() == "MINUTE")) && (drExcel[i].ToString() != ""))
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

                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}