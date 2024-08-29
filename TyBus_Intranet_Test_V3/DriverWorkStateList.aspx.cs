using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverWorkStateList : System.Web.UI.Page
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
        private DataTable dtWSList = new DataTable("WorkStateList");
        private string[] vaTypeStr = { "", "平", "休", "例", "國" };

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

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        eSheetYear.Text = (DateTime.Today.AddMonths(-1).Year - 1911).ToString("D3");
                        eSheetMonth.Text = DateTime.Today.AddMonths(-1).Month.ToString("D2");
                        eDepNo_S.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "03") || (vLoginDepNo == "05") || (vLoginDepNo == "06"));
                        eDepNo_E.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "03") || (vLoginDepNo == "05") || (vLoginDepNo == "06"));
                        eDepNo_S.Text = ((Int32.Parse(vLoginDepNo) >= 11) && (Int32.Parse(vLoginDepNo) < 30)) ? vLoginDepNo : "";
                        plReport.Visible = false;
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

        private string GetSelStr()
        {
            string vWorkYear = (eSheetYear.Text.Trim() != "") ? (Int32.Parse(eSheetYear.Text.Trim()) + 1911).ToString("D4") : DateTime.Today.AddMonths(-1).Year.ToString("D4");
            string vWorkMonth = (eSheetMonth.Text.Trim() != "") ? eSheetMonth.Text.Trim() : DateTime.Today.AddMonths(-1).Month.ToString("D2");
            DateTime vWorkYM_S = DateTime.Parse(vWorkYear + "/" + vWorkMonth + "/01");
            DateTime vWorkYM_E = DateTime.Parse(PF.GetMonthLastDay(vWorkYM_S, "B"));
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and e.DepNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and e.DepNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and e.DepNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
            string vSelStr = "select EmpNo, [Name] , DepNo, (select [Name] from Department where DepNo = e.DepNo) DepName " + Environment.NewLine +
                             "  from Employee e " + Environment.NewLine +
                             " where (e.Leaveday is null or e.LeaveDay >= '" + vWorkYM_S.ToString("yyyy/MM/dd") + "') " + Environment.NewLine +
                             "   and e.AssumeDay <= '" + vWorkYM_E.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                             "   and e.[Type] = '20' " + Environment.NewLine +
                             vWStr_DepNo +
                             " order by e.DepNo, e.EmpNo";
            return vSelStr;
        }

        /// <summary>
        /// 資料取回
        /// </summary>
        /// <param name="fEmpNoListStr">取得指定範圍的員工編號</param>
        /// <param name="ShowType">指定顯示方式 1: 畫面顯示 2:準備供轉出 EXCEL 用</param>
        private void WorkStateListDataBind(string fEmpNoListStr, string fShowType)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vWorkYear = (eSheetYear.Text.Trim() != "") ? (Int32.Parse(eSheetYear.Text.Trim()) + 1911).ToString("D4") : DateTime.Today.AddMonths(-1).Year.ToString("D4");
            string vWorkMonth = (eSheetMonth.Text.Trim() != "") ? eSheetMonth.Text.Trim() : DateTime.Today.AddMonths(-1).Month.ToString("D2");
            DateTime vWorkYM_S = DateTime.Parse(vWorkYear + "/" + vWorkMonth + "/01");
            DateTime vWorkYM_E = DateTime.Parse(PF.GetMonthLastDay(vWorkYM_S, "B"));
            int vWorkMonthDays = PF.GetMonthDays(vWorkYM_S);

            string vSelStr = "";
            string vColName = "";
            string vHeaderText = "";
            int vHeaderDays = 0;
            string vWS_Str = "";
            int vESCType1 = 0;
            int vESCType2 = 0;
            int vESCType3 = 0;

            using (SqlConnection connEmpList = new SqlConnection(vConnStr))
            {
                SqlCommand cmdEmpList = new SqlCommand(fEmpNoListStr, connEmpList);
                connEmpList.Open();
                SqlDataReader drEmpList = cmdEmpList.ExecuteReader();
                if (drEmpList.HasRows)
                {
                    //查詢駕駛員有筆數時才執行
                    dtWSList.Clear();
                    dtWSList.Columns.Add("DepNo", typeof(String));
                    dtWSList.Columns.Add("DepName", typeof(String));
                    dtWSList.Columns.Add("Driver", typeof(String));
                    dtWSList.Columns.Add("DriverName", typeof(String));
                    dtWSList.Columns.Add("WS01", typeof(String));
                    dtWSList.Columns.Add("WS02", typeof(String));
                    dtWSList.Columns.Add("WS03", typeof(String));
                    dtWSList.Columns.Add("WS04", typeof(String));
                    dtWSList.Columns.Add("WS05", typeof(String));
                    dtWSList.Columns.Add("WS06", typeof(String));
                    dtWSList.Columns.Add("WS07", typeof(String));
                    dtWSList.Columns.Add("WS08", typeof(String));
                    dtWSList.Columns.Add("WS09", typeof(String));
                    dtWSList.Columns.Add("WS10", typeof(String));
                    dtWSList.Columns.Add("WS11", typeof(String));
                    dtWSList.Columns.Add("WS12", typeof(String));
                    dtWSList.Columns.Add("WS13", typeof(String));
                    dtWSList.Columns.Add("WS14", typeof(String));
                    dtWSList.Columns.Add("WS15", typeof(String));
                    dtWSList.Columns.Add("WS16", typeof(String));
                    dtWSList.Columns.Add("WS17", typeof(String));
                    dtWSList.Columns.Add("WS18", typeof(String));
                    dtWSList.Columns.Add("WS19", typeof(String));
                    dtWSList.Columns.Add("WS20", typeof(String));
                    dtWSList.Columns.Add("WS21", typeof(String));
                    dtWSList.Columns.Add("WS22", typeof(String));
                    dtWSList.Columns.Add("WS23", typeof(String));
                    dtWSList.Columns.Add("WS24", typeof(String));
                    dtWSList.Columns.Add("WS25", typeof(String));
                    dtWSList.Columns.Add("WS26", typeof(String));
                    dtWSList.Columns.Add("WS27", typeof(String));
                    dtWSList.Columns.Add("WS28", typeof(String));
                    dtWSList.Columns.Add("WS29", typeof(String));
                    dtWSList.Columns.Add("WS30", typeof(String));
                    dtWSList.Columns.Add("WS31", typeof(String));
                    dtWSList.Columns.Add("ESCType1", typeof(Int32));
                    dtWSList.Columns.Add("ESCType2", typeof(Int32));
                    dtWSList.Columns.Add("ESCType3", typeof(Int32));

                    while (drEmpList.Read())
                    {
                        vESCType1 = 0;
                        vESCType2 = 0;
                        vESCType3 = 0;
                        DataRow rowsTemp = dtWSList.NewRow();
                        rowsTemp["DepNo"] = drEmpList["DepNo"].ToString().Trim();
                        rowsTemp["DepName"] = drEmpList["DepName"].ToString().Trim();
                        rowsTemp["Driver"] = drEmpList["EmpNo"].ToString().Trim();
                        rowsTemp["DriverName"] = drEmpList["Name"].ToString().Trim();
                        vSelStr = "select Driver, BuDate, " + Environment.NewLine +
                                  "       left((select Classtxt from DBDICB where ClassNo = a.WorkState and FKey = '行車記錄單A     runsheeta       WORKSTATE'), 1) WS " + Environment.NewLine +
                                  "  from RunSheetA a " + Environment.NewLine +
                                  " where BuDate Between '" + vWorkYM_S.Year.ToString("D4") + "/" + vWorkYM_S.ToString("MM/dd") + "' and '" + vWorkYM_E.Year.ToString("D4") + "/" + vWorkYM_E.ToString("MM/dd") + "' " + Environment.NewLine +
                                  "   and Driver = '" + drEmpList["EmpNo"].ToString().Trim() + "' " + Environment.NewLine +
                                  //2023.08.31 新增取回請假資料
                                  "union all " + Environment.NewLine +
                                  "select a.ApplyMan as Driver, a.RealDay as BuDate, " + Environment.NewLine +
                                  "       case when LEN(b.Classtxt) > 2 then left(b.ClassTxt,2) else left(b.ClassTxt, 1) end as WS " + Environment.NewLine +
                                  "  from ESCDuty a left join DBDICB b on b.FKey = '請假資料檔      ESCDUTY         ESCTYPE' and b.ClassNo = a.ESCType " + Environment.NewLine +
                                  " where a.RealDay Between '" + vWorkYM_S.Year.ToString("D4") + "/" + vWorkYM_S.ToString("MM/dd") + "' and '" + vWorkYM_E.Year.ToString("D4") + "/" + vWorkYM_E.ToString("MM/dd") + "' " + Environment.NewLine +
                                  "   and a.ApplyMan = '" + drEmpList["EmpNo"].ToString().Trim() + "' " + Environment.NewLine +
                                  //====================================================================================================
                                  " order by BuDate";
                        using (SqlConnection connWorkState = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdWorkState = new SqlCommand(vSelStr, connWorkState);
                            connWorkState.Open();
                            SqlDataReader drWorkState = cmdWorkState.ExecuteReader();
                            while (drWorkState.Read())
                            {
                                vColName = "WS" + DateTime.Parse(drWorkState["BuDate"].ToString().Trim()).Day.ToString("D2");
                                vWS_Str = drWorkState["WS"].ToString().Trim();
                                rowsTemp[vColName] = vWS_Str;
                                switch (vWS_Str)
                                {
                                    case "休":
                                        vESCType1++;
                                        break;
                                    case "例":
                                        vESCType2++;
                                        break;
                                    case "國":
                                        vESCType3++;
                                        break;
                                }
                            }
                        }
                        rowsTemp["ESCType1"] = vESCType1;
                        rowsTemp["ESCType2"] = vESCType2;
                        rowsTemp["ESCType3"] = vESCType3;
                        dtWSList.Rows.Add(rowsTemp);
                    }

                    if (fShowType == "1")
                    {
                        //產生畫面預覽
                        gridDriverWorkStateList.DataSourceID = null;
                        gridDriverWorkStateList.DataSource = dtWSList;
                        for (int i = 0; i < gridDriverWorkStateList.Columns.Count; i++)
                        {
                            vHeaderText = gridDriverWorkStateList.Columns[i].HeaderText;
                            if ((vHeaderText != "站別編號") &&
                                (vHeaderText != "站別") &&
                                (vHeaderText != "駕駛工號") &&
                                (vHeaderText != "駕駛姓名"))
                            {
                                vHeaderDays = Int32.Parse(vHeaderText);
                                gridDriverWorkStateList.Columns[i].Visible = (vHeaderDays <= vWorkMonthDays);
                            }
                        }
                        gridDriverWorkStateList.DataBind();
                        string vSheetYMStr = eSheetYear.Text.Trim() + " 年 " + eSheetMonth.Text.Trim() + " 月";
                        string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "自 " + eDepNo_S.Text.Trim() + " 站至 " + eDepNo_E.Text.Trim() + " 站 " :
                                           ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() + " 站" :
                                           ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() + " 站" : "全部";
                        string vRecordNote = "查詢資料_駕駛員排班狀況表" + Environment.NewLine +
                                             "DriverWorkStateList.aspx" + Environment.NewLine +
                                             "查詢年月：" + vSheetYMStr + Environment.NewLine +
                                             "站別：" + vDepNoStr;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    }
                    else if (fShowType == "2")
                    {
                        //準備匯出EXCEL
                        HSSFWorkbook wbExcel = new HSSFWorkbook();
                        //Excel 工作表
                        HSSFSheet wsExcel;

                        //設定標題欄位的格式
                        HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
                        csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                        csTitle.Alignment = HorizontalAlignment.Center; //水平置中

                        //設定標題欄位的字體格式
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

                        vHeaderText = "";

                        string vFileName = Int32.Parse(eSheetYear.Text.Trim()).ToString() + "年" + Int32.Parse(eSheetMonth.Text.Trim()).ToString() + "月駕駛員排班狀況表";
                        int vLinesNo = 0;

                        if (vConnStr == "")
                        {
                            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                        }
                        //新增一個工作表
                        wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                        //寫入標題列
                        vLinesNo = 0;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < dtWSList.Columns.Count; i++)
                        {
                            vHeaderText = (dtWSList.Columns[i].Caption == "DepNo") ? "站別編號" :
                                          (dtWSList.Columns[i].Caption == "DepName") ? "站別" :
                                          (dtWSList.Columns[i].Caption == "Driver") ? "駕駛工號" :
                                          (dtWSList.Columns[i].Caption == "DriverName") ? "駕駛姓名" :
                                          dtWSList.Columns[i].Caption.Trim().Replace("WS", "");
                            if ((vHeaderText == "站別編號") ||
                                (vHeaderText == "站別") ||
                                (vHeaderText == "駕駛工號") ||
                                (vHeaderText == "駕駛姓名"))
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                            }
                            else if ((vHeaderText.Length > 7) && (vHeaderText.Substring(0, 7) == "ESCType"))
                            {
                                switch (vHeaderText)
                                {
                                    case "ESCType1":
                                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue("休加");
                                        break;
                                    case "ESCType2":
                                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue("例加");
                                        break;
                                    case "ESCType3":
                                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue("國加");
                                        break;
                                }
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                            }
                            else if (Int32.Parse(vHeaderText) <= vWorkMonthDays)
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                            }
                        }
                        for (int i = 0; i < dtWSList.Rows.Count; i++)
                        {
                            vLinesNo++;
                            wsExcel.CreateRow(vLinesNo);
                            for (int j = 0; j < dtWSList.Columns.Count; j++)
                            {
                                vHeaderText = dtWSList.Columns[j].Caption.Trim();
                                if ((vHeaderText == "DepNo") || (vHeaderText == "DepName") || (vHeaderText == "Driver") || (vHeaderText == "DriverName"))
                                {
                                    wsExcel.GetRow(vLinesNo).CreateCell(j);
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(dtWSList.Rows[i][j].ToString().Trim());
                                    wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
                                }
                                else if ((vHeaderText.Length > 7) && (vHeaderText.Substring(0, 7) == "ESCType"))
                                {
                                    wsExcel.GetRow(vLinesNo).CreateCell(j);
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(dtWSList.Rows[i][j].ToString().Trim());
                                    wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData_Int;
                                }
                                else if (Int32.Parse(vHeaderText.Replace("WS", "")) <= vWorkMonthDays)
                                {
                                    wsExcel.GetRow(vLinesNo).CreateCell(j);
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(dtWSList.Rows[i][j].ToString().Trim());
                                    //出勤狀況不是 "平、例、休、國" 或空白的話就用紅字表示
                                    wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = (Array.IndexOf(vaTypeStr, dtWSList.Rows[i][j].ToString().Trim()) > 0) ? csData : csData_Red;
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
                                string vSheetYMStr = eSheetYear.Text.Trim() + " 年 " + eSheetMonth.Text.Trim() + " 月";
                                string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "自 " + eDepNo_S.Text.Trim() + " 站至 " + eDepNo_E.Text.Trim() + " 站 " :
                                                   ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() + " 站" :
                                                   ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() + " 站" : "全部";
                                string vRecordNote = "匯出檔案_駕駛員排班狀況表" + Environment.NewLine +
                                                     "DriverWorkStateList.aspx" + Environment.NewLine +
                                                     "查詢年月：" + vSheetYMStr + Environment.NewLine +
                                                     "站別：" + vDepNoStr;
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
                    else if (fShowType == "3") //預覽列印
                    {
                        string vESCDutyMember_11 = "";
                        string vESCDutyMember_13 = "";
                        string vESCDutyMember_15 = "";
                        string vESCDutyMember_16 = "";
                        string vESCDutyMember_17 = "";
                        string vESCDutyMember_18 = "";
                        string vESCDutyMember_19 = "";
                        string vESCDutyMember_20 = "";
                        string vESCDutyMember_22 = "";
                        string vESCDutyMember_23 = "";
                        string vESCDutyMember_25 = "";
                        string vTempDepNo = "";
                        string vTempEmpName = "";
                        if (dtWSList.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtWSList.Rows.Count; i++)
                            {
                                vTempDepNo = dtWSList.Rows[i]["DepNo"].ToString().Trim();
                                vTempEmpName = dtWSList.Rows[i]["DriverName"].ToString().Trim();
                                if (Int32.Parse(dtWSList.Rows[i]["ESCType2"].ToString().Trim()) > 0)
                                {
                                    switch (vTempDepNo)
                                    {
                                        case "11":
                                            vESCDutyMember_11 = (vESCDutyMember_11.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_11 + ", " + vTempEmpName;
                                            break;
                                        case "13":
                                            vESCDutyMember_13 = (vESCDutyMember_13.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_13 + ", " + vTempEmpName;
                                            break;
                                        case "15":
                                            vESCDutyMember_15 = (vESCDutyMember_15.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_15 + ", " + vTempEmpName;
                                            break;
                                        case "16":
                                            vESCDutyMember_16 = (vESCDutyMember_16.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_16 + ", " + vTempEmpName;
                                            break;
                                        case "17":
                                            vESCDutyMember_17 = (vESCDutyMember_17.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_17 + ", " + vTempEmpName;
                                            break;
                                        case "18":
                                            vESCDutyMember_18 = (vESCDutyMember_18.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_18 + ", " + vTempEmpName;
                                            break;
                                        case "19":
                                            vESCDutyMember_19 = (vESCDutyMember_19.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_19 + ", " + vTempEmpName;
                                            break;
                                        case "20":
                                            vESCDutyMember_20 = (vESCDutyMember_20.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_20 + ", " + vTempEmpName;
                                            break;
                                        case "22":
                                            vESCDutyMember_22 = (vESCDutyMember_22.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_22 + ", " + vTempEmpName;
                                            break;
                                        case "23":
                                            vESCDutyMember_23 = (vESCDutyMember_23.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_23 + ", " + vTempEmpName;
                                            break;
                                        case "25":
                                            vESCDutyMember_25 = (vESCDutyMember_25.Trim() == "") ? "有出現例加人員：" + vTempEmpName : vESCDutyMember_25 + ", " + vTempEmpName;
                                            break;
                                    }
                                }
                            }
                            string vYM = (Int32.Parse(eSheetYear.Text.Trim()) + 1911).ToString("D4") + eSheetMonth.Text.Trim();
                            string vTempStr = "select cast(cast(left(YYYYMM, 4) as int) - 1911 as varchar) + ' 年 ' + " + Environment.NewLine +
                                              "       cast(cast(right(YYYYMM, 2) as int) as varchar) + ' 月份共計 ' + " + Environment.NewLine +
                                              "       cast(WorkStateDays1 as varchar) + ' 休 ' + " + Environment.NewLine +
                                              "       cast(WorkStateDays2 as varchar) + ' 例 ' + " + Environment.NewLine +
                                              "       cast(WorkStateDays3 as varchar) + ' 國' as DayTypeList " + Environment.NewLine +
                                              "  from DriverMonthDays " + Environment.NewLine +
                                              " where YYYYMM = '" + vYM + "' ";
                            string vDayTypeList = PF.GetValue(vConnStr, vTempStr, "DayTypeList");
                            vTempStr = "select [Name] from Custom where Types = 'O' and Code = 'A000' ";
                            string vCompanyName = PF.GetValue(vConnStr, vTempStr, "Name");
                            ReportDataSource rdsPrint = new ReportDataSource("DriverWorkStateListP", dtWSList);

                            rvPrint.LocalReport.DataSources.Clear();
                            rvPrint.LocalReport.ReportPath = @"Report\DriverWorkStateListP.rdlc";
                            rvPrint.LocalReport.DataSources.Add(rdsPrint);
                            rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "駕駛員排班狀況表"));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("DayTypeList", vDayTypeList));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_11", vESCDutyMember_11));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_13", vESCDutyMember_13));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_15", vESCDutyMember_15));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_16", vESCDutyMember_16));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_17", vESCDutyMember_17));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_18", vESCDutyMember_18));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_19", vESCDutyMember_19));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_20", vESCDutyMember_20));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_22", vESCDutyMember_22));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_23", vESCDutyMember_23));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ESCDutyMember_25", vESCDutyMember_25));
                            rvPrint.LocalReport.Refresh();
                            plReport.Visible = true;
                            plShowData.Visible = false;

                            string vSheetYMStr = eSheetYear.Text.Trim() + " 年 " + eSheetMonth.Text.Trim() + " 月";
                            string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "自 " + eDepNo_S.Text.Trim() + " 站至 " + eDepNo_E.Text.Trim() + " 站 " :
                                               ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() + " 站" :
                                               ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() + " 站" : "全部";
                            string vRecordNote = "預覽報表_駕駛員排班狀況表" + Environment.NewLine +
                                                 "DriverWorkStateList.aspx" + Environment.NewLine +
                                                 "查詢年月：" + vSheetYMStr + Environment.NewLine +
                                                 "站別：" + vDepNoStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                        }
                    }
                }
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            WorkStateListDataBind(GetSelStr(), "1");
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            WorkStateListDataBind(GetSelStr(), "2");
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        //根據繫結內容決定要不要把 GridView 欄位裡的文字改成紅色
        protected void gridDriverWorkStateList_RowDataBound(object sender, System.Web.UI.WebControls.GridViewRowEventArgs e)
        {
            for (int i = 0; i < gridDriverWorkStateList.Columns.Count; i++)
            {
                if ((i > 3) && (Array.IndexOf(vaTypeStr, e.Row.Cells[i].Text.Trim()) < 0) && (!Int32.TryParse(e.Row.Cells[i].Text.Trim(), out int vResultInt)))
                { //前四個欄位不處理，第一行是標題欄也不處理
                  //其他欄位如果是空字串或是 "平、例、國、休" 四個字之一也不處理，剩下的都把文字轉紅
                    e.Row.Cells[i].ForeColor = Color.Red;
                }
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
            plSearch.Visible = true;
            plShowData.Visible = true;
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            WorkStateListDataBind(GetSelStr(), "3");
        }
    }
}