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
    public partial class DriverWorkState : System.Web.UI.Page
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
                        eCalYear_Search.Text = (vToday.Month - 1 > 0) ? (vToday.Year - 1911).ToString() : (vToday.Year - 1912).ToString();
                        eCalMonth_Search.Text = (vToday.Month - 1 > 0) ? (vToday.Month - 1).ToString() : "12";
                        eDepNo_S_Search.Text = vLoginDepNo;
                        eDepName_S_Search.Text = (string)Session["DepName"];
                        /* 2021.07.29 修改為只要是總公司部門都不鎖單位
                        eDepNo_S_Search.Enabled = ((vLoginDepNo == "06") || (vLoginDepNo == "09") || (vLoginDepNo == "03") || (vLoginDepNo == "05"));
                        eDepName_S_Search.Visible = ((vLoginDepNo == "06") || (vLoginDepNo == "09") || (vLoginDepNo == "03") || (vLoginDepNo == "05"));
                        eDepNo_E_Search.Enabled = ((vLoginDepNo == "06") || (vLoginDepNo == "09") || (vLoginDepNo == "03") || (vLoginDepNo == "05"));
                        eDepName_E_Search.Visible = ((vLoginDepNo == "06") || (vLoginDepNo == "09") || (vLoginDepNo == "03") || (vLoginDepNo == "05"));
                        //*/
                        if (!int.TryParse(vLoginDepNo, out int vDepIndex))
                        {
                            vDepIndex = -1;
                        }
                        if ((vDepIndex > 0) && (vDepIndex < 11))
                        {
                            eDepNo_S_Search.Enabled = true;
                            //eDepName_S_Search.Visible = true;
                            eDepNo_E_Search.Enabled = true;
                            eDepName_E_Search.Visible = true;
                        }
                        else
                        {
                            eDepNo_S_Search.Enabled = false;
                            //eDepName_S_Search.Visible = false;
                            eDepNo_E_Search.Enabled = false;
                            eDepName_E_Search.Visible = false;
                        }
                        //=====================================================================================================================================
                    }
                    ErrorDataListBind();
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
        /// 取得 SQL 查詢字串
        /// </summary>
        /// <param name="fIndex">選擇範圍：0...全部資料   1...異常資料</param>
        /// <returns>回傳查詢字串</returns>
        private string GetSelectStr(int fIndex)
        {
            string vCalYear = (eCalYear_Search.Text.Trim() == "") ? DateTime.Today.Year.ToString("D4") :
                              (Int32.Parse(eCalYear_Search.Text.Trim()) < 1911) ? (Int32.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString("D4") :
                              eCalYear_Search.Text.Trim();
            string vCalMonth = Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2");
            string vReturnStr = "";
            string vCalYM = vCalYear + vCalMonth;
            string vCalDate_S = vCalYear + "/" + vCalMonth + "/01";
            string vCalDate_E = PF.GetMonthLastDay(DateTime.Parse(vCalDate_S), "B");
            string vWStr_Driver = ((eDriver_Search.Text.Trim() != "") && (fIndex == 1)) ? " and e.EmpNo = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   and e.DepNo between '" + eDepNo_S_Search.Text.Trim() + "' and '" + eDepNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? "   and e.DepNo = '" + eDepNo_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   and e.DepNo = '" + eDepNo_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            vReturnStr = "select t.DepNo, (select [Name] from Department where Department.DepNo = t.DepNo) as DepName, t.EmpNo, t.[Name], t.WorkType, " + Environment.NewLine +
                         "       (t.NormalDays + t.HoliDays + t.RuleHoliDays + t.NationHoliDays + t.ESCDays) as MonthDays_Cal, t.MonthDays, t.WorkDays, " + Environment.NewLine +
                         "       t.ESCDays, (t.AllowHoliDays + t.AllowRuleHoliDays + t.AllowNationHoliDays) as MonthAllowESCDays, " + Environment.NewLine +
                         "       NoneWorkDays, t.NormalDays, t.AllowNormalDays, t.HoliDays, t.AllowHoliDays, t.RuleHoliDays, " + Environment.NewLine +
                         "       t.AllowRuleHoliDays, t.NationHoliDays, t.AllowNationHoliDays, " + Environment.NewLine +
                         "       case when (t.HoliDays - t.AllowHoliDays) > 0 then(t.HoliDays - t.AllowHoliDays) else 0 end as HoliDayDiff, " + Environment.NewLine +
                         "       case when (t.RuleHoliDays - t.AllowRuleHoliDays) > 0 then(t.RuleHoliDays - t.AllowRuleHoliDays) else 0 end as RuleHoliDatDiff, " + Environment.NewLine +
                         "       case when (t.NationHoliDays - t.AllowNationHoliDays) > 0 then(t.NationHoliDays - t.AllowNationHoliDays) else 0 end as NationHoliDayDiff, " + Environment.NewLine +
                         "       case when ((t.HoliDays + t.RuleHoliDays + t.NationHoliDays + t.NoneWorkDays) <> (t.AllowHoliDays + t.AllowRuleHoliDays + t.AllowNationHoliDays + t.ESCDays)) " + Environment.NewLine +
                         "            then ((t.HoliDays + t.RuleHoliDays + t.NationHoliDays + t.NoneWorkDays) - (t.AllowHoliDays + t.AllowRuleHoliDays + t.AllowNationHoliDays + t.ESCDays)) " + Environment.NewLine +
                         "            else 0 end as NoneWorkDayDiff, " + Environment.NewLine +
                         "       case when ((t.NoneWorkDays + t.NormalDays + t.HoliDays + t.RuleHoliDays + t.NationHoliDays) <> t.MonthDays) " + Environment.NewLine +
                         "            then (t.NoneWorkDays + t.NormalDays + t.HoliDays + t.RuleHoliDays + t.NationHoliDays) - t.MonthDays " + Environment.NewLine +
                         "            else 0 end as MonthDayDiff, t.AssumeDay" + Environment.NewLine +
                         "  from ( " + Environment.NewLine +
                         "        select day(dateadd(m, 1, '" + vCalDate_S + "') - day('" + vCalDate_S + "')) MonthDays, " + Environment.NewLine +
                         "               e.WorkType, e.EmpNo, e.[Name], e.DepNo, " + Environment.NewLine +
                         "               (select count(AssignNo) from RunSheetA where RunSheetA.BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and RunSheetA.Driver = e.EmpNo) as WorkDays, " + Environment.NewLine +
                         "               day(dateadd(m, 1, '" + vCalDate_S + "') - day('" + vCalDate_S + "')) - (select count(AssignNo) from RunSheetA where RunSheetA.BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and RunSheetA.Driver = e.EmpNo) as NoneWorkDays, " + Environment.NewLine +
                         "               (select count(AssignNo) from RunSheetA where RunSheetA.BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and RunSheetA.Driver = e.EmpNo and RunSheetA.WORKSTATE = '0') as NormalDays, " + Environment.NewLine +
                         "               (select  WORKSTATEDAYS0 from DriverMonthDays where YYYYMM = '" + vCalYM + "') as AllowNormalDays, " + Environment.NewLine +
                         "               (select count(AssignNo) from RunSheetA where RunSheetA.BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and RunSheetA.Driver = e.EmpNo and RunSheetA.WORKSTATE = '1') as HoliDays, " + Environment.NewLine +
                         "               (select  WORKSTATEDAYS1 from DriverMonthDays where YYYYMM = '" + vCalYM + "') as AllowHoliDays, " + Environment.NewLine +
                         "               (select count(AssignNo) from RunSheetA where RunSheetA.BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and RunSheetA.Driver = e.EmpNo and RunSheetA.WORKSTATE = '2') as RuleHoliDays, " + Environment.NewLine +
                         "               (select  WORKSTATEDAYS2 from DriverMonthDays where YYYYMM = '" + vCalYM + "') as AllowRuleHoliDays, " + Environment.NewLine +
                         "               (select count(AssignNo) from RunSheetA where RunSheetA.BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and RunSheetA.Driver = e.EmpNo and RunSheetA.WORKSTATE = '3') as NationHoliDays, " + Environment.NewLine +
                         "               (select  WORKSTATEDAYS3 from DriverMonthDays where YYYYMM = '" + vCalYM + "') as AllowNationHoliDays, " + Environment.NewLine +
                         "               (select count(AssignNo) from RunSheetA where RunSheetA.BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and RunSheetA.Driver = e.EmpNo and RunSheetA.WORKSTATE <> '0') as TotalHoliDays, " + Environment.NewLine +
                         "               (select Count(Hours_T) from(select sum(Hours) Hours_T from ESCDUTY where RealDay between '" + vCalDate_S + "' and '" + vCalDate_E + "' and ESCDuty.ApplyMan = e.EmpNo group by RealDay) z where Hours_T >= 8) as ESCDays, " + Environment.NewLine +
                         "               case when AssumeDay between '" + vCalDate_S + "' and '" + vCalDate_E + "' then AssumeDay else null end as AssumeDay " + Environment.NewLine +
                         "          from Employee as e " + Environment.NewLine +
                         "         where e.Type = '20' " + Environment.NewLine + vWStr_Driver +
                         "           and (e.LeaveDay is null or e.LeaveDay >= '" + vCalDate_S + "') " + Environment.NewLine +
                         "           and e.AssumeDay <= '" + vCalDate_E + "' " + Environment.NewLine + vWStr_DepNo +
                         ") t ";
            if (fIndex == 1)
            {
                vReturnStr = vReturnStr + Environment.NewLine +
                             " where t.WorkDays > 0 " + Environment.NewLine +
                             "    and ((t.HoliDays > t.AllowHoliDays) or(t.RuleHoliDays > t.AllowRuleHoliDays) or(t.NationHoliDays > t.AllowNationHoliDays) " + Environment.NewLine +
                             "    or ((t.HoliDays + t.RuleHoliDays + t.NationHoliDays + t.NoneWorkDays) <> (t.AllowHoliDays + t.AllowRuleHoliDays + t.AllowNationHoliDays + t.ESCDays)) " + Environment.NewLine +
                             "    or (t.NoneWorkDays + t.NormalDays + t.HoliDays + t.RuleHoliDays + t.NationHoliDays) <> t.MonthDays)";
            }
            else
            {
                vReturnStr = vReturnStr + Environment.NewLine +
                             " where t.WorkDays > 0 ";
            }
            return vReturnStr;
        }

        private void ErrorDataListBind()
        {
            string vSelectStr = GetSelectStr(1);
            sdsDriverWorkStateErrorList.SelectCommand = "";
            sdsDriverWorkStateErrorList.SelectCommand = vSelectStr;
            gridDriverWorkStateErrorList.DataBind();
            lbRemark.Visible = (gridDriverWorkStateErrorList.Rows.Count > 0);
        }

        private void SaveExcel(string fFileName, string fSelectStr)
        {
            DateTime vAssumeDay;
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
            string vFieldValue = "";
            int vLinesNo = 0;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(fSelectStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(fFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "DepNo") ? "部門編號" :
                                      (drExcel.GetName(i) == "DepName") ? "部門" :
                                      (drExcel.GetName(i) == "EmpNo") ? "員工編號" :
                                      (drExcel.GetName(i) == "Name") ? "駕駛員姓名" :
                                      (drExcel.GetName(i) == "WorkType") ? "在職情況" :
                                      (drExcel.GetName(i) == "MonthDays_Cal") ? "憑單及請假天數" :
                                      (drExcel.GetName(i) == "MonthDays") ? "當月天數" :
                                      (drExcel.GetName(i) == "WorkDays") ? "憑單天數" :
                                      (drExcel.GetName(i) == "ESCDays") ? "請假天數" :
                                      (drExcel.GetName(i) == "MonthAllowESCDays") ? "當月可休天數" :
                                      (drExcel.GetName(i) == "NoneWorkDays") ? "憑單未出勤天數" :
                                      (drExcel.GetName(i) == "NormalDays") ? "憑單平常日天數" :
                                      (drExcel.GetName(i) == "AllowNormalDays") ? "當月上班日天數" :
                                      (drExcel.GetName(i) == "HoliDays") ? "休假上班天數" :
                                      (drExcel.GetName(i) == "AllowHoliDays") ? "當月應休假天數" :
                                      (drExcel.GetName(i) == "RuleHoliDays") ? "例假上班天數" :
                                      (drExcel.GetName(i) == "AllowRuleHoliDays") ? "當月應休例假天數" :
                                      (drExcel.GetName(i) == "NationHoliDays") ? "國定假日上班天數" :
                                      (drExcel.GetName(i) == "AllowNationHoliDays") ? "當月國定假日天數" :
                                      (drExcel.GetName(i) == "HoliDayDiff") ? "休假天數差異" :
                                      (drExcel.GetName(i) == "RuleHoliDatDiff") ? "例假天數差異" :
                                      (drExcel.GetName(i) == "NationHoliDayDiff") ? "國定假日天數差異" :
                                      (drExcel.GetName(i) == "NoneWorkDayDiff") ? "當月請休假加班天數差異" :
                                      (drExcel.GetName(i) == "MonthDayDiff") ? "當月總天數差異" :
                                      (drExcel.GetName(i) == "AssumeDay") ? "到職日期" : drExcel.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    //以下開始寫入資料
                    vLinesNo++;
                    while (drExcel.Read())
                    {
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if ((drExcel.GetName(i) == "DepNo") ||
                                (drExcel.GetName(i) == "DepName") ||
                                (drExcel.GetName(i) == "EmpNo") ||
                                (drExcel.GetName(i) == "Name") ||
                                (drExcel.GetName(i) == "WorkType"))
                            {
                                //字串欄位
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                            }
                            else if ((drExcel.GetName(i) == "AssumeDay") && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vAssumeDay = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vAssumeDay.Year) - 1911).ToString("D3") + "/" + vAssumeDay.ToString("MM/dd"));
                            }
                            else
                            {
                                //數值欄位
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                vFieldValue = (drExcel[i].ToString().Trim() != "") ? drExcel[i].ToString().Trim() : "0";
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(vFieldValue));
                            }
                            wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                        }
                        vLinesNo++;
                    }
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    wsExcel.GetRow(vLinesNo).CreateCell(0, CellType.String);
                    wsExcel.GetRow(vLinesNo).GetCell(0).SetCellValue("備註：新進、退休、離職、停職及未滿一個月之人員，請用人工核定！");
                    wsExcel.GetRow(vLinesNo).GetCell(0).CellStyle = csData;
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
                            string vRecordDepNoStr = ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? eDepNo_S_Search.Text.Trim() + "~" + eDepNo_E_Search.Text.Trim() :
                                                     ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? eDepNo_S_Search.Text.Trim() :
                                                     ((eDepNo_S_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? eDepNo_E_Search.Text.Trim() : "全部";
                            string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_" + fFileName + Environment.NewLine +
                                                 "DriverWorkState.aspx" + Environment.NewLine +
                                                 "行車日期：" + eCalYear_Search.Text.Trim() + "年" + eCalMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + fFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fFileName + ".xls"));
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

        protected void eDepNo_S_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDepNo_S_Search.Text.Trim();
            string vDepName_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_S_Search.Text = vDepNo_Temp;
            eDepName_S_Search.Text = vDepName_Temp;
        }

        protected void eDepNo_E_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDepNo_E_Search.Text.Trim();
            string vDepName_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_E_Search.Text = vDepNo_Temp;
            eDepName_E_Search.Text = vDepName_Temp;
        }

        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo = eDriver_Search.Text.Trim();
            string vEmpName = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' and LeaveDay is null";
            vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName + "' and LeaveDay is null";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eDriver_Search.Text = vEmpNo;
            eDriverName_Search.Text = vEmpName;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            ErrorDataListBind();

            string vRecordDepNoStr = ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? eDepNo_S_Search.Text.Trim() + "~" + eDepNo_E_Search.Text.Trim() :
                                     ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? eDepNo_S_Search.Text.Trim() :
                                     ((eDepNo_S_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? eDepNo_E_Search.Text.Trim() : "全部";
            string vRecordDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
            string vRecordNote = "查詢資料_駕駛員出勤狀況異常查核" + Environment.NewLine +
                                 "DriverWorkState.aspx" + Environment.NewLine +
                                 "行車日期：" + eCalYear_Search.Text.Trim() + "年" + eCalMonth_Search.Text.Trim() + "月" + Environment.NewLine +
                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                 "駕駛員：" + vRecordDriverStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
        }

        protected void bbAll2Excel_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelectStr(0);
            SaveExcel("駕駛員出勤狀況查核", vSelectStr);
        }

        protected void bbError2Excel_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelectStr(1);
            SaveExcel("駕駛員出勤狀況查核", vSelectStr);
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}