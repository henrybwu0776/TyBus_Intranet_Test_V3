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

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverMonthNoneWorkList : System.Web.UI.Page
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
                        eCalYear.Text = (vToday.Year - 1911).ToString().Trim();
                        eCalMonth.Text = vToday.AddMonths(-1).Month.ToString().Trim();
                        plShowData.Visible = true;
                        plPrint.Visible = false;
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
            string vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "                 and e.DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EMpNo = (eEmpNo.Text.Trim() != "") ? "                 and e.EmpNo = '" + eEmpNo.Text.Trim() + "' " + Environment.NewLine : "";
            DateTime vStartDate = DateTime.Parse(eCalYear.Text.Trim() + "/" + eCalMonth.Text.Trim() + "/01");
            string vStartDate_Str = PF.GetMonthFirstDay(vStartDate, "B");
            String vEndDate_Str = PF.GetMonthLastDay(vStartDate, "B");
            string vResultStr = "select z.DepNo, (select [Name] from Department where DepNo = z.DepNo) DepName, " + Environment.NewLine +
                                "       z.EmpNo, z.[Name], z.WorkType, z.LeaveDay, z.BeginDay, z.StopDay, " + Environment.NewLine +
                                "       sum(z.Hours01) / 8 TDay01, sum(z.Hours03) / 8 TDay03, sum(z.Hours04) / 8 TDay04, sum(z.Hours05) / 8 TDay05, " + Environment.NewLine +
                                "       sum(z.OtherHours) / 8 TOtherDays, " + Environment.NewLine +
                                "       case when sum(z.Hours01 +z.Hours03 + z.Hours04 + z.Hours05 + z.OtherHours) = 0 then '無請假記錄' " + Environment.NewLine +
                                "            else cast(sum(z.Hours01 + z.Hours03 + z.Hours04 + z.Hours05 + z.OtherHours) / 8 as varchar) end TotalESCDays " + Environment.NewLine +
                                "  from ( " + Environment.NewLine +
                                "       select t.DepNo, t.EmpNo, t.[Name], t.LeaveDay, t.BeginDay, t.StopDay, t.WorkType, " + Environment.NewLine +
                                "              case when t.ESCType = '01' then t.[Hours] else 0 end Hours01, " + Environment.NewLine +
                                "              case when t.ESCType = '03' then t.[Hours] else 0 end Hours03, " + Environment.NewLine +
                                "              case when t.ESCType = '04' then t.[Hours] else 0 end Hours04, " + Environment.NewLine +
                                "              case when t.ESCType = '05' then t.[Hours] else 0 end Hours05, " + Environment.NewLine +
                                "              case when t.ESCType not in ('01','03','04','05') then t.[Hours] else 0 end OtherHours " + Environment.NewLine +
                                "         from ( " + Environment.NewLine +
                                "              select e.DepNo, e.EmpNo, e.[Name], e.WorkType, e.LeaveDay, e.BeginDay, e.StopDay, d.ESCType, d.[Hours] " + Environment.NewLine +
                                "                from Employee e left join ESCDuty d on d.ApplyMan = e.EmpNo and d.RealDay Between '" + vStartDate_Str + "' and '" + vEndDate_Str + "' " + Environment.NewLine +
                                "               where e.LeaveDay >= '" + vStartDate_Str + "' " + Environment.NewLine +
                                "                 and e.[Type] = '20' " + Environment.NewLine +
                                "                 and e.EmpNo not in (select distinct Driver from RunSheetA where BuDate between '" + vStartDate_Str + "' and '" + vEndDate_Str + "') " + Environment.NewLine +
                                "                 and e.DepNo in (select DepNo from Department where IsStation = 'V' and DepGroup in ('1','2') and patindex('Z%', DepNo) = 0) " + Environment.NewLine +
                                vWStr_DepNo + vWStr_EMpNo +
                                "               union all " + Environment.NewLine +
                                "              select e.DepNo, e.EmpNo, e.[Name], e.WorkType, e.LeaveDay, e.BeginDay, e.StopDay, d.ESCType, d.[Hours] " + Environment.NewLine +
                                "                from Employee e left join ESCDuty d on d.ApplyMan = e.EmpNo and d.RealDay Between '" + vStartDate_Str + "' and '" + vEndDate_Str + "' " + Environment.NewLine +
                                "               where (isnull(e.BeginDay, '') <> '' and (e.StopDay is null or e.StopDay >= '" + vStartDate_Str + "')) " + Environment.NewLine +
                                "                 and e.[Type] = '20' " + Environment.NewLine +
                                "                 and e.EmpNo not in (select distinct Driver from RunSheetA where BuDate between '" + vStartDate_Str + "' and '" + vEndDate_Str + "') " + Environment.NewLine +
                                "                 and e.DepNo in (select DepNo from Department where IsStation = 'V' and DepGroup in ('1','2') and patindex('Z%', DepNo) = 0) " + Environment.NewLine +
                                vWStr_DepNo + vWStr_EMpNo +
                                "               union all " + Environment.NewLine +
                                "              select e.DepNo, e.EmpNo, e.[Name], e.WorkType, e.LeaveDay, e.BeginDay, e.StopDay, d.ESCType, d.[Hours] " + Environment.NewLine +
                                "                from Employee e left join ESCDuty d on d.ApplyMan = e.EmpNo and d.RealDay Between '" + vStartDate_Str + "' and '" + vEndDate_Str + "' " + Environment.NewLine +
                                "               where e.BeginDay is null " + Environment.NewLine +
                                "                 and e.LeaveDay is null " + Environment.NewLine +
                                "                 and e.[Type] = '20' " + Environment.NewLine +
                                "                 and e.EmpNo not in (select distinct Driver from RunSheetA where BuDate between '" + vStartDate_Str + "' and '" + vEndDate_Str + "') " + Environment.NewLine +
                                "                 and e.DepNo in (select DepNo from Department where IsStation = 'V' and DepGroup in ('1','2') and patindex('Z%', DepNo) = 0) " + Environment.NewLine +
                                vWStr_DepNo + vWStr_EMpNo +
                                "             ) as t " + Environment.NewLine +
                                "       ) z " + Environment.NewLine +
                                " group by z.DepNo, z.EmpNo, z.[Name], z.LeaveDay, z.BeginDay, z.StopDay, z.WorkType " + Environment.NewLine +
                                " order by z.EmpNo";
            return vResultStr;
        }

        private void GridDataBind()
        {
            string vSelStr = GetSelectStr();
            sdsShowData.SelectCommand = "";
            sdsShowData.SelectCommand = vSelStr;
            gridShowData.DataBind();
        }

        protected void eDepNo_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = eDepNo.Text.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo.Text = vDepNo_Temp.Trim();
            eDepName.Text = vDepName_Temp.Trim();
        }

        protected void eEmpNo_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo_Temp = eEmpNo.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp.Trim() + "' ";
            string vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = eEmpNo.Text.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Temp.Trim() + "' order by isnull(LeaveDay,'9999/12/31') DESC";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eEmpNo.Text = vEmpNo_Temp.Trim();
            eEmpName.Text = vEmpName_Temp.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            GridDataBind();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vSelStr = GetSelectStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                try
                {
                    SqlDataAdapter daPrint = new SqlDataAdapter(vSelStr, connPrint);
                    connPrint.Open();
                    DataTable dtPrint = new DataTable();
                    daPrint.Fill(dtPrint);
                    if (dtPrint.Rows.Count > 0)
                    {
                        string vReportName = (eDepNo.Text.Trim() != "") ?
                                             eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月 [" + eDepName.Text.Trim() + "] 全月未出勤駕駛員明細" :
                                             eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月 全月未出勤駕駛員明細";
                        ReportDataSource rdsPrint = new ReportDataSource("DriverMonthNoneWorkListP", dtPrint);

                        rvPrint.LocalReport.DataSources.Clear();
                        rvPrint.LocalReport.ReportPath = @"Report\DriverMonthNoneWorkListP.rdlc";
                        rvPrint.LocalReport.DataSources.Add(rdsPrint);
                        rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                        rvPrint.LocalReport.Refresh();
                        plShowData.Visible = false;
                        plPrint.Visible = true;

                        string vRecordDepNoStr = (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : "全部";
                        string vRecordDriverStr = (eEmpNo.Text.Trim() != "") ? eEmpNo.Text.Trim() : "全部";
                        string vRecordNote = "列印資料_全月未出勤駕駛員查詢" + Environment.NewLine +
                                             "DriverMonthNoneWorkList.aspx" + Environment.NewLine +
                                             "站別：" + vRecordDepNoStr + Environment.NewLine +
                                             "駕駛員：" + vRecordDriverStr + Environment.NewLine +
                                             "統計年月：" + eCalYear.Text.Trim() + "年" + eDepName.Text.Trim() + "月";
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
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

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelStr = GetSelectStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
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
            string vFileName = (eDepNo.Text.Trim() != "") ?
                               eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月 [" + eDepName.Text.Trim() + "] 全月未出勤駕駛員明細" :
                               eCalYear.Text.Trim() + "年" + eCalMonth.Text.Trim() + "月 全月未出勤駕駛員明細";
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
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "DEPNO") ? "" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "單位" :
                                      (drExcel.GetName(i).ToUpper() == "EMPNO") ? "工號" :
                                      (drExcel.GetName(i).ToUpper() == "NAME") ? "姓名" :
                                      (drExcel.GetName(i).ToUpper() == "WORKTYPE") ? "在職狀況" :
                                      (drExcel.GetName(i).ToUpper() == "LEAVEDAY") ? "離職日" :
                                      (drExcel.GetName(i).ToUpper() == "BEGINDAY") ? "留停起" :
                                      (drExcel.GetName(i).ToUpper() == "STOPDAY") ? "留停迄" :
                                      (drExcel.GetName(i).ToUpper() == "TDAY01") ? "特休" :
                                      (drExcel.GetName(i).ToUpper() == "TDAY03") ? "公傷假" :
                                      (drExcel.GetName(i).ToUpper() == "TDAY04") ? "事假" :
                                      (drExcel.GetName(i).ToUpper() == "TDAY05") ? "病假" :
                                      (drExcel.GetName(i).ToUpper() == "TOTHERDAYS") ? "其他假別" :
                                      (drExcel.GetName(i).ToUpper() == "TOTALESCDAYS") ? "總請假天數" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i).ToUpper() == "LEAVEDAY") ||
                                 (drExcel.GetName(i).ToUpper() == "BEGINDAY") ||
                                 (drExcel.GetName(i).ToUpper() == "STOPDAY")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                            }
                            else if (((drExcel.GetName(i).ToUpper() == "TDAY01") ||
                                      (drExcel.GetName(i).ToUpper() == "TDAY03") ||
                                      (drExcel.GetName(i).ToUpper() == "TDAY04") ||
                                      (drExcel.GetName(i).ToUpper() == "TDAY05") ||
                                      (drExcel.GetName(i).ToUpper() == "TOTHERDAYS")) && (drExcel[i].ToString() != ""))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
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
                            string vRecordDepNoStr = (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : "全部";
                            string vRecordDriverStr = (eEmpNo.Text.Trim() != "") ? eEmpNo.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_全月未出勤駕駛員查詢" + Environment.NewLine +
                                                 "DriverMonthNoneWorkList.aspx" + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vRecordDriverStr + Environment.NewLine +
                                                 "統計年月：" + eCalYear.Text.Trim() + "年" + eDepName.Text.Trim() + "月";
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

        protected void bbCLoseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plShowData.Visible = true;
        }
    }
}