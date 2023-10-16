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
    public partial class RunSheetCheck : System.Web.UI.Page
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
                        eCheckYear.Text = (vToday.Month - 1 < 0) ? (vToday.Year - 1912).ToString() : (vToday.Year - 1911).ToString();
                        eCheckMonth.Text = (vToday.Month - 1 > 0) ? (vToday.Month - 1).ToString() : "12";
                        if ((vLoginDepNo == "03") || (vLoginDepNo == "06") || (vLoginDepNo == "09"))
                        {
                            eDepNo_S.Text = "";
                            eDepName_S.Text = "";
                            eDepNo_S.Enabled = true;
                            eDepNo_E.Text = "";
                            eDepName_E.Text = "";
                            eDepNo_E.Enabled = true;
                        }
                        else
                        {
                            eDepNo_S.Text = vLoginDepNo;
                            eDepName_S.Text = (string)Session["DepName"];
                            eDepNo_S.Enabled = false;
                            eDepNo_E.Text = "";
                            eDepName_E.Text = "";
                            eDepNo_E.Enabled = false;
                        }
                    }
                    ListDataBind();
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

        private string GetSelectStr()
        {
            string vResultStr = "";
            int vYear = Int32.Parse(eCheckYear.Text.Trim());
            int vMonth = Int32.Parse(eCheckMonth.Text.Trim());
            int vLastDay = DateTime.DaysInMonth(vYear, vMonth);
            DateTime vMonthFirstDate = new DateTime(vYear, vMonth, 1);
            DateTime vMonthLastDate = new DateTime(vYear, vMonth, vLastDay);
            string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "   and DepNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " :
                                 ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "   and DepNo = '" + eDepNo_S.Text.Trim() + "' " :
                                 ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "   and DepNo = '" + eDepNo_E.Text.Trim() + "' " : "";
            //2020.10.10 加入取回當天請假假別和時數
            vResultStr = "select ASSIGNNO, BUDATE, DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = ra.DEPNO)) AS DepName, " + Environment.NewLine +
                          "       DRIVER, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = ra.DRIVER)) AS DriverName, " + Environment.NewLine +
                          "       WORKHR, workmin, ActualKm, " + Environment.NewLine +
                          "       (select ClassTxt from DBDICB where ClassNo = ra.WORKSTATE and FKey = '行車記錄單A     runsheeta       WORKSTATE') WorkState, " + Environment.NewLine +
                          //"       case when (select top 1 SerialNo from ESCDuty where ESCDUty.RealDay = RunSheetA.BuDate and ESCDuty.ApplyMan = RunSheetA.Driver) is null then '' else 'V' end HasESCDuty, " + Environment.NewLine +
                          "       (select top 1 ResultStr = (stuff(( " + Environment.NewLine +
                          "               select ',' + (RTrim((select ClassTxt from DBDICB where FKey = '請假資料檔      ESCDUTY         ESCTYPE' and ClassNo = a.ESCType)) + '_' + cast(a.[Hours] as varchar) + '小時') " + Environment.NewLine +
                          "                 from ESCDuty a " + Environment.NewLine +
                          "                where a.RealDay = ra.BuDate and a.ApplyMan = ra.Driver " + Environment.NewLine +
                          "                  for xml path('')), 1, 1, '')) " + Environment.NewLine +
                          "          from ESCDuty e " + Environment.NewLine +
                          "         where e.RealDay = ra.BuDate and e.ApplyMan = ra.Driver) as HasESCDuty, " + Environment.NewLine +
                          "       cast(null as varchar) as Remark " + Environment.NewLine +
                          "  from RUNSHEETA ra " + Environment.NewLine +
                          " where BuDate between '" + PF.TransDateString(vMonthFirstDate, "B") + "' and '" + PF.TransDateString(vMonthLastDate, "B") + "' " + Environment.NewLine +
                          "   and ((WorkHR = 0 and WorkMin = 0) or (ActualKM = 0) or (WorkHR * 60 + WorkMin < 480)) " + Environment.NewLine +
                          vWStr_DepNo +
                          " union all " + Environment.NewLine +
                          "select ASSIGNNO, BUDATE, DEPNO, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = ra.DEPNO)) AS DepName, " + Environment.NewLine +
                          "       DRIVER, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = ra.DRIVER)) AS DriverName, " + Environment.NewLine +
                          "       WORKHR, workmin, ActualKm, " + Environment.NewLine +
                          "       (select ClassTxt from DBDICB where ClassNo = ra.WORKSTATE and FKey = '行車記錄單A     runsheeta       WORKSTATE') WorkState, " + Environment.NewLine +
                          //"       case when (select top 1 SerialNo from ESCDuty where ESCDUty.RealDay = RunSheetA.BuDate and ESCDuty.ApplyMan = RunSheetA.Driver) is null then '' else 'V' end HasESCDuty, " + Environment.NewLine +
                          "       (select top 1 ResultStr = (stuff(( " + Environment.NewLine +
                          "               select ',' + (RTrim((select ClassTxt from DBDICB where FKey = '請假資料檔      ESCDUTY         ESCTYPE' and ClassNo = a.ESCType)) + '_' + cast(a.[Hours] as varchar) + '小時') " + Environment.NewLine +
                          "                 from ESCDuty a " + Environment.NewLine +
                          "                where a.RealDay = ra.BuDate and a.ApplyMan = ra.Driver " + Environment.NewLine +
                          "                  for xml path('')), 1, 1, '')) " + Environment.NewLine +
                          "          from ESCDuty e " + Environment.NewLine +
                          "         where e.RealDay = ra.BuDate and e.ApplyMan = ra.Driver) as HasESCDuty, " + Environment.NewLine +
                          "       cast('明細公里數異常' as varchar) as Remark " + Environment.NewLine +
                          "  from RUNSHEETA ra " + Environment.NewLine +
                          " where BuDate between '" + PF.TransDateString(vMonthFirstDate, "B") + "' and '" + PF.TransDateString(vMonthLastDate, "B") + "' " + Environment.NewLine +
                          "   and AssignNo in (select AssignNo from RunsheetB where RunSheetB.ActualKM > 500 and RunSheetB.CARTYPE not in ('7', '8')) " + Environment.NewLine +
                          vWStr_DepNo;
            return vResultStr;
        }

        private void ListDataBind()
        {
            string vSelectStr = GetSelectStr();
            sdsRunSheetCheckList.SelectCommand = vSelectStr;
            gridRunSheetCheckList.DataBind();
        }

        private void SaveExcel(string fFileName, string fSelectStr)
        {
            DateTime vBudate;
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
                        vHeaderText = (drExcel.GetName(i) == "ASSIGNNO") ? "行車憑單單號" :
                                      (drExcel.GetName(i) == "BUDATE") ? "行車日期" :
                                      (drExcel.GetName(i) == "DEPNO") ? "站別編號" :
                                      (drExcel.GetName(i) == "DepName") ? "站別" :
                                      (drExcel.GetName(i) == "DRIVER") ? "駕駛員工號" :
                                      (drExcel.GetName(i) == "DriverName") ? "駕駛員" :
                                      (drExcel.GetName(i) == "WORKHR") ? "行車小時" :
                                      (drExcel.GetName(i) == "workmin") ? "行車分鐘" :
                                      (drExcel.GetName(i) == "ActualKm") ? "總公里數" :
                                      (drExcel.GetName(i) == "HasESCDuty") ? "當日有請假單" :
                                      (drExcel.GetName(i) == "WorkState") ? "出勤狀況" :
                                      (drExcel.GetName(i) == "Remark") ? "其他異常情況" : drExcel.GetName(i);
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
                            if ((drExcel.GetName(i) == "ASSIGNNO") ||
                                (drExcel.GetName(i) == "DEPNO") ||
                                (drExcel.GetName(i) == "DepName") ||
                                (drExcel.GetName(i) == "DRIVER") ||
                                (drExcel.GetName(i) == "DriverName") ||
                                (drExcel.GetName(i) == "HasESCDuty") ||
                                (drExcel.GetName(i) == "WorkState") ||
                                (drExcel.GetName(i) == "Remark"))
                            {
                                //字串欄位
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                            }
                            else if ((drExcel.GetName(i) == "BUDATE") && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBudate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBudate.Year) - 1911).ToString("D3") + "/" + vBudate.ToString("MM/dd"));
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
                            string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "自 " + eDepNo_S.Text.Trim() + " 站至 " + eDepNo_E.Text.Trim() + " 站" :
                                               ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepName_S.Text.Trim() :
                                               ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() : "全部";
                            string vCheckYMStr = eCheckYear.Text.Trim() + " 年 " + eCheckMonth.Text.Trim() + " 月";
                            string vRecordNote = "匯出檔案_行車憑單錯誤檢查" + Environment.NewLine +
                                                 "RunSheetCheck.aspx" + Environment.NewLine +
                                                 "站別：" + vDepNoStr + Environment.NewLine +
                                                 "檢查年月：" + vCheckYMStr;
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

        protected void eDepNo_S_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_S.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Departmwnt where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_S.Text = vDepNo_Temp;
            eDepName_S.Text = vDepName_Temp;
        }

        protected void eDepNo_E_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_E.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Departmwnt where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_E.Text = vDepNo_Temp;
            eDepName_E.Text = vDepName_Temp;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            ListDataBind();
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelectStr();
            SaveExcel("行車憑單錯誤檢查", vSelectStr);
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}