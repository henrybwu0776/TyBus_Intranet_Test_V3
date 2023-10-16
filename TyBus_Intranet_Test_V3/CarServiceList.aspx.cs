using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class CarServiceList : Page
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
                        /* 變更站別過濾方式
                        string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' and isnull(InSHReport, 'X') = 'V' ";
                        string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                        if (vDepName_Temp != "")
                        {
                            eDepNo_Search.Text = vDepNo;
                            ddlDepNo_Search.SelectedIndex = ddlDepNo_Search.Items.IndexOf(ddlDepNo_Search.Items.FindByValue(vDepNo));
                            ddlDepNo_Search.Enabled = false;
                        }
                        else
                        {
                            eDepNo_Search.Text = "";
                            ddlDepNo_Search.SelectedIndex = 0;
                            ddlDepNo_Search.Enabled = true;
                        }
                        //*/
                        //改從 "員工多單位設定" 這個 TABLE 抓資料
                        string vSQLStr_Temp = "";
                        if (Int32.Parse(vLoginDepNo) <= 10)
                        {
                            vSQLStr_Temp = "SELECT DEPNO, DEPNO + '-' + NAME AS DepName FROM DEPARTMENT WHERE (ISNULL(InSHReport, 'X') = 'V')";
                        }
                        else
                        {
                            vSQLStr_Temp = "SELECT d.DEPNO, d.DEPNO + '-' + d.NAME AS DepName " + Environment.NewLine +
                                           "  FROM DEPARTMENT d " + Environment.NewLine +
                                           " WHERE (ISNULL(InSHReport, 'X') = 'V') " + Environment.NewLine +
                                           "   AND d.DepNo in (SELECT DepNo FROM EmployeeDepNo e where e.EmpNo = '" + vLoginID + "' and isnull(e.IsUsed, 0) = 1)";
                        }

                        sdsDepList.SelectCommand = vSQLStr_Temp;
                        ddlDepNo_Search.DataBind();

                        //using (SqlConnection connTemp=new SqlConnection(vConnStr))
                        //{
                        //    SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        //    connTemp.Open();
                        //    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        //    if (drTemp.HasRows)
                        //    {
                        //        ddlDepNo_Search.Items.Clear();
                        //        while (drTemp.Read())
                        //        {
                        //            ddlDepNo_Search.Items.Add(new ListItem(drTemp[1].ToString().Trim(), drTemp[0].ToString().Trim()));
                        //        }
                        //    }
                        //}

                        if (ddlDepNo_Search.Items.Count > 0)
                        {
                            ddlDepNo_Search.SelectedIndex = 0;
                            eDepNo_Search.Text = ddlDepNo_Search.Items[0].Value;
                        }

                        //2020.03.19 新增 2020.04.17 修改，不是修護人員看不到 "計算預排" 和 "預排關帳" 這兩個按鈕
                        if ((vLoginDepNo == "09") || (vLoginEmpType == "10"))
                        {
                            bbGetList.Visible = true; //2020.03.19
                            bbStopCalYM.Visible = true; //2020.04.17
                        }
                        else
                        {
                            bbGetList.Visible = false; //2020.03.19
                            bbStopCalYM.Visible = false; //2020.04.17
                        }

                        if ((vLoginDepNo == "09") || ((vLoginDepNo == "30") && ((vLoginTitle == "130") || (vLoginTitle == "230"))))
                        {
                            bbReOpenYM.Visible = true;
                        }
                        else
                        {
                            bbReOpenYM.Visible = false;
                        }
                        //==============================================================================================================================
                        eCalYear_Search.Text = (DateTime.Today.AddMonths(1).Year - 1911).ToString();
                        eCalMonth_Search.Text = DateTime.Today.AddMonths(1).Month.ToString();
                        eSecondServiceKM_Search.Text = "5500";
                        eThirdServiceKM_Search.Text = "35000";
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

        private string GetCarListStr(string fShowDate)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vWStr_CalYM = ((eCalYear_Search.Text.Trim() != "") && (eCalMonth_Search.Text.Trim() != "")) ?
                                 " where ServicePreNo like '" + eCalYear_Search.Text.Trim() + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2") + "%' " + Environment.NewLine :
                                 " where ServicePreNo like '" + (DateTime.Today.AddMonths(-1).Year - 1911).ToString() + DateTime.Today.AddMonths(-1).Month.ToString("D2") + "%' " + Environment.NewLine;
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine :
                                 (ddlDepNo_Search.SelectedValue != "") ? "   and DepNo = '" + ddlDepNo_Search.SelectedValue.ToString().Trim() + "' " + Environment.NewLine : "";
            string vWStr_ServiceDate = (fShowDate != "") ? "   and ServiceDate = '" + fShowDate + "' " + Environment.NewLine : "";
            string vSelStr = "select c.Car_ID, c.ServiceDate, " + Environment.NewLine +
                             "       (select ClassTXT from DBDICB where ClassNo = c.ServiceType and FKey = '工作單A         FixworkA        SERVICE') ServiceType_C " + Environment.NewLine +
                             "  from CarServicePre c" + Environment.NewLine +
                             vWStr_CalYM + vWStr_DepNo + vWStr_ServiceDate +
                             " order by c.DepNo, c.ServiceDate, c.Car_ID ";
            return vSelStr;
        }

        private string GetSelectStr(string fShowDate)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vWStr_CalYM = ((eCalYear_Search.Text.Trim() != "") && (eCalMonth_Search.Text.Trim() != "")) ?
                                 " where ServicePreNo like '" + eCalYear_Search.Text.Trim() + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2") + "%' " + Environment.NewLine :
                                 " where ServicePreNo like '" + (DateTime.Today.AddMonths(-1).Year - 1911).ToString() + DateTime.Today.AddMonths(-1).Month.ToString("D2") + "%' " + Environment.NewLine;
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine :
                                 (ddlDepNo_Search.SelectedValue != "") ? "   and DepNo = '" + ddlDepNo_Search.SelectedValue.ToString().Trim() + "' " + Environment.NewLine : "";
            string vWStr_ServiceDate = (fShowDate != "") ? "   and ServiceDate = '" + fShowDate + "' " + Environment.NewLine : "";
            string vSelStr = "select c.ServicePreNo, c.DepNo, (select [Name] from Department where DepNo = c.DepNo) DepName, " + Environment.NewLine +
                             "       c.Car_ID, c.Driver, (select [Name] from Employee where EmpNo = c.Driver) DriverName, " + Environment.NewLine +
                             "       c.ServiceType, (select ClassTXT from DBDICB where ClassNo = c.ServiceType and FKey = '工作單A         FixworkA        SERVICE') ServiceType_C, " + Environment.NewLine +
                             "       c.LastDate, c.ServiceDate, c.Remark, " + Environment.NewLine +
                             "       c.BuMan, (select [Name] from Employee where EmpNo = c.BuMan) BuManName, c.BuDate, " + Environment.NewLine +
                             "       c.ModifyMan, (select [Name] from Employee where EmpNo = c.ModifyMan) ModifyManName, c.ModifyDate " + Environment.NewLine +
                             "  from CarServicePre c" + Environment.NewLine +
                             vWStr_CalYM + vWStr_DepNo + vWStr_ServiceDate +
                             " order by c.DepNo, c.ServiceDate, c.Car_ID ";
            return vSelStr;
        }

        protected void ddlDepNo_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDepNo_Search.Text = ddlDepNo_Search.SelectedValue.Trim();
        }

        /// <summary>
        /// 匯出預排資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExportData_Click(object sender, EventArgs e)
        {
            if ((vLoginDepNo != "09") && (eDepNo_Search.Text == ""))
            {
                eDepNo_Search.Text = ddlDepNo_Search.Items[0].Value.Trim();
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
            string vSelectStr = GetSelectStr("");
            string vFileName = "桃園客運車輛_" + eCalYear_Search.Text.Trim() + "_年_" + eCalMonth_Search.Text.Trim() + "_月_二_三級保養預排清冊";
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
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "ServicePreNo") ? "預排單號" :
                                      (drExcel.GetName(i) == "DepNo") ? "" :
                                      (drExcel.GetName(i) == "DepName") ? "場站名" :
                                      (drExcel.GetName(i) == "Car_ID") ? "車號" :
                                      (drExcel.GetName(i) == "Driver") ? "" :
                                      (drExcel.GetName(i) == "DriverName") ? "駕駛員姓名" :
                                      (drExcel.GetName(i) == "ServiceType") ? "" :
                                      (drExcel.GetName(i) == "ServiceType_C") ? "保養級別" :
                                      (drExcel.GetName(i) == "LastDate") ? "上次保養日" :
                                      (drExcel.GetName(i) == "ServiceDate") ? "預排保養日" :
                                      (drExcel.GetName(i) == "Remark") ? "備註" :
                                      (drExcel.GetName(i) == "BuDate") ? "建檔日期" :
                                      (drExcel.GetName(i) == "BuMan") ? "" :
                                      (drExcel.GetName(i) == "BuManName") ? "建檔人" :
                                      (drExcel.GetName(i) == "ModifyDate") ? "異動日期" :
                                      (drExcel.GetName(i) == "ModifyMan") ? "" :
                                      (drExcel.GetName(i) == "ModifyManName") ? "異動人" : drExcel.GetName(i);
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    wsExcel.GetRow(vLinesNo).CreateCell(drExcel.FieldCount).SetCellValue("實際保養日");
                    wsExcel.GetRow(vLinesNo).GetCell(drExcel.FieldCount).CellStyle = csTitle;
                    vLinesNo++;
                    while (drExcel.Read())
                    {
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if (((drExcel.GetName(i) == "LastDate") ||
                                 (drExcel.GetName(i) == "ServiceDate") ||
                                 (drExcel.GetName(i) == "BuDate") ||
                                 (drExcel.GetName(i) == "ModifyDate")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
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
                            string vCalYMStr = ((eCalYear_Search.Text.Trim() != "") && (eCalMonth_Search.Text.Trim() != "")) ?
                                               eCalYear_Search.Text.Trim() + " 年 " + eCalMonth_Search.Text.Trim() + " 月" + Environment.NewLine :
                                               (DateTime.Today.AddMonths(-1).Year - 1911).ToString() + " 年 " + DateTime.Today.AddMonths(-1).Month.ToString("D2") + " 月" + Environment.NewLine;
                            string vCalDepNo = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() :
                                               (ddlDepNo_Search.SelectedValue != "") ? ddlDepNo_Search.SelectedValue.ToString().Trim() :
                                               "全部站別";
                            string vRecordNote = "匯出檔案_車輛保養預排" + Environment.NewLine +
                                                 "CarServiceList.aspx" + Environment.NewLine +
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

        /// <summary>
        /// 計算預排資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbGetList_Click(object sender, EventArgs e)
        {
            if ((vLoginDepNo != "09") && (eDepNo_Search.Text == ""))
            {
                eDepNo_Search.Text = ddlDepNo_Search.Items[0].Value.Trim();
            }
            string vControlItem = "a_ServicePreClose" + eDepNo_Search.Text.Trim();
            if ((eCalYear_Search.Text.Trim() != "") && (eCalMonth_Search.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCheckDateStr = (Int32.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString("D4") + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2") + "01";

                //判斷是否已經關帳
                if ((vLoginDepNo != "09") && (PF.DateIsClosed(vConnStr, vControlItem, vCheckDateStr) == true))
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('指定要計算的保養年月已經關帳！')");
                    Response.Write("</" + "Script>");
                    eCalYear_Search.Focus();
                }
                else
                {
                    string vCalYM = eCalYear_Search.Text.Trim() + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2");
                    string vSQLStr_Temp = "select count(ServicePreNo) RCount from CarServicePre where ServicePreNo like '" + vCalYM + "%' and DepNo = '" + eDepNo_Search.Text.Trim() + "' ";
                    int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount"));
                    if (vRCount > 0)
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('系統已有預排資料，不允許再次產生！')");
                        Response.Write("</" + "Script>");
                    }
                    else
                    {
                        //要計算預排件月的起迄日期
                        DateTime vCalDate_Start = DateTime.Parse(eCalYear_Search.Text.Trim() + "/" + eCalMonth_Search.Text.Trim() + "/01");
                        DateTime vCalDate_End = DateTime.Parse(eCalYear_Search.Text.Trim() + "/" + eCalMonth_Search.Text.Trim() + "/" + PF.GetMonthDays(vCalDate_Start).ToString());

                        //取回維修記錄的起迄日期
                        DateTime vStartDate = (vCalDate_Start > DateTime.Today) ? DateTime.Today.AddMonths(-1) : vCalDate_Start.AddMonths(-1);
                        string vStartDate_Str = vStartDate.Year.ToString() + "/" + vStartDate.ToString("MM/dd");
                        DateTime vEndDate = (vCalDate_End > DateTime.Today) ? DateTime.Today : vCalDate_End.AddMonths(-1);
                        string vEndDate_Str = vEndDate.Year.ToString() + "/" + vEndDate.ToString("MM/dd");

                        //決定要計算的站別
                        string vDepNo_Temp = (eDepNo_Search.Text.Trim() != "") ? " and DepNo = '" + eDepNo_Search.Text.Trim() + "' " : "";
                        string vCompanyNo_Temp = (eDepNo_Search.Text.Trim() != "") ? "		   and CompanyNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";

                        //取回維修記錄
                        /* 2020.02.10 重新改寫法
                        string vSQLStr = "select Car_ID, sum(isnull(ActualKM,0)) TotalKM, count(distinct AssignNo) RCount, " + Environment.NewLine +
                                         "       (sum(isnull(ActualKM, 0)) / count(distinct AssignNo)) AvgKM " + Environment.NewLine +
                                         "  from RunSheetB " + Environment.NewLine +
                                         " where AssignNo in (select AssignNo from RunSheetA where BuDate between '" + vStartDate_Str + "' and '" + vEndDate_Str + "') " + Environment.NewLine +
                                         "   and isnull(ReduceReason, '') = '' " + Environment.NewLine +
                                         " group by Car_ID " + Environment.NewLine +
                                         " union all " + Environment.NewLine +
                                         "select Car_ID, cast(0 as float) TotalKM, cast(0 as int) RCount, cast(0 as float) AvgKM " + Environment.NewLine +
                                         "  from Car_InfoA " + Environment.NewLine +
                                         " where Tran_Type = '1' and LEFT(Point, 1) <> 'X' " + Environment.NewLine +
                                         "   and Car_ID NOT IN ( " + Environment.NewLine +
                                         "                      select distinct Car_ID " + Environment.NewLine +
                                         "                        from RunSheetB " + Environment.NewLine +
                                         "                       where AssignNo in ( " + Environment.NewLine +
                                         "                                          select AssignNo from RunSheetA " + Environment.NewLine +
                                         "                                           where BuDate between '" + vStartDate_Str + "' and '" + vEndDate_Str + "') " + Environment.NewLine +
                                         "                         and isnull(ReduceReason, '') = '') " + Environment.NewLine +
                                         " order by Car_ID ";
                        //*/
                        string vSQLStr = "select Car_ID, sum(t.TotalKM) TotalKM, sum(t.RCount) RCount, " + Environment.NewLine +
                                         "       case when sum(t.RCount) <> 0 then sum(t.TotalKM) / sum(t.RCount) else 0 end AvgKM " + Environment.NewLine +
                                         "  from ( " + Environment.NewLine +
                                         "        select Car_ID, sum(isnull(ActualKM, 0)) TotalKM, count(distinct AssignNo) RCount " + Environment.NewLine +
                                         "          from RunSheetB " + Environment.NewLine +
                                         "         where AssignNo in (select AssignNo from RunSheetA where BuDate between '" + vStartDate_Str + "' and '" + vEndDate_Str + "' " + vDepNo_Temp + ") " + Environment.NewLine +
                                         "           and isnull(ReduceReason, '') = '' " + Environment.NewLine +
                                         "         group by Car_ID " + Environment.NewLine +
                                         "         union all " + Environment.NewLine +
                                         "        select Car_ID, cast(0 as float) TotalKM, cast(0 as int) RCount " + Environment.NewLine +
                                         "          from Car_InfoA " + Environment.NewLine +
                                         "         where Tran_Type = '1' and LEFT(Point, 1) <> 'X' " + Environment.NewLine + vCompanyNo_Temp +
                                         ") t " + Environment.NewLine +
                                         " Group by t.Car_ID " + Environment.NewLine +
                                         " order by t.Car_ID ";
                        string vCarID = "";
                        string vDriver = "";
                        string vDepNo_Car = "";
                        int vTotalRunDays = 0;
                        int vNewIndex = 0;
                        double vDayAvgKM = 0.0;
                        string vGetTempData = "";
                        string vTempStr = "";
                        string vServiceType = "";
                        string vExecStr = "";
                        string vClearDataStr = "";
                        DateTime vLastDate_Second;
                        DateTime vLastDate_Third;
                        //取回二級保養間隔里程 (預設 5500 KM )
                        int vSecondServiceKM = (eSecondServiceKM_Search.Text.Trim() != "") ? Int32.Parse(eSecondServiceKM_Search.Text.Trim()) : 5500;
                        int vSecondServiceDays = 0;
                        //取回三級保養間隔里程 (預設 35000 KM )
                        int vThirdServiceKM = (eThirdServiceKM_Search.Text.Trim() != "") ? Int32.Parse(eThirdServiceKM_Search.Text.Trim()) : 35000;
                        int vThirdServiceDays = 0;
                        //日期差距
                        int vDateDiff = 0;
                        //取回保養次數的語法
                        string vSQLStr_GetRCount = "";
                        int vCaseCount = 0;
                        string vRemark = "";

                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                //先取得每車的行駛天數跟日平均里程
                                vCarID = drTemp["Car_ID"].ToString().Trim();

                                if (vCarID != "(暫)")
                                {
                                    vGetTempData = "select (isnull(Driver, '') + ',' + isnull(CompanyNo,'')) TempData from Car_InfoA where Car_ID = '" + vCarID + "' ";
                                    vTempStr = PF.GetValue(vConnStr, vGetTempData, "TempData");
                                    string[] vTempArray = vTempStr.Split(',');
                                    vDriver = (vTempStr != "") ? "'" + vTempArray[0].Trim() + "'" : "NULL";
                                    vDepNo_Car = (vTempStr != "") ? "'" + vTempArray[1].Trim() + "'" : "NULL";
                                    vTotalRunDays = Int32.Parse(drTemp["RCount"].ToString().Trim());
                                    vDayAvgKM = double.Parse(drTemp["AvgKM"].ToString().Trim());

                                    vSecondServiceDays = (vDayAvgKM != 0) ? (int)(vSecondServiceKM / vDayAvgKM) : 20; //二級保養間隔天數，預設20天
                                    vThirdServiceDays = (vDayAvgKM != 0) ? (int)(vThirdServiceKM / vDayAvgKM) : 140; //三級保養間隔天數，預設140天
                                    vCalDate_Start = DateTime.Parse(eCalYear_Search.Text.Trim() + "/" + eCalMonth_Search.Text.Trim() + "/01");
                                    vCalDate_End = DateTime.Parse(eCalYear_Search.Text.Trim() + "/" + eCalMonth_Search.Text.Trim() + "/" + PF.GetMonthDays(vCalDate_Start).ToString());

                                    //先清空該車在計算年月範圍內已有的預排資料
                                    vClearDataStr = "delete CarServicePre where Car_ID = '" + vCarID + "' and ServiceDate between '" + vCalDate_Start.Year.ToString() + "/" + vCalDate_Start.ToString("MM/dd") + "' and '" + vCalDate_End.Year.ToString() + "/" + vCalDate_End.ToString("MM/dd") + "' ";
                                    PF.ExecSQL(vConnStr, vClearDataStr);

                                    //取得最後一次保養的資料
                                    vGetTempData = "select top 1 (cast((Year(ServiceDate) - 1911) as varchar) + ', ' + cast(Month(ServiceDate) as varchar) + ', ' + cast(Day(ServiceDate) as varchar) + ', ' + ServiceType) TempData " + Environment.NewLine +
                                                   "  from ( " + Environment.NewLine +
                                                   "        select FixDateOut as ServiceDate, [Service] as ServiceType " + Environment.NewLine +
                                                   "          from FixWorkA " + Environment.NewLine +
                                                   "         where Car_ID = '" + vCarID + "' " + Environment.NewLine +
                                                   "           and [Service] in ('1', '2') " + Environment.NewLine +
                                                   "         union all " + Environment.NewLine +
                                                   "        select ServiceDate, ServiceType " + Environment.NewLine +
                                                   "          from CarServicePre " + Environment.NewLine +
                                                   "         where Car_ID = '" + vCarID + "' " + Environment.NewLine +
                                                   "           and ServiceDate >= '" + vEndDate_Str + "' " + Environment.NewLine +
                                                   "           and isnull(IsPassed, 'X') = 'X' " + Environment.NewLine +
                                                   "       ) t order by ServiceDate DESC ";
                                    vTempStr = PF.GetValue(vConnStr, vGetTempData, "TempData");
                                    Array.Clear(vTempArray, 0, vTempArray.Length); //先清空前一次存入陣列的內容
                                    vTempArray = vTempStr.Split(','); //把從資料庫取回來的資料用按照逗號切割存入陣列
                                    string vGetLicDateStr = "";
                                    if (vTempStr != "")
                                    {
                                        vLastDate_Second = DateTime.Parse(vTempArray[0].Trim() + "/" + vTempArray[1].Trim() + "/" + vTempArray[2].Trim());
                                        vServiceType = vTempArray[3].Trim();
                                    }
                                    else
                                    {
                                        //如果找不到任何保養資料的時候就用領照日期當做上次保養日
                                        vGetLicDateStr = "select cast((Year(GetLicDate) - 1911) as varchar) + '/' + cast(Month(GetLicDate) as varchar) + '/' + cast(Day(GetLicDate) as varchar) ServiceDate " + Environment.NewLine +
                                                         "  from Car_InfoA " + Environment.NewLine +
                                                         " where Car_ID = '" + vCarID + "' ";
                                        vLastDate_Second = DateTime.Parse(PF.GetValue(vConnStr, vGetLicDateStr, "ServiceDate"));
                                        vServiceType = "2"; //把領照日當做最段一次三級保養日
                                    }
                                    if (vServiceType == "1")
                                    {
                                        //最後一次是二級保養時
                                        vGetTempData = "select top 1 cast((Year(FixDate) - 1911) as varchar) + ', ' + cast(Month(FixDate) as varchar) + ', ' + cast(Day(FixDate) as varchar) TempData " + Environment.NewLine +
                                                       "  from ( " + Environment.NewLine +
                                                       "        select FixDateOut as FixDate from FixWorkA where Car_ID = '" + vCarID + "' and [Service] = '2' " + Environment.NewLine +
                                                       "         union all " + Environment.NewLine +
                                                       "        select ServiceDate as FixDate from CarServicePre where Car_ID = '" + vCarID + "' and ServiceType = '2' and isnull(IsPassed, 'X') = 'X' " + Environment.NewLine +
                                                       ") t order by t.FixDate DESC";
                                        vTempStr = PF.GetValue(vConnStr, vGetTempData, "TempData");
                                        //取回上一次三級保養的日期
                                        Array.Clear(vTempArray, 0, vTempArray.Length); //先清空前一次存入的內容
                                        vTempArray = vTempStr.Split(',');
                                        //如果系統只有記錄二級保養沒有三級保養，那就用領照日當做上次三級保養日
                                        if (vTempStr.Trim() != "")
                                        {
                                            vLastDate_Third = DateTime.Parse(vTempArray[0].Trim() + "/" + vTempArray[1].Trim() + "/" + vTempArray[2].Trim());
                                        }
                                        else
                                        {
                                            vGetLicDateStr = "select cast((Year(GetLicDate) - 1911) as varchar) + '/' + cast(Month(GetLicDate) as varchar) + '/' + cast(Day(GetLicDate) as varchar) ServiceDate " + Environment.NewLine +
                                                             "  from Car_InfoA " + Environment.NewLine +
                                                             " where Car_ID = '" + vCarID + "' ";
                                            vLastDate_Third = DateTime.Parse(PF.GetValue(vConnStr, vGetLicDateStr, "ServiceDate"));
                                        }
                                    }
                                    else
                                    {
                                        //最後一次是三級保養時
                                        vLastDate_Third = vLastDate_Second;
                                    }

                                    //第一次三級保養日
                                    DateTime vThirdDate_Temp = (vLastDate_Third.AddDays(vThirdServiceDays) < vCalDate_Start) ? vCalDate_Start : vLastDate_Third.AddDays(vThirdServiceDays);

                                    //第一次二級保養日
                                    DateTime vDate_Temp = (vLastDate_Second.AddDays(vSecondServiceDays) < vCalDate_Start) ? vCalDate_Start : vLastDate_Second.AddDays(vSecondServiceDays);
                                    vDateDiff = Math.Abs(new TimeSpan(vThirdDate_Temp.Ticks - vDate_Temp.Ticks).Days); //計算跟前面的三級保養日差異天數
                                    DateTime vSecondDate_Temp = (vDateDiff <= 7) ? vThirdDate_Temp : vDate_Temp; //差距不足 7 天的話就跟三級合併

                                    //第二次二級保養日
                                    vDate_Temp = vSecondDate_Temp.AddDays(vSecondServiceDays);
                                    vDateDiff = Math.Abs(new TimeSpan(vThirdDate_Temp.Ticks - vDate_Temp.Ticks).Days); //計算跟前面的三級保養日差異天數
                                    DateTime vSecondDate_Temp2 = (vDateDiff <= 7) ? vThirdDate_Temp : vDate_Temp; //差距不足 7 天的話就跟三級合併

                                    //第三次二級保養日
                                    vDate_Temp = vSecondDate_Temp2.AddDays(vSecondServiceDays);
                                    vDateDiff = Math.Abs(new TimeSpan(vThirdDate_Temp.Ticks - vDate_Temp.Ticks).Days); //計算跟前面的三級保養日差異天數
                                    DateTime vSecondDate_Temp3 = (vDateDiff <= 7) ? vThirdDate_Temp : vDate_Temp; //差距不足 7 天的話就跟三級合併

                                    //取回前次三級之後的保養次數
                                    vSQLStr_GetRCount = "select count(t.CaseNo) RCount " + Environment.NewLine +
                                                        "  from ( " + Environment.NewLine +
                                                        "        select WorkNo as CaseNo from FixWorkA " + Environment.NewLine +
                                                        "         where Car_ID = '" + vCarID + "' " + Environment.NewLine +
                                                        "           and FixDateOut > '" + vThirdDate_Temp.Year.ToString() + "/" + vThirdDate_Temp.ToString("MM/dd") + "' " + Environment.NewLine +
                                                        "           and [Service] in ('1', '2') " + Environment.NewLine +
                                                        "         union all " + Environment.NewLine +
                                                        "        select ServicePreNo as CaseNo from CarServicePre " + Environment.NewLine +
                                                        "         where Car_ID = '" + vCarID + "' " + Environment.NewLine +
                                                        "           and ServiceDate > '" + vThirdDate_Temp.Year.ToString() + "/" + vThirdDate_Temp.ToString("MM/dd") + "' " + Environment.NewLine +
                                                        "           and isnull(IsPassed, 'X') = 'X' " + Environment.NewLine +
                                                        "       ) t";

                                    try
                                    {
                                        //先寫入三級保養
                                        if ((vThirdDate_Temp >= vCalDate_Start) && (vThirdDate_Temp <= vCalDate_End))
                                        {
                                            vRemark = "'三級，換機油'";
                                            vGetTempData = "select MAX(ServicePreNo) MaxIndex from CarServicePre where ServicePreNo like '" + vCalYM + "%' ";
                                            vTempStr = PF.GetValue(vConnStr, vGetTempData, "MaxIndex");
                                            vNewIndex = (vTempStr != "") ? Int32.Parse(vTempStr.Replace(vCalYM, "")) + 1 : 1;
                                            vExecStr = "insert into CarServicePre " + Environment.NewLine +
                                                       "       (ServicePreNo, DepNo, Car_ID, Driver, ServiceType, " + Environment.NewLine +
                                                       "        LastDate, ServiceDate, OriDate, Remark, BuMan, BuDate, IsPassed)" + Environment.NewLine +
                                                       "values ('" + vCalYM + vNewIndex.ToString("D4") + "', " + vDepNo_Car + ", '" + vCarID + "', " + vDriver + ", '2', " + Environment.NewLine +
                                                       "        '" + vLastDate_Third.Year.ToString() + "/" + vLastDate_Third.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        '" + vThirdDate_Temp.Year.ToString() + "/" + vThirdDate_Temp.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        '" + vThirdDate_Temp.Year.ToString() + "/" + vThirdDate_Temp.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        " + vRemark + ", '" + vLoginID + "', GetDate(), 'X')";
                                            PF.ExecSQL(vConnStr, vExecStr);
                                        }

                                        //第一次二級保養
                                        if ((vSecondDate_Temp >= vCalDate_Start) && (vSecondDate_Temp <= vCalDate_End) && (vSecondDate_Temp != vThirdDate_Temp))
                                        {
                                            vCaseCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_GetRCount, "RCount")) + 1;
                                            vRemark = (vCaseCount == 2) ? "'二級, 換機油'" : "NULL";

                                            vGetTempData = "select MAX(ServicePreNo) MaxIndex from CarServicePre where ServicePreNo like '" + vCalYM + "%' ";
                                            vTempStr = PF.GetValue(vConnStr, vGetTempData, "MaxIndex");
                                            vNewIndex = (vTempStr != "") ? Int32.Parse(vTempStr.Replace(vCalYM, "")) + 1 : 1;
                                            vExecStr = "insert into CarServicePre " + Environment.NewLine +
                                                       "       (ServicePreNo, DepNo, Car_ID, Driver, ServiceType, " + Environment.NewLine +
                                                       "        LastDate, ServiceDate, OriDate, Remark, BuMan, BuDate, IsPassed)" + Environment.NewLine +
                                                       "values ('" + vCalYM + vNewIndex.ToString("D4") + "', " + vDepNo_Car + ", '" + vCarID + "', " + vDriver + ", '1', " + Environment.NewLine +
                                                       "        '" + vLastDate_Second.Year.ToString() + "/" + vLastDate_Second.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        '" + vSecondDate_Temp.Year.ToString() + "/" + vSecondDate_Temp.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        '" + vSecondDate_Temp.Year.ToString() + "/" + vSecondDate_Temp.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        " + vRemark + ", '" + vLoginID + "', GetDate(), 'X')";
                                            PF.ExecSQL(vConnStr, vExecStr);
                                        }

                                        //第二次二級保養
                                        if ((vSecondDate_Temp2 >= vCalDate_Start) && (vSecondDate_Temp2 <= vCalDate_End) && (vSecondDate_Temp2 != vThirdDate_Temp))
                                        {
                                            vCaseCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_GetRCount, "RCount")) + 1;
                                            vRemark = (vCaseCount == 2) ? "'二級, 換機油'" : "NULL";

                                            vGetTempData = "select MAX(ServicePreNo) MaxIndex from CarServicePre where ServicePreNo like '" + vCalYM + "%' ";
                                            vTempStr = PF.GetValue(vConnStr, vGetTempData, "MaxIndex");
                                            vNewIndex = (vTempStr != "") ? Int32.Parse(vTempStr.Replace(vCalYM, "")) + 1 : 1;
                                            vExecStr = "insert into CarServicePre " + Environment.NewLine +
                                                       "       (ServicePreNo, DepNo, Car_ID, Driver, ServiceType, " + Environment.NewLine +
                                                       "        LastDate, ServiceDate, OriDate, Remark, BuMan, BuDate, IsPassed)" + Environment.NewLine +
                                                       "values ('" + vCalYM + vNewIndex.ToString("D4") + "', " + vDepNo_Car + ", '" + vCarID + "', " + vDriver + ", '1', " + Environment.NewLine +
                                                       "        '" + vLastDate_Second.Year.ToString() + "/" + vLastDate_Second.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        '" + vSecondDate_Temp2.Year.ToString() + "/" + vSecondDate_Temp2.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        '" + vSecondDate_Temp2.Year.ToString() + "/" + vSecondDate_Temp2.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        " + vRemark + ", '" + vLoginID + "', GetDate(), 'X')";
                                            PF.ExecSQL(vConnStr, vExecStr);
                                        }

                                        //第三次二級保養
                                        if ((vSecondDate_Temp3 >= vCalDate_Start) && (vSecondDate_Temp3 <= vCalDate_End) && (vSecondDate_Temp3 != vThirdDate_Temp))
                                        {
                                            vCaseCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_GetRCount, "RCount")) + 1;
                                            vRemark = (vCaseCount == 2) ? "'二級, 換機油'" : "NULL";

                                            vGetTempData = "select MAX(ServicePreNo) MaxIndex from CarServicePre where ServicePreNo like '" + vCalYM + "%' ";
                                            vTempStr = PF.GetValue(vConnStr, vGetTempData, "MaxIndex");
                                            vNewIndex = (vTempStr != "") ? Int32.Parse(vTempStr.Replace(vCalYM, "")) + 1 : 1;
                                            vExecStr = "insert into CarServicePre " + Environment.NewLine +
                                                       "       (ServicePreNo, DepNo, Car_ID, Driver, ServiceType, " + Environment.NewLine +
                                                       "        LastDate, ServiceDate, OriDate, Remark, BuMan, BuDate, IsPassed)" + Environment.NewLine +
                                                       "values ('" + vCalYM + vNewIndex.ToString("D4") + "', " + vDepNo_Car + ", '" + vCarID + "', " + vDriver + ", '1', " + Environment.NewLine +
                                                       "        '" + vLastDate_Second.Year.ToString() + "/" + vLastDate_Second.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        '" + vSecondDate_Temp3.Year.ToString() + "/" + vSecondDate_Temp3.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        '" + vSecondDate_Temp3.Year.ToString() + "/" + vSecondDate_Temp3.ToString("MM/dd") + "', " + Environment.NewLine +
                                                       "        " + vRemark + ", '" + vLoginID + "', GetDate(), 'X')";
                                            PF.ExecSQL(vConnStr, vExecStr);
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
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先輸入要計算的保養年月！')");
                Response.Write("</" + "Script>");
                eCalYear_Search.Focus();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvServicePreDetail_DataBound(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDateURL = "";
            string vDateScript = "";
            string vCheckDateStr = calCarServiceList.SelectedDate.Year.ToString("D4") + calCarServiceList.SelectedDate.Month.ToString("D2") + "01";
            string vControlItem = "a_ServicePreClose" + eDepNo_Search.Text.Trim();
            Boolean vIsClose = PF.DateIsClosed(vConnStr, vControlItem, vCheckDateStr);
            switch (fvServicePreDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Button bbInsert = (Button)fvServicePreDetail.FindControl("bbInsert");
                    if (bbInsert != null)
                    {
                        bbInsert.Visible = (vIsClose == false);
                    }
                    Button bbEdit = (Button)fvServicePreDetail.FindControl("bbEdit");
                    if (bbEdit != null)
                    {
                        bbEdit.Visible = (vIsClose == false);
                    }
                    Button bbDelete = (Button)fvServicePreDetail.FindControl("bbDelete");
                    if (bbDelete != null)
                    {
                        bbDelete.Visible = (vIsClose == false);
                    }
                    break;

                case FormViewMode.Edit:
                    if (!vIsClose)
                    {
                        //異動人
                        Label eModifyMan_Edit = (Label)fvServicePreDetail.FindControl("eModifyMan_Edit");
                        eModifyMan_Edit.Text = vLoginID;
                        Label eModifyManName_Edit = (Label)fvServicePreDetail.FindControl("eModifyManName_Edit");
                        eModifyManName_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                        //異動日期
                        Label eModifyDate_Edit = (Label)fvServicePreDetail.FindControl("eModifyDate_Edit");
                        eModifyDate_Edit.Text = PF.GetChinsesDate(DateTime.Today);
                        //預排保養日
                        TextBox eServiceDate_Edit = (TextBox)fvServicePreDetail.FindControl("eServiceDate_Edit");
                        eServiceDate_Edit.Text = PF.GetChinsesDate(calCarServiceList.SelectedDate);
                        vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eServiceDate_Edit.ClientID;
                        vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eServiceDate_Edit.Attributes["onClick"] = vDateScript;
                        //保養類別
                        Label eServiceType_Edit = (Label)fvServicePreDetail.FindControl("eServiceType_Edit");
                        DropDownList ddlServiceType_Edit = (DropDownList)fvServicePreDetail.FindControl("ddlServiceType_Edit");
                        if (ddlServiceType_Edit != null)
                        {
                            int vIndex = ddlServiceType_Edit.Items.IndexOf(ddlServiceType_Edit.Items.FindByValue(eServiceType_Edit.Text.Trim()));
                            ddlServiceType_Edit.SelectedIndex = (eServiceType_Edit.Text.Trim() == "") ? 0 : ddlServiceType_Edit.Items.IndexOf(ddlServiceType_Edit.Items.FindByValue(eServiceType_Edit.Text.Trim()));
                        }
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('指定要計算的保養日期已經關帳！')");
                        Response.Write("</" + "Script>");
                        fvServicePreDetail.ChangeMode(FormViewMode.ReadOnly);
                    }
                    break;

                case FormViewMode.Insert:
                    if (!vIsClose)
                    {
                        //建檔人
                        Label eBuMan_INS = (Label)fvServicePreDetail.FindControl("eBuMan_INS");
                        eBuMan_INS.Text = vLoginID;
                        Label eBuManName_INS = (Label)fvServicePreDetail.FindControl("eBuManNAme_INS");
                        eBuManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                        //建檔日期
                        Label eBuDate_INS = (Label)fvServicePreDetail.FindControl("eBuDate_INS");
                        eBuDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                        //預排保養日
                        TextBox eServiceDate_INS = (TextBox)fvServicePreDetail.FindControl("eServiceDate_INS");
                        if (calCarServiceList.SelectedDate.Year >= 1912)
                        {
                            eServiceDate_INS.Text = PF.GetChinsesDate(calCarServiceList.SelectedDate);
                        }
                        vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eServiceDate_INS.ClientID;
                        vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eServiceDate_INS.Attributes["onClick"] = vDateScript;
                        //保養類別
                        DropDownList ddlServiceType_INS = (DropDownList)fvServicePreDetail.FindControl("ddlServiceType_INS");
                        if (ddlServiceType_INS != null)
                        {
                            ddlServiceType_INS.SelectedIndex = 0;
                            Label eServiceType_INS = (Label)fvServicePreDetail.FindControl("eServiceType_INS");
                            eServiceType_INS.Text = ddlServiceType_INS.SelectedValue;
                        }
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('指定要計算的保養日期已經關帳！')");
                        Response.Write("</" + "Script>");
                        fvServicePreDetail.ChangeMode(FormViewMode.ReadOnly);
                    }
                    break;
            }
        }

        /// <summary>
        /// 確定修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eServicePreNo_Edit = (Label)fvServicePreDetail.FindControl("eServicePreNo_Edit");
            if (eServicePreNo_Edit != null)
            {
                string vServicePreNo_Edit = eServicePreNo_Edit.Text.Trim();
                Label eServiceType_Edit = (Label)fvServicePreDetail.FindControl("eServiceType_Edit");
                string vServiceType_Edit = eServiceType_Edit.Text.Trim();
                TextBox eServiceDate_Edit = (TextBox)fvServicePreDetail.FindControl("eServiceDate_Edit");
                DateTime vServiceDate_Edit = DateTime.Parse(eServiceDate_Edit.Text.Trim());
                Label eOldDate_Edit = (Label)fvServicePreDetail.FindControl("eOldDate_Edit");
                DateTime vOldDate_Edit = DateTime.Parse(eOldDate_Edit.Text.Trim());
                //日期差距
                int vDateDiff = Math.Abs(new TimeSpan(vOldDate_Edit.Ticks - vServiceDate_Edit.Ticks).Days);
                if (vDateDiff <= 7)
                {
                    string vServiceDate_Str = vServiceDate_Edit.ToString("yyyy/MM/dd");
                    TextBox eDriver_Edit = (TextBox)fvServicePreDetail.FindControl("eDriver_Edit");
                    string vDriver_Edit = eDriver_Edit.Text.Trim();
                    TextBox eRemark_Edit = (TextBox)fvServicePreDetail.FindControl("eRemark_Edit");
                    string vRemark_Edit = eRemark_Edit.Text.Trim();
                    Label eModifyMan_Edit = (Label)fvServicePreDetail.FindControl("eModifyMan_Edit");
                    string vModifyMan_Edit = eModifyMan_Edit.Text.Trim();
                    Label eMosifyDate_Edit = (Label)fvServicePreDetail.FindControl("eModifyDate_Edit");
                    DateTime vModifyDate_Edit = DateTime.Parse(eMosifyDate_Edit.Text.Trim());
                    string vModifyDate_Str = vModifyDate_Edit.ToString("yyyy/MM/dd");

                    string vSQLStr_Temp = "select MAX(Items) MaxIndex from  CarServicePre_History where ServicePreNo = '" + vServicePreNo_Edit + "' ";
                    string vMaxIndexStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    int vNewIndex = (vMaxIndexStr.Trim() == "") ? 1 : Int32.Parse(vMaxIndexStr) + 1;
                    string vNewServicePreNoItem = vServicePreNo_Edit + vNewIndex.ToString("D4");
                    vSQLStr_Temp = "insert into CarServicePre_History " + Environment.NewLine +
                                   "       (ServicePreNoItem, Items, ModifyMode, ServicePreNo, DepNo, Car_ID, Driver, ServiceType, LastDate, ServiceDate," + Environment.NewLine +
                                   "        Remark, BuMan, BuDate, ModifyMan, ModifyDate, OriDate) " + Environment.NewLine +
                                   "select '" + vNewServicePreNoItem.Trim() + "', '" + vNewIndex.ToString("D4") + "', 'EDIT', " + Environment.NewLine +
                                   "       ServicePreNo, DepNo, Car_ID, Driver, ServiceType, LastDate, ServiceDate," + Environment.NewLine +
                                   "       Remark, BuMan, BuDate, ModifyMan, ModifyDate, OriDate " + Environment.NewLine +
                                   "  from CarServicePre " + Environment.NewLine +
                                   " where ServicePreNo = '" + vServicePreNo_Edit.Trim() + "' ";
                    try
                    {
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        sdsShowDetail.UpdateParameters.Clear();
                        sdsShowDetail.UpdateParameters.Add(new Parameter("ServiceType", DbType.String, vServiceType_Edit.Trim()));
                        sdsShowDetail.UpdateParameters.Add(new Parameter("ServiceDate", DbType.Date, vServiceDate_Str.Trim()));
                        sdsShowDetail.UpdateParameters.Add(new Parameter("Driver", DbType.String, vDriver_Edit.Trim()));
                        sdsShowDetail.UpdateParameters.Add(new Parameter("ServicePreNo", DbType.String, vServicePreNo_Edit.Trim()));
                        sdsShowDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, vRemark_Edit.Trim()));
                        sdsShowDetail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vModifyMan_Edit.Trim()));
                        sdsShowDetail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.Date, vModifyDate_Str.Trim()));
                        sdsShowDetail.Update();
                        fvServicePreDetail.ChangeMode(FormViewMode.ReadOnly);
                        DateTime vSelectDate = calCarServiceList.SelectedDate;
                        string vSelDateStr = vSelectDate.Year.ToString() + "/" + vSelectDate.ToString("MM/dd");
                        vSQLStr_Temp = GetSelectStr(vSelDateStr);
                        sdsShowList.SelectCommand = vSQLStr_Temp;
                        sdsShowList.DataBind();
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
                    Response.Write("alert('日期調整不可超過七天！')");
                    Response.Write("</" + "Script>");
                    eServiceDate_Edit.Focus();
                }
            }
        }

        protected void eDriver_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Edit = (TextBox)fvServicePreDetail.FindControl("eDriver_Edit");
            Label eDriverName_Edit = (Label)fvServicePreDetail.FindControl("eDriverName_Edit");
            string vDriver_Temp = eDriver_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp.Trim() == "")
            {
                vDriverName_Temp = vDriver_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp + "' order by isnull(LeaveDay, '9999/12/31') DESC ";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Edit.Text = vDriver_Temp.Trim();
            eDriverName_Edit.Text = vDriverName_Temp.Trim();
        }

        protected void ddlServiceType_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlServiceType_Edit = (DropDownList)fvServicePreDetail.FindControl("ddlServiceType_Edit");
            Label eServiceType_Edit = (Label)fvServicePreDetail.FindControl("eServiceType_Edit");
            eServiceType_Edit.Text = ddlServiceType_Edit.SelectedValue.Trim();
        }

        /// <summary>
        /// 輸入車號取回車輛資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eCar_ID_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eCar_ID_INS = (TextBox)fvServicePreDetail.FindControl("eCar_ID_INS");
            TextBox eDepNo_INS = (TextBox)fvServicePreDetail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvServicePreDetail.FindControl("eDepName_INS");
            TextBox eDriver_INS = (TextBox)fvServicePreDetail.FindControl("eDriver_INS");
            Label eDriverName_INS = (Label)fvServicePreDetail.FindControl("eDriverName_INS");
            Label eLastDate_INS = (Label)fvServicePreDetail.FindControl("eLastDate_INS");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "select (isnull(CompanyNo, '') + ', ' + isnull((select [Name] from Department where DepNo = a.CompanyNo), '') + ', ' + isnull(Driver, '') + ', ' + isnull((select [Name] from Employee where EmpNo = a.Driver), '')) CarData from Car_InfoA a where Car_ID = '" + eCar_ID_INS.Text.Trim() + "' ";
            string vTempStr = PF.GetValue(vConnStr, vSQLStr_Temp, "CarData");
            if (vTempStr.Trim() != "")
            {
                string[] vTempArray = vTempStr.Split(',');
                eDepNo_INS.Text = vTempArray[0].Trim();
                eDepName_INS.Text = vTempArray[1].Trim();
                eDriver_INS.Text = vTempArray[2].Trim();
                eDriverName_INS.Text = vTempArray[3].Trim();
                vSQLStr_Temp = "select top 1 (cast((Year(ServiceDate) - 1911) as varchar) + '/' + cast(Month(ServiceDate) as varchar) + '/' + cast(Day(ServiceDate) as varchar)) TempData " + Environment.NewLine +
                               "  from ( " + Environment.NewLine +
                               "        select FixDateOut as ServiceDate, [Service] as ServiceType " + Environment.NewLine +
                               "          from FixWorkA " + Environment.NewLine +
                               "         where Car_ID = '" + eCar_ID_INS.Text.Trim() + "' " + Environment.NewLine +
                               "           and [Service] in ('1', '2') " + Environment.NewLine +
                               "         union all " + Environment.NewLine +
                               "        select ServiceDate, ServiceType " + Environment.NewLine +
                               "          from CarServicePre " + Environment.NewLine +
                               "         where Car_ID = '" + eCar_ID_INS.Text.Trim() + "' " + Environment.NewLine +
                               "           and ServiceDate >= '" + DateTime.Today.Year.ToString() + "/" + DateTime.Today.ToString("MM/dd") + "' " + Environment.NewLine +
                               "       ) t order by ServiceDate DESC ";
                eLastDate_INS.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "TempData");
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_INS = (TextBox)fvServicePreDetail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvServicePreDetail.FindControl("eDepName_INS");
            string vDepNo_INS = eDepNo_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_INS.Trim() + "' ";
            string vDepName_INS = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_INS.Trim() == "")
            {
                vDepName_INS = vDepNo_INS.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_INS.Trim() + "' ";
                vDepNo_INS = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_INS.Text = vDepNo_INS.Trim();
            eDepName_INS.Text = vDepName_INS.Trim();
        }

        protected void eDriver_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_INS = (TextBox)fvServicePreDetail.FindControl("eDriver_INS");
            Label eDriverName_INS = (Label)fvServicePreDetail.FindControl("eDriverName_INS");
            string vDriver_INS = eDriver_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_INS + "' ";
            string vDriverName_INS = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_INS == "")
            {
                vDriverName_INS = vDriver_INS.Trim();
                vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vDriverName_INS.Trim() + "' order by isnull(Leaveday, '9999/12/31') DESC";
                vDriver_INS = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_INS.Text = vDriver_INS.Trim();
            eDriverName_INS.Text = vDriverName_INS.Trim();
        }

        protected void ddlServiceType_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlServiceType_INS = (DropDownList)fvServicePreDetail.FindControl("ddlServiceType_INS");
            Label eServiceType_INS = (Label)fvServicePreDetail.FindControl("eServiceType_INS");
            eServiceType_INS.Text = ddlServiceType_INS.SelectedValue.Trim();
        }

        /// <summary>
        /// 刪除資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eServicePreNo_List = (Label)fvServicePreDetail.FindControl("eServicePreNo_List");
            string vServicePreNo_Temp = eServicePreNo_List.Text.Trim();
            string vSQLStr_Temp = "select MAX(Items) MaxIndex from  CarServicePre_History where ServicePreNo = '" + vServicePreNo_Temp + "' ";
            string vMaxIndexStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
            int vNewIndex = (vMaxIndexStr.Trim() == "") ? 1 : Int32.Parse(vMaxIndexStr) + 1;
            string vNewServicePreNoItem = vServicePreNo_Temp + vNewIndex.ToString("D4");
            vSQLStr_Temp = "insert into CarServicePre_History " + Environment.NewLine +
                           "       (ServicePreNoItem, Items, ModifyMode, ServicePreNo, DepNo, Car_ID, Driver, ServiceType, LastDate, ServiceDate," + Environment.NewLine +
                           "        Remark, BuMan, BuDate, ModifyMan, ModifyDate, OriDate) " + Environment.NewLine +
                           "select '" + vNewServicePreNoItem.Trim() + "', '" + vNewIndex.ToString("D4") + "', 'DEL', " + Environment.NewLine +
                           "       ServicePreNo, DepNo, Car_ID, Driver, ServiceType, LastDate, ServiceDate," + Environment.NewLine +
                           "       Remark, BuMan, BuDate, ModifyMan, ModifyDate, OriDate " + Environment.NewLine +
                           "  from CarServicePre " + Environment.NewLine +
                           " where ServicePreNo = '" + vServicePreNo_Temp.Trim() + "' ";
            try
            {
                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                sdsShowDetail.DeleteParameters.Clear();
                sdsShowDetail.DeleteParameters.Add(new Parameter("ServicePreNo", DbType.String, vServicePreNo_Temp.Trim()));
                sdsShowDetail.Delete();
                DateTime vSelectDate = calCarServiceList.SelectedDate;
                string vSelDateStr = vSelectDate.Year.ToString() + "/" + vSelectDate.ToString("MM/dd");
                vSQLStr_Temp = GetSelectStr(vSelDateStr);
                sdsShowList.SelectCommand = vSQLStr_Temp;
                sdsShowList.DataBind();
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        protected void calCarServiceList_DayRender(object sender, DayRenderEventArgs e)
        {
            if (e.Day.IsOtherMonth)
            {
                e.Cell.ForeColor = System.Drawing.Color.White;
            }
            string vServiceDateStr = e.Day.Date.Year.ToString("D4") + "/" + e.Day.Date.ToString("MM/dd");
            string vSelStr = GetCarListStr(vServiceDateStr);

            //取回要顯示的資料
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    ListBox lbTemp = new ListBox();
                    int vIndex = 0;
                    while (drTemp.Read())
                    {
                        lbTemp.Items.Insert(vIndex, drTemp["Car_ID"].ToString() + " : " + drTemp["ServiceType_C"].ToString());
                        vIndex++;
                    }
                    lbTemp.Rows = vIndex;
                    lbTemp.Attributes.Remove("Style");
                    lbTemp.Attributes.Add("Style", "Width: 95%; Height:120px");

                    e.Cell.Controls.Add(new LiteralControl("<br />"));
                    e.Cell.Controls.Add(lbTemp);
                }
            }
        }

        protected void calCarServiceList_SelectionChanged(object sender, EventArgs e)
        {
            //行事曆日期變更
            string vSQLStr_Temp = "";
            DateTime vSelectDate = calCarServiceList.SelectedDate;
            string vSelDateStr = vSelectDate.Year.ToString() + "/" + vSelectDate.ToString("MM/dd");

            vSQLStr_Temp = GetSelectStr(vSelDateStr);
            sdsShowList.SelectCommand = vSQLStr_Temp;
            gridDailyList.DataBind();
            plShowData.Visible = true;
        }

        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vErrorMSG = "";

            TextBox eServiceDate_INS = (TextBox)fvServicePreDetail.FindControl("eServiceDate_INS");
            if (eServiceDate_INS != null)
            {
                DateTime vServiceDate = DateTime.Parse(eServiceDate_INS.Text.Trim());
                DateTime vOriDate = vServiceDate;
                DateTime vBuDate = DateTime.Today;
                string vCalYM_INS = eCalYear_Search.Text.Trim() + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2");
                string vSQLStr_Temp = "select MAX(ServicePreNo) ReturnStr from CarServicePre where ServicePreNo like '" + vCalYM_INS.Trim() + "%' ";
                string vTempStr = PF.GetValue(vConnStr, vSQLStr_Temp, "ReturnStr");
                int vNewIndex = (vTempStr == "") ? 1 : Int32.Parse(vTempStr.Replace(vCalYM_INS, "")) + 1;
                string vServicePreNo_New = vCalYM_INS + vNewIndex.ToString("D4");
                TextBox eDepNo_INS = (TextBox)fvServicePreDetail.FindControl("eDepNo_INS");
                string vDepNo_INS = (eDepNo_INS.Text.Trim() != "") ? eDepNo_INS.Text.Trim() : vLoginDepNo;
                TextBox eCar_ID_INS = (TextBox)fvServicePreDetail.FindControl("eCar_ID_INS");
                string vCar_ID_INS = (eCar_ID_INS.Text.Trim() != "") ? eCar_ID_INS.Text.Trim() : "";
                if (vCar_ID_INS == "")
                {
                    vErrorMSG = (vErrorMSG.Trim() == "") ? "車號不可空白！" : vErrorMSG + Environment.NewLine + "車號不可空白！";
                }
                TextBox eDriver_INS = (TextBox)fvServicePreDetail.FindControl("eDriver_INS");
                string vDriver_INS = (eDriver_INS.Text.Trim() != "") ? eDriver_INS.Text.Trim() : "";
                if (vDriver_INS.Trim() == "")
                {
                    vErrorMSG = (vErrorMSG == "") ? "駕駛員不可空白！" : vErrorMSG + Environment.NewLine + "駕駛員不可空白！";
                }
                Label eServiceType_INS = (Label)fvServicePreDetail.FindControl("eServiceType_INS");
                string vServiceType = eServiceType_INS.Text.Trim();
                Label eLastDate_INS = (Label)fvServicePreDetail.FindControl("eLastDate_INS");
                DateTime vLastDate = (eLastDate_INS.Text.Trim() != "") ? DateTime.Parse(eLastDate_INS.Text.Trim()) : DateTime.Today;
                TextBox eRemark_INS = (TextBox)fvServicePreDetail.FindControl("eRemark_INS");
                string vRemark_INS = eRemark_INS.Text.Trim();

                if (vErrorMSG.Trim() == "")
                {
                    string vSQL_INSStr = "INSERT INTO CarServicePre (ServicePreNo, DepNo, Car_ID, Driver, ServiceType, LastDate, ServiceDate, Remark, BuMan, BuDate, OriDate) " + Environment.NewLine +
                                         "                   VALUES (@ServicePreNo, @DepNo, @Car_ID, @Driver, @ServiceType, @LastDate, @ServiceDate, @Remark, @BuMan, @BuDate, @OriDate)";
                    sdsShowDetail.InsertParameters.Clear();
                    sdsShowDetail.InsertCommand = vSQL_INSStr;
                    sdsShowDetail.InsertParameters.Add(new Parameter("ServicePreNo", DbType.String, vServicePreNo_New));
                    sdsShowDetail.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_INS));
                    sdsShowDetail.InsertParameters.Add(new Parameter("Car_ID", DbType.String, (vCar_ID_INS.Trim() != "") ? vCar_ID_INS : String.Empty));
                    sdsShowDetail.InsertParameters.Add(new Parameter("Driver", DbType.String, (vDriver_INS.Trim() != "") ? vDriver_INS : String.Empty));
                    sdsShowDetail.InsertParameters.Add(new Parameter("ServiceType", DbType.String, (vServiceType.Trim() != "") ? vServiceType.Trim() : String.Empty));
                    sdsShowDetail.InsertParameters.Add(new Parameter("LastDate", DbType.Date, (eLastDate_INS.Text.Trim() != "") ? vLastDate.ToString("yyyy/MM/dd") : String.Empty));
                    sdsShowDetail.InsertParameters.Add(new Parameter("ServiceDate", DbType.Date, (eServiceDate_INS.Text.Trim() != "") ? vServiceDate.ToString("yyyy/MM/dd") : String.Empty));
                    sdsShowDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark_INS.Trim() != "") ? vRemark_INS.Trim() : String.Empty));
                    sdsShowDetail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsShowDetail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.Today.ToString("yyyy/MM/dd")));
                    sdsShowDetail.InsertParameters.Add(new Parameter("OriDate", DbType.Date, (eServiceDate_INS.Text.Trim() != "") ? vServiceDate.ToString("yyyy/MM/dd") : String.Empty));
                    try
                    {
                        sdsShowDetail.Insert();
                        fvServicePreDetail.ChangeMode(FormViewMode.ReadOnly);
                        DateTime vSelectDate = calCarServiceList.SelectedDate;
                        string vSelDateStr = vSelectDate.Year.ToString() + "/" + vSelectDate.ToString("MM/dd");
                        vSQLStr_Temp = GetSelectStr(vSelDateStr);
                        sdsShowList.SelectCommand = vSQLStr_Temp;
                        sdsShowList.DataBind();
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
                    Response.Write("alert('" + vErrorMSG + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// 設定預排關帳
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbStopCalYM_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            try
            {
                string vControlItem = "a_ServicePreClose" + eDepNo_Search.Text.Trim();
                string vSQLStr_Temp = "select cast(count(Content) as nvarchar) + ',' + Content as RCount " + Environment.NewLine +
                                      "  from SysFlag " + Environment.NewLine +
                                      " where FormName = 'unSysflag' and ControlItem = '" + vControlItem + "' " + Environment.NewLine +
                                      " group by Content";
                string vItemData_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount").Trim();
                string[] vItemData = vItemData_Str.Split(',');
                int vHasCloseDate = (vItemData_Str != "") ? Int32.Parse(vItemData[0].ToString().Trim()) : 0;
                string vCloseDate_New = (Int32.Parse(eCalYear_Search.Text.Trim()) + 1911).ToString("D4") + Int32.Parse(eCalMonth_Search.Text.Trim()).ToString("D2") + "01";
                vSQLStr_Temp = (vHasCloseDate == 0) ?
                               "Insert into SysFlag(FormName, ControlItem, Content) values('unSysFlag', '" + vControlItem + "', '" + vCloseDate_New.Trim() + "')" :
                               (Int32.Parse(vItemData[1].ToString().Trim()) < Int32.Parse(vCloseDate_New)) ?
                               "update SysFlag set Content = '" + vCloseDate_New.Trim() + "' where FormName = 'unSysflag' and ControlItem = '" + vControlItem + "'" : "";
                if (vSQLStr_Temp != "")
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('指定站別 [" + ddlDepNo_Search.Text.Trim() + "] 關帳設定完成！')");
                    Response.Write("</" + "Script>");
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('指定站別 [" + ddlDepNo_Search.Text.Trim() + "] 保養預排單關帳日期已經設定於 [" + vItemData[1].ToString().Trim() + "] ')");
                    Response.Write("</" + "Script>");
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
        /// 開放修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbReOpenYM_Click(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDepNo_Search.Text.Trim();
            string vControlItem = "a_ServicePreClose" + vDepNo_Temp;
            if (vDepNo_Temp.Trim() != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSQLStr_Temp = "select count(Content) as RCount " + Environment.NewLine +
                                      "  from SysFlag " + Environment.NewLine +
                                      " where FormName = 'unSysflag' and ControlItem = '" + vControlItem + "' ";
                string vItemCount_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount").Trim();
                if ((vItemCount_Str != "") && (Int32.Parse(vItemCount_Str) > 0))
                {
                    if ((eCalYear_Search.Text.Trim() != "") && (eCalMonth_Search.Text.Trim() != ""))
                    {
                        DateTime vCalDate = DateTime.Parse(eCalYear_Search.Text.Trim() + "/" + eCalMonth_Search.Text.Trim() + "/01").AddMonths(-1);
                        string vStopDateStr = vCalDate.Year.ToString("D4") + vCalDate.Month.ToString("D2") + vCalDate.Day.ToString("D2");
                        vSQLStr_Temp = "update SysFlag set Content = '" + vStopDateStr.Trim() + "' where FormName = 'unSysflag' and ControlItem = '" + vControlItem + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('指定站別 [" + ddlDepNo_Search.Text.Trim() + "] 關帳設定完成！')");
                        Response.Write("</" + "Script>");
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('計算年月不可空白！')");
                        Response.Write("</" + "Script>");

                        if (eCalYear_Search.Text.Trim() == "")
                        {
                            eCalYear_Search.Focus();
                        }
                        else
                        {
                            eCalMonth_Search.Focus();
                        }
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('指定站別 [" + ddlDepNo_Search.Text.Trim() + "] 未進行關帳作業，不需解鎖！')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請選擇站別！')");
                Response.Write("</" + "Script>");
                ddlDepNo_Search.Focus();
            }
        }
    }
}