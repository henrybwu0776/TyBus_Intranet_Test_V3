using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverStatus : System.Web.UI.Page
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

        private string vEmpDataURL = "";
        private string vEmpDataScript = "";
        private DataTable dtDriverStatus;

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
                    /* 在這裡設定員工工號輸入欄位的新視窗相關設定*/
                    //vEmpDataURL = "SearchEmp.aspx?TextBoxID=" + eDriver_Search.ClientID + "&LabelID=" + eDriverName_Search.ClientID;
                    vEmpDataURL = "SearchEmp.aspx?TextBoxID=" + eDriver_Search.ClientID;
                    vEmpDataScript = "window.open('" + vEmpDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                    eDriver_Search.Attributes["onClick"] = vEmpDataScript;
                    //======================================================================================================
                    if (!IsPostBack)
                    {
                        eCalYear_Search.Text = (DateTime.Today.AddMonths(-1).Year - 1911).ToString();
                        ddlCalMonth_Search.SelectedIndex = DateTime.Today.AddMonths(-1).Month - 1;
                    }
                    else
                    {
                        /*
                        if (eDriver_Search.Text.Trim() != "")
                        {
                            OpenData();
                        }
                        //*/
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

        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            /* 2023.02.23 修改，同一姓名比對到多筆資料時的處理方式
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDriverNo = eDriver_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriverNo + "' ";
            string vDriverName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriverNo.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName + "' order by EmpNo DESC ";
                vDriverNo = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Search.Text = vDriverNo.Trim();
            eDriverName_Search.Text = vDriverName.Trim();
            */
        }

        private string GetSelStr()
        {
            string vReturnStr = "";
            int vCalYear = 0;
            if (!Int32.TryParse(eCalYear_Search.Text.Trim(), out vCalYear))
            {
                vCalYear = DateTime.Today.Year;
            }
            string vCalYear_Str = (vCalYear < 1911) ? (vCalYear + 1911).ToString() : vCalYear.ToString();
            string vCalDate = vCalYear_Str + "/" + ddlCalMonth_Search.SelectedValue.Trim() + "/01";
            vReturnStr = "select TA.DT, TA.[Name], TA.Birthday, TA.IDCardNo, TA.WorkType, TA.YearCount, TA.Assumeday, TA.DepNo, TA.EmpNo, TA.TelNo1, TA.CellPhone, TA.Addr1, TA.Title_C as Title, " + Environment.NewLine +
                         "       isnull(RA.WorkHR, 0) WorkHR, (select ClassTxt from DBDICB where FKey = '行車記錄單A     runsheeta       WORKSTATE       ' and ClassNo = RA.WorkState) WorkState, " + Environment.NewLine +
                         "       TA.ESCType_04, TA.ESCType_05, TA.MonthPay " + Environment.NewLine +
                         "  from( " + Environment.NewLine +
                         "       select aa.DT, e.Name, e.Birthday, e.IDCardNo, e.WorkType, datediff(Month, e.Expday, getdate()) / 12.0 as YearCount, e.Assumeday, e.DepNo, " + Environment.NewLine +
                         "              e.EmpNo, e.TelNo1, e.CellPhone, e.Addr1, (select ClassTxt from DBDICB where FKey = '人事資料檔      EMPLOYEE        TITLE' and ClassNo = e.Title) Title_C, " + Environment.NewLine +
                         "              cast(null as varchar) ESCType_04, cast(null as varchar) as ESCType_05, " + Environment.NewLine +
                         "              isnull((select GivCash from PayRec where PayDate = '" + vCalDate + "' and PayDur = '1' and EmpNo = '" + eDriver_Search.Text.Trim() + "'), 0) MonthPay " + Environment.NewLine +
                         "         from( " + Environment.NewLine +
                         "              SELECT number + 1  N, DATEADD(DD, number, CONVERT(CHAR(8), '" + vCalDate + "', 120) + '01') DT " + Environment.NewLine +
                         "                FROM master.dbo.spt_values " + Environment.NewLine +
                         "               WHERE name IS NULL " + Environment.NewLine +
                         "                 AND number < DAY(DATEADD(MM, 1, CONVERT(CHAR(8), '" + vCalDate + "', 120) + '01') - 1) " + Environment.NewLine +
                         "             ) as aa, Employee as e --left join Alters as ea on ea.EmpNo = e.EmpNo " + Environment.NewLine +
                         "        where e.empno = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine +
                         "      ) as TA left join RunSheetA as RA on RA.BuDate = TA.DT and RA.Driver = TA.EmpNo";
            return vReturnStr;
        }

        private void CalData()
        {
            string vSQLStr_Temp = "";
            string vSelStr_Temp = GetSelStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vCalDate;
            string vCalDateStr = "";
            string vDriverNo = eDriver_Search.Text.Trim();
            string vESCType = "";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelStr_Temp, connTemp);
                connTemp.Open();
                dtDriverStatus = new DataTable();
                daTemp.Fill(dtDriverStatus);
                if (dtDriverStatus.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDriverStatus.Rows.Count; i++)
                    {
                        vCalDate = DateTime.Parse(dtDriverStatus.Rows[i]["DT"].ToString());
                        vCalDateStr = PF.TransDateString(vCalDate, "B");
                        //先找有沒有事假
                        vSQLStr_Temp = "select ESCType from ESCDuty where ESCType = '04' and ApplyMan = '" + vDriverNo + "' and RealDay = '" + vCalDateStr + "' ";
                        vESCType = PF.GetValue(vConnStr, vSQLStr_Temp, "ESCType");
                        dtDriverStatus.Rows[i]["ESCType_04"] = (vESCType != "") ? "事假" : dtDriverStatus.Rows[i]["ESCType_04"].ToString();
                        //再看有沒有病假
                        vSQLStr_Temp = "select ESCType from ESCDuty where ESCType = '05' and ApplyMan = '" + vDriverNo + "' and RealDay = '" + vCalDateStr + "' ";
                        vESCType = PF.GetValue(vConnStr, vSQLStr_Temp, "ESCType");
                        dtDriverStatus.Rows[i]["ESCType_05"] = (vESCType != "") ? "病假" : dtDriverStatus.Rows[i]["ESCType_05"].ToString();
                    }
                    dtDriverStatus.AcceptChanges();
                }
            }
        }

        private void OpenData()
        {
            CalData();
            gridDriverStatus.DataSourceID = "";
            gridDriverStatus.DataSource = dtDriverStatus;
            gridDriverStatus.DataBind();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            if (eDriver_Search.Text.Trim() != "")
            {
                OpenData();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請輸入駕駛員工號！')");
                Response.Write("</" + "Script>");
                eDriver_Search.Focus();
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (eDriver_Search.Text.Trim() != "")
            {
                ExportExcel();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請輸入駕駛員工號！')");
                Response.Write("</" + "Script>");
                eDriver_Search.Focus();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        private void ExportExcel()
        {
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

            CalData(); //產生資料 TABLE
            if (dtDriverStatus.Rows.Count > 0)
            {
                int vLinesNo = 0;
                string vColName = "";
                string vColHeadText = "";
                string vTempStr = "";
                DateTime vCalDate;

                string vFileName = eCalYear_Search.Text.Trim() + "年" + ddlCalMonth_Search.SelectedItem.Text.Trim() + "駕駛員：" + eDriverName_Search.Text.Trim() + "出勤及請休假狀況表";
                //開始準備轉入 EXCEL 的資料
                wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                //寫入標題列
                wsExcel.CreateRow(vLinesNo);
                for (int ColIndex = 0; ColIndex < dtDriverStatus.Columns.Count; ColIndex++)
                {
                    vColName = dtDriverStatus.Columns[ColIndex].ColumnName.ToUpper().Trim();
                    vColHeadText = (vColName == "DT") ? "" :
                                   (vColName == "NAME") ? "姓名" :
                                   (vColName == "BIRTHDAY") ? "出生日期" :
                                   (vColName == "IDCARDNO") ? "身分證字號" :
                                   (vColName == "WORKTYPE") ? "在職情形" :
                                   (vColName == "YEARCOUNT") ? "年資" :
                                   (vColName == "ASSUMEDAY") ? "受雇日期" :
                                   (vColName == "DEPNO") ? "部門" :
                                   (vColName == "EMPNO") ? "工號" :
                                   (vColName == "TELNO1") ? "電話" :
                                   (vColName == "CELLPHONE") ? "行動電話" :
                                   (vColName == "ADDR1") ? "地址" :
                                   (vColName == "TITLE") ? "職稱" :
                                   (vColName == "WORKHR") ? "每日工時" :
                                   (vColName == "WORKSTATE") ? "每日排班" :
                                   (vColName == "ESCTYPE_04") ? "事假紀錄" :
                                   (vColName == "ESCTYPE_05") ? "病假紀錄" :
                                   (vColName == "MONTHPAY") ? "當月薪資" : "";
                    wsExcel.GetRow(vLinesNo).CreateCell(ColIndex).SetCellValue(vColHeadText);
                    wsExcel.GetRow(vLinesNo).GetCell(ColIndex).CellStyle = csTitle;
                }
                //寫入資料
                for (int DataIndex = 0; DataIndex < dtDriverStatus.Rows.Count; DataIndex++)
                {
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int FieldCount = 0; FieldCount < dtDriverStatus.Columns.Count; FieldCount++)
                    {
                        vColName = dtDriverStatus.Columns[FieldCount].ColumnName.ToUpper().Trim();
                        vTempStr = dtDriverStatus.Rows[DataIndex][FieldCount].ToString().Trim();
                        if (vTempStr != "")
                        {
                            if ((vColName == "DT") || (vColName == "BIRTHDAY") || (vColName == "ASSUMEDAY")) //日期欄位
                            {
                                vCalDate = DateTime.Parse(vTempStr);
                                wsExcel.GetRow(vLinesNo).CreateCell(FieldCount).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(FieldCount).SetCellValue((vCalDate.Year - 1911).ToString("D3") + "/" + vCalDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(FieldCount).CellStyle = csData;

                            }
                            else if ((vColName == "YEARCOUNT") || (vColName == "WORKHR") || (vColName == "MONTHPAY")) //數字欄位
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(FieldCount).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(FieldCount).SetCellValue(double.Parse(vTempStr));
                                wsExcel.GetRow(vLinesNo).GetCell(FieldCount).CellStyle = csData_Int;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(FieldCount).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(FieldCount).SetCellValue(vTempStr);
                                wsExcel.GetRow(vLinesNo).GetCell(FieldCount).CellStyle = csData;
                            }
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
                        string vEmpNoStr = eDriver_Search.Text.Trim();
                        string vCalYMStr = eCalYear_Search.Text.Trim() + " 年 " + ddlCalMonth_Search.SelectedItem.Text + "止 ";
                        string vRecordNote = "匯出檔案_駕駛員月出勤及請休假狀況表" + Environment.NewLine +
                                             "DriverStatus.aspx" + Environment.NewLine +
                                             "計算年月：" + vCalYMStr + Environment.NewLine +
                                             "員工工號：" + vEmpNoStr;
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
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
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
}