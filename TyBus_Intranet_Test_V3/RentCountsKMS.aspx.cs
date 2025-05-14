using Amaterasu_Function;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class RentCountsKMS : Page
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
        //private string vErrorMessage = "";
        private DataTable dtTarget = new DataTable();

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

                UnobtrusiveValidationMode = UnobtrusiveValidationMode.None;

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vDateURL = "InputDate.aspx?TextboxID=" + eSearchDate_S.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eSearchDate_S.Attributes["onClick"] = vDateScript;
                    vDateURL = "InputDate.aspx?TextboxID=" + eSearchDate_E.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eSearchDate_E.Attributes["onClick"] = vDateScript;

                    if (!IsPostBack)
                    {
                        DateTime dtLastMonth = DateTime.Today.AddMonths(-1);
                        eSearchDate_S.Text = PF.GetMonthFirstDay(dtLastMonth, "B");
                        eSearchDate_E.Text = PF.GetMonthLastDay(dtLastMonth, "B");
                        eDepNo.Text = vLoginDepNo;
                        eDepName.Text = vLoginDepName;
                        eDepNo.Enabled = ((Int32.Parse(vLoginDepNo) <= 9) || (vLoginDepNo == "30"));
                        GetCarType();
                        ddlCarType.SelectedIndex = 0;
                        eCarType.Text = ddlCarType.Items[0].Value.Trim();
                        rbDataShowType.SelectedIndex = 0;
                        eDataShowType.Text = rbDataShowType.Items[0].Value.Trim();
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

        private void GetCarType()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                string vTempStr = "select cast('0' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                  " union all " + Environment.NewLine +
                                  "select ClassNo, ClassTxt from DBDICB where FKey = '行車記錄單b     runsheetb       CARTYPE'";
                SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    ddlCarType.Items.Clear();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                        ddlCarType.Items.Add(liTemp);
                    }
                }
            }
        }

        private string GetSelectStr()
        {
            string vResultStr = "";
            string vWStr_LinesNo = "";
            switch (ddlCarType.SelectedValue.Trim())
            {
                case "0": //不選擇車種
                    if (rbDataShowType.SelectedValue == "0")
                    {
                        vWStr_LinesNo = (eLinesNo.Text.Trim() != "") ? "          and b.LinesNo = '" + eLinesNo.Text.Trim() + "' " + Environment.NewLine : "";
                        vResultStr = "select t.BuDate 行車日期, t.DepNo 站別代碼, t.DepName 站別, t.Car_ID 牌照號碼, t.Driver 駕駛員工號, t.DriverName 駕駛員, t.CellPhone 手機號碼, " + Environment.NewLine +
                                     "       sum(t.NoneBussKM) 非營運里程, sum(t.TotalKM) 營運里程, sum(t.RCount) 趟次 " + Environment.NewLine +
                                     "  from (" + Environment.NewLine +
                                     "       select a.BuDate, a.DepNo, c.[Name] as DepName, b.Car_ID, a.Driver, f.[Name] as DriverName, isnull(f.Cellphone, '') as CellPhone, " + Environment.NewLine +
                                     "              cast(0 as float) as NoneBussKM, sum(b.ActualKM) as TotalKM, count(b.AssignNo) as RCount " + Environment.NewLine +
                                     "         from RunSheetA as a left join RunSheetB as b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                     "                             left join Department as c on c.DepNo = a.DepNo " + Environment.NewLine +
                                     "                             left join Employee as f on f.EmpNo = a.Driver " + Environment.NewLine +
                                     "        where a.BuDate between @StartDate and @EndDate " + Environment.NewLine +
                                     "          and a.DepNo = @DepNo " + Environment.NewLine +
                                     "          and b.CarType != '9' " + Environment.NewLine +
                                     "          and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                     vWStr_LinesNo+
                                     "        group by a.BuDate, a.DepNo, c.[Name], b.Car_ID, a.Driver, f.[Name], isnull(f.Cellphone, '') " + Environment.NewLine +
                                     "        union all " + Environment.NewLine +
                                     "       select a.BuDate, a.DepNo, c.[Name] as DepName, b.Car_ID, a.Driver, f.[Name] as DriverName, isnull(f.Cellphone, '') as CellPhone, " + Environment.NewLine +
                                     "              sum(b.ActualKM) as NoneBussKM, cast(0 as float) as TotalKM, count(b.AssignNo) as RCount " + Environment.NewLine +
                                     "         from RunSheetA as a left join RunSheetB as b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                     "                             left join Department as c on c.DepNo = a.DepNo " + Environment.NewLine +
                                     "                             left join Employee as f on f.EmpNo = a.Driver " + Environment.NewLine +
                                     "        where a.BuDate between @StartDate and @EndDate " + Environment.NewLine +
                                     "          and a.DepNo = @DepNo " + Environment.NewLine +
                                     "          and b.CarType = '9' " + Environment.NewLine +
                                     "          and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                     vWStr_LinesNo +
                                     "        group by a.BuDate, a.DepNo, c.[Name], b.Car_ID, a.Driver, f.[Name], isnull(f.Cellphone, '') " + Environment.NewLine +
                                     ") t group by t.BuDate, t.DepNo, t.DepName, t.Car_ID, t.Driver, t.DriverName, t.CellPhone " + Environment.NewLine +
                                     " order by t.BuDate, t.DepNo, t.Car_ID ";
                    }
                    else
                    {
                        vWStr_LinesNo = (eLinesNo.Text.Trim() != "") ? "   and b.LinesNo = '" + eLinesNo.Text.Trim() + "' " + Environment.NewLine : "";
                        vResultStr = "select a.BuDate 行車日期, a.DepNo 站別代碼, c.[Name] 站別, b.LinesNo 路線編號, e.LineName 路線名稱, " + Environment.NewLine +
                                     "       b.ToTime 去程時間, b.ToLine 去程路線, b.BackTime 回程時間, b.BackLine 回程路線, b.Car_ID 牌照號碼, " + Environment.NewLine +
                                     "       d.ClassTXT 車種, a.Driver 駕駛員工號, f.[Name] 駕駛員, isnull(f.Cellphone, '') 手機號碼, b.ActualKM 行駛里程 " + Environment.NewLine +
                                     "  from RunSheetA as a left join RunSheetB as b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                     "                      left join Department as c on c.DepNo = a.DepNo " + Environment.NewLine +
                                     "                      left join DBDICB as d on d.ClassNo = b.CarType and d.FKey = '行車記錄單b     runsheetb       CARTYPE' " + Environment.NewLine +
                                     "                      left join Lines as e on e.LinesNo = b.LinesNo " + Environment.NewLine +
                                     "                      left join Employee as f on f.EmpNo = a.Driver " + Environment.NewLine +
                                     " where a.BuDate between @StartDate and @EndDate " + Environment.NewLine +
                                     "   and a.DepNo = @DepNo " + Environment.NewLine +
                                     "   and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                     vWStr_LinesNo +
                                     " order by a.BuDate, a.DepNo, b.Car_ID, b.ToTime, b.LinesNo ";
                    }
                    break;
                default: //有選擇車種
                    if (rbDataShowType.SelectedValue == "0")
                    {
                        vWStr_LinesNo = (eLinesNo.Text.Trim() != "") ? "   and b.LinesNo = '" + eLinesNo.Text.Trim() + "' " + Environment.NewLine : "";
                        vResultStr = "select a.BuDate 行車日期, a.DepNo 站別代碼, c.[Name] 站別, b.LinesNo 路線編號, e.LineName 路線名稱, " + Environment.NewLine +
                                     "       b.ToLine 去程路線, b.BackLine 回程路線, b.Car_ID 牌照號碼, d.ClassTXT 車種, a.Driver 駕駛員工號, f.[Name] 駕駛員, " + Environment.NewLine +
                                     "       isnull(f.Cellphone, '') 手機號碼, sum(b.ActualKM) 里程小計, count(b.AssignNo) 趟次 " + Environment.NewLine +
                                     "  from RunSheetA as a left join RunSheetB as b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                     "                      left join Department as c on c.DepNo = a.DepNo " + Environment.NewLine +
                                     "                      left join DBDICB as d on d.ClassNo = b.CarType and d.FKey = '行車記錄單b     runsheetb       CARTYPE' " + Environment.NewLine +
                                     "                      left join Lines as e on e.LinesNo = b.LinesNo " + Environment.NewLine +
                                     "                      left join Employee as f on f.EmpNo = a.Driver " + Environment.NewLine +
                                     " where a.BuDate between @StartDate and @EndDate " + Environment.NewLine +
                                     "   and a.DepNo = @DepNo " + Environment.NewLine +
                                     "   and b.CarType = @CarType " + Environment.NewLine +
                                     "   and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                     vWStr_LinesNo +
                                     " group by a.BuDate, a.DepNo, c.[Name], b.LinesNo, e.LineName, b.ToLine, " + Environment.NewLine +
                                     "       b.BackLine, b.Car_ID, d.ClassTXT, a.Driver, f.[Name], isnull(f.Cellphone, '') " + Environment.NewLine +
                                     " order by a.BuDate, a.DepNo, b.Car_ID, b.LinesNo ";
                    }
                    else
                    {
                        vWStr_LinesNo = (eLinesNo.Text.Trim() != "") ? "   and b.LinesNo = '" + eLinesNo.Text.Trim() + "' " + Environment.NewLine : "";
                        vResultStr = "select a.BuDate 行車日期, a.DepNo 站別代碼, c.[Name] 站別, b.LinesNo 路線編號, e.LineName 路線名稱, " + Environment.NewLine +
                                     "       b.ToTime 去程時間, b.ToLine 去程路線, b.BackTime 回程時間, b.BackLine 回程路線, b.Car_ID 牌照號碼, " + Environment.NewLine +
                                     "       d.ClassTXT 車種, a.Driver 駕駛員工號, f.[Name] 駕駛員, isnull(f.Cellphone, '') 手機號碼, b.ActualKM 行駛里程 " + Environment.NewLine +
                                     "  from RunSheetA as a left join RunSheetB as b on b.AssignNo = a.AssignNo " + Environment.NewLine +
                                     "                      left join Department as c on c.DepNo = a.DepNo " + Environment.NewLine +
                                     "                      left join DBDICB as d on d.ClassNo = b.CarType and d.FKey = '行車記錄單b     runsheetb       CARTYPE' " + Environment.NewLine +
                                     "                      left join Lines as e on e.LinesNo = b.LinesNo " + Environment.NewLine +
                                     "                      left join Employee as f on f.EmpNo = a.Driver " + Environment.NewLine +
                                     " where a.BuDate between @StartDate and @EndDate " + Environment.NewLine +
                                     "   and a.DepNo = @DepNo " + Environment.NewLine +
                                     "   and isnull(b.ReduceReason, '') = '' " + Environment.NewLine +
                                     "   and b.CarType = @CarType " + Environment.NewLine +
                                     vWStr_LinesNo +
                                     " order by a.BuDate, a.DepNo, b.Car_ID, b.ToTime, b.LinesNo ";
                    }
                    break;
            }
            return vResultStr;
        }

        private DataTable OpenData()
        {
            DataTable dtResult = new DataTable();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSelectStr();
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelectStr, connTemp);
                daTemp.SelectCommand.Parameters.Clear();
                daTemp.SelectCommand.Parameters.Add(new SqlParameter("StartDate", eSearchDate_S.Text.Trim() + " 00:00:00"));
                daTemp.SelectCommand.Parameters.Add(new SqlParameter("EndDate", eSearchDate_E.Text.Trim() + " 23:59:59"));
                daTemp.SelectCommand.Parameters.Add(new SqlParameter("DepNo", eDepNo.Text.Trim()));
                if (ddlCarType.SelectedValue != "0")
                {
                    daTemp.SelectCommand.Parameters.Add(new SqlParameter("CarType", ddlCarType.SelectedValue.Trim()));
                }
                connTemp.Open();
                daTemp.Fill(dtResult);
            }
            return dtResult;
        }

        protected void eDepNo_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo.Text.Trim();
            string vDepName = "";
            string vTempStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            vDepName = PF.GetValue(vConnStr, vTempStr, "Name");
            if (vDepName.Trim() == "")
            {
                vDepName = vDepNo.Trim();
                vTempStr = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' ";
                vDepNo = PF.GetValue(vConnStr, vTempStr, "DepNo");
            }
            eDepNo.Text = vDepNo;
            eDepName.Text = vDepName;
        }

        protected void ddlCarType_SelectedIndexChanged(object sender, EventArgs e)
        {
            eCarType.Text = ddlCarType.SelectedValue.Trim();
        }

        protected void rbDataShowType_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDataShowType.Text = rbDataShowType.SelectedValue.Trim();
        }

        protected void bbShowData_Click(object sender, EventArgs e)
        {
            plShowData.Visible = false;
            dtTarget = OpenData();
            if (dtTarget.Rows.Count > 0)
            {
                plShowData.Visible = true;
                gvShowData.DataSource = dtTarget;
                gvShowData.DataBind();
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            DateTime vBuDate;
            double vTempFloat = 0;
            dtTarget = OpenData();
            if (dtTarget.Rows.Count > 0)
            {
                //準備匯出EXCEL
                XSSFWorkbook wbExcel = new XSSFWorkbook();
                //Excel 工作表
                XSSFSheet wsExcel;

                //設定標題欄位的格式
                XSSFCellStyle csTitle = (XSSFCellStyle)wbExcel.CreateCellStyle();
                csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csTitle.Alignment = HorizontalAlignment.Center; //水平置中

                //設定標題欄位的字體格式
                XSSFFont fontTitle = (XSSFFont)wbExcel.CreateFont();
                //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
                fontTitle.IsBold = true;
                fontTitle.FontHeightInPoints = 12; //字體大小
                csTitle.SetFont(fontTitle);

                //設定資料內容欄位的格式
                XSSFCellStyle csData = (XSSFCellStyle)wbExcel.CreateCellStyle();
                csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                XSSFFont fontData = (XSSFFont)wbExcel.CreateFont();
                fontData.FontHeightInPoints = 12;
                csData.SetFont(fontData);

                XSSFCellStyle csData_Red = (XSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                XSSFFont fontData_Red = (XSSFFont)wbExcel.CreateFont();
                fontData_Red.FontHeightInPoints = 12; //字體大小
                fontData_Red.Color = IndexedColors.Red.Index; //字體顏色
                csData_Red.SetFont(fontData_Red);

                XSSFCellStyle csData_Int = (XSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                XSSFDataFormat format = (XSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Int.DataFormat = format.GetFormat("###,##0");

                XSSFCellStyle csData_Float = (XSSFCellStyle)wbExcel.CreateCellStyle();
                csData_Float.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                csData_Float.Alignment = HorizontalAlignment.Right;

                XSSFDataFormat format_Float = (XSSFDataFormat)wbExcel.CreateDataFormat();
                csData_Float.DataFormat = format.GetFormat("##0.00");

                string vHeaderText = "";

                string vFileName = "車輛趟次里程統計作業";
                int vLinesNo = 0;

                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                //新增一個工作表
                wsExcel = (XSSFSheet)wbExcel.CreateSheet(vFileName);
                //寫入標題列
                vLinesNo = 0;
                wsExcel.CreateRow(vLinesNo);
                for (int i = 0; i < dtTarget.Columns.Count; i++)
                {
                    vHeaderText = dtTarget.Columns[i].ColumnName.Trim();
                    wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                }
                for (int i = 0; i < dtTarget.Rows.Count; i++)
                {
                    vLinesNo++;
                    wsExcel.CreateRow(vLinesNo);
                    for (int j = 0; j < dtTarget.Columns.Count; j++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(j);
                        if (dtTarget.Columns[j].ColumnName == "行車日期")
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue((DateTime.TryParse(dtTarget.Rows[i][j].ToString().Trim(), out vBuDate)) ? vBuDate.ToShortDateString() : String.Empty);
                            wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
                        }
                        else if ((dtTarget.Columns[j].ColumnName == "行駛里程") ||
                                 (dtTarget.Columns[j].ColumnName == "里程小計") ||
                                 (dtTarget.Columns[j].ColumnName == "趟次") ||
                                 (dtTarget.Columns[j].ColumnName == "營運里程") ||
                                 (dtTarget.Columns[j].ColumnName == "非營運里程"))
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.Numeric);
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue((double.TryParse(dtTarget.Rows[i][j].ToString().Trim(), out vTempFloat)) ? vTempFloat.ToString() : "0");
                            wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData_Float;
                        }
                        else
                        {
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(dtTarget.Rows[i][j].ToString().Trim());
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
                        //先判斷是不是IE
                        HttpContext.Current.Response.ContentType = "application/octet-stream";
                        HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                        string TourVerision = brObject.Type;

                        if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                        {
                            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xlsx", System.Text.Encoding.UTF8));
                        }
                        else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                        {
                            // 設定強制下載標頭
                            Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xlsx"));
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
                    //throw;
                }
            }
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}