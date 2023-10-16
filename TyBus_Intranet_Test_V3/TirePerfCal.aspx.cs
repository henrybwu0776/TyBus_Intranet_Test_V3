using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;

namespace TyBus_Intranet_Test_V3
{
    public partial class TirePerfCal : System.Web.UI.Page
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
        DataTable tbResult = new DataTable();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                System.Threading.Thread.CurrentThread.CurrentCulture = Cal;

                DateTime vToday = DateTime.Today;

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
                        eCalYear_Search.Text = (vToday.Year - 1911).ToString();
                        eCalMonth_Search.Text = vToday.Month.ToString();
                        plResultData.Visible = false;
                    }
                    CaseDataBind();
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

        private void CaseDataBind()
        {
            string vSQLStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select a.TireNo, a.DepNo, (select name from Department where DepNo = a.DepNo) as DepName, a.DateS, a.DateE, a.TireKM, " + Environment.NewLine +
                      "       (select ClassTxt from DBDICB where fkey = '輪胎資料        Tires           BRAND' and ClassNo = b.Brand) as [Brand_Name], " + Environment.NewLine +
                      "       (select ClassTxt from DBDICB where fkey = '輪胎資料        Tires           SPEC'  and ClassNo = b.Spec)  as [Spec_Name], " + Environment.NewLine +
                      "       (select ClassTxt from DBDICB where fkey = '輪胎資料        Tires           TYPE'  and ClassNo = b.Type)  as [Type_Name] " + Environment.NewLine +
                      "  from TirePerf a, Tires b " + Environment.NewLine +
                      " where a.TireNo = b.TireNo ";
            int vTempYear = (eCalYear_Search.Text.Trim() != "") ? int.Parse(eCalYear_Search.Text.Trim()) : DateTime.Today.Year;
            vTempYear = (vTempYear < 1911) ? vTempYear + 1911 : vTempYear;
            int vTempMonth = (eCalMonth_Search.Text.Trim() != "") ? int.Parse(eCalMonth_Search.Text.Trim()) : DateTime.Today.Month;
            string vSearchYM = vTempYear.ToString("D4") + vTempMonth.ToString("D2");
            string vWStr_CalYM = " and Convert(varchar(6), a.DateE, 112) = '" + vSearchYM + "' ";
            string vWStr_TireNo = (eTireNo_Search.Text.Trim() != "") ? " and a.TireNo = '" + eTireNo_Search.Text.Trim() + "' " : "";
            vSQLStr = vSQLStr + Environment.NewLine +
                      vWStr_CalYM + Environment.NewLine +
                      vWStr_TireNo + Environment.NewLine +
                      " order by a.TireNo";
            sdsTirePerfCalMain.SelectCommand = "";
            sdsTirePerfCalMain.SelectCommand = vSQLStr;
            gridTirePerfCalMain.DataBind();
        }

        private void OutToOldExcel(string TireNo)
        {
            DateTime vTempDate;
            string vGetDate = "";
            //Excel 檔案
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
            fontTitle.FontHeightInPoints = 20; //字體大小
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

            int vLinesNo = 0;
            if (gridResultData.Rows.Count > 0)
            {
                wsExcel = (HSSFSheet)wbExcel.CreateSheet(TireNo);
                for (int i = 0; i < gridResultData.Rows.Count; i++)
                {
                    if (i == 0)
                    {
                        //建立標題列
                        wsExcel.CreateRow(i);
                        for (int j = 0; j < gridResultData.Columns.Count; j++)
                        {
                            wsExcel.GetRow(i).CreateCell(j).SetCellValue(gridResultData.Columns[j].HeaderText);
                            wsExcel.GetRow(i).GetCell(j).CellStyle = csTitle;
                        }
                    }
                    vLinesNo = i + 1;
                    wsExcel.CreateRow(vLinesNo);
                    for (int j = 0; j < gridResultData.Columns.Count; j++)
                    {
                        switch (j)
                        {
                            case 0:
                                wsExcel.GetRow(vLinesNo).CreateCell(j);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                                vTempDate = DateTime.Parse(gridResultData.Rows[i].Cells[j].Text.Trim());
                                vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(vGetDate);
                                wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
                                break;
                            case 1:
                                wsExcel.GetRow(vLinesNo).CreateCell(j);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(gridResultData.Rows[i].Cells[j].Text.Trim());
                                wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
                                break;
                            default:
                                wsExcel.GetRow(vLinesNo).CreateCell(j);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(double.Parse(gridResultData.Rows[i].Cells[j].Text.Trim()));
                                wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
                                break;
                        }
                    }
                }
                vLinesNo++;
                wsExcel.CreateRow(vLinesNo);
                wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("合計：");
                wsExcel.GetRow(vLinesNo).GetCell(1).CellStyle = csData;
                wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellFormula("sum(C2:C" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).GetCell(2).CellStyle = csData;
                wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellFormula("sum(D2:D" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).GetCell(3).CellStyle = csData;
                wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellFormula("sum(E2:E" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).GetCell(4).CellStyle = csData;
                wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellFormula("sum(F2:F" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).GetCell(5).CellStyle = csData;
                wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellFormula("sum(G2:G" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).GetCell(6).CellStyle = csData;
                wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellFormula("sum(H2:H" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).GetCell(7).CellStyle = csData;
                wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellFormula("sum(I2:I" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).GetCell(8).CellStyle = csData;
                wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellFormula("sum(J2:J" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).GetCell(9).CellStyle = csData;
                wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellFormula("sum(K2:K" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).GetCell(10).CellStyle = csData;

                try
                {
                    /*
                    MemoryStream msTarget = new MemoryStream();
                    wbExcel.Write(msTarget);
                    // 設定強制下載標頭
                    Response.AddHeader("Content-Disposition", string.Format("attachment; filename=輪胎_" + TireNo + "_公里數展開.xls"));
                    // 輸出檔案
                    Response.BinaryWrite(msTarget.ToArray());
                    msTarget.Close();

                    Response.End(); //*/
                    string vFileName = "輪胎_" + TireNo + "_公里數展開";
                    var msTarget = new NPOIMemoryStream();
                    msTarget.AllowClose = false;
                    wbExcel.Write(msTarget);
                    msTarget.Flush();
                    msTarget.Seek(0, SeekOrigin.Begin);
                    msTarget.AllowClose = true;

                    if (msTarget.Length > 0)
                    {
                        string vCalYMStr = eCalYear_Search.Text.Trim() + " 年 " + eCalMonth_Search.Text.Trim() + " 月";
                        string vRecordNote = "匯出檔案_輪胎公里數展開" + Environment.NewLine +
                                             "TirePerfCal.aspx" + Environment.NewLine +
                                             "結算年月：" + vCalYMStr + Environment.NewLine +
                                             "輪胎號碼：" + TireNo;
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

        private void OutToNewExcel(string TireNo)
        {
            DateTime vTempDate;
            string vGetDate = "";
            //Excel 檔案
            XSSFWorkbook wbExcel = new XSSFWorkbook();
            //Excel 工作表
            XSSFSheet wsExcel;
            //設定標題欄位的格式
            XSSFCellStyle csTitle = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中
                                                            //設定字體格式
            XSSFFont fontTitle = (XSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 20; //字體大小
            csTitle.SetFont(fontTitle);

            int vLinesNo = 0;
            if (gridResultData.Rows.Count > 0)
            {
                wsExcel = (XSSFSheet)wbExcel.CreateSheet(TireNo);
                for (int i = 0; i < gridResultData.Rows.Count; i++)
                {
                    if (i == 0)
                    {
                        //建立標題列
                        wsExcel.CreateRow(i);
                        for (int j = 0; j < gridResultData.Columns.Count; j++)
                        {
                            wsExcel.GetRow(i).CreateCell(j).SetCellValue(gridResultData.Columns[j].HeaderText);
                        }
                    }
                    vLinesNo = i + 1;
                    wsExcel.CreateRow(vLinesNo);
                    for (int j = 0; j < gridResultData.Columns.Count; j++)
                    {
                        switch (j)
                        {
                            case 0:
                                wsExcel.GetRow(vLinesNo).CreateCell(j);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                                vTempDate = DateTime.Parse(gridResultData.Rows[i].Cells[j].Text.Trim());
                                vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(vGetDate);
                                break;
                            case 1:
                                wsExcel.GetRow(vLinesNo).CreateCell(j);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(gridResultData.Rows[i].Cells[j].Text.Trim());
                                break;
                            default:
                                wsExcel.GetRow(vLinesNo).CreateCell(j);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(j).SetCellValue(double.Parse(gridResultData.Rows[i].Cells[j].Text.Trim()));
                                break;
                        }
                    }
                }
                vLinesNo++;
                wsExcel.CreateRow(vLinesNo);
                wsExcel.GetRow(vLinesNo).CreateCell(1).SetCellValue("合計：");
                wsExcel.GetRow(vLinesNo).CreateCell(2).SetCellFormula("sum(C2:C" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).CreateCell(3).SetCellFormula("sum(D2:D" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).CreateCell(4).SetCellFormula("sum(E2:E" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).CreateCell(5).SetCellFormula("sum(F2:F" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).CreateCell(6).SetCellFormula("sum(G2:G" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).CreateCell(7).SetCellFormula("sum(H2:H" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).CreateCell(8).SetCellFormula("sum(I2:I" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).CreateCell(9).SetCellFormula("sum(J2:J" + vLinesNo.ToString() + ")");
                wsExcel.GetRow(vLinesNo).CreateCell(10).SetCellFormula("sum(K2:K" + vLinesNo.ToString() + ")");

                try
                {
                    /*
                    MemoryStream msTarget = new MemoryStream();
                    wbExcel.Write(msTarget);
                    // 設定強制下載標頭
                    Response.AddHeader("Content-Disposition", string.Format("attachment; filename=輪胎_" + TireNo + "_公里數展開.xlsx"));
                    // 輸出檔案
                    Response.BinaryWrite(msTarget.ToArray());
                    msTarget.Close();

                    Response.End(); //*/
                    string vFileName = "輪胎_" + TireNo + "_公里數展開";
                    var msTarget = new NPOIMemoryStream();
                    msTarget.AllowClose = false;
                    wbExcel.Write(msTarget);
                    msTarget.Flush();
                    msTarget.Seek(0, SeekOrigin.Begin);
                    msTarget.AllowClose = true;

                    if (msTarget.Length > 0)
                    {
                        string vCalYMStr = eCalYear_Search.Text.Trim() + " 年 " + eCalMonth_Search.Text.Trim() + " 月";
                        string vRecordNote = "匯出檔案_輪胎公里數展開" + Environment.NewLine +
                                             "TirePerfCal.aspx" + Environment.NewLine +
                                             "結算年月：" + vCalYMStr + Environment.NewLine +
                                             "輪胎號碼：" + TireNo;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

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
                    Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            CaseDataBind();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void gridTirePerfCalMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            plResultData.Visible = false;
            tbResult.Clear();

            string vTireNo_Selected = gridTirePerfCalMain.SelectedRow.Cells[1].Text.Trim();
            string vKind_Old = "";
            string vKind = "";
            string vDateB = "";
            bool vKindState = false;
            bool vReNew = false;
            string vSQLStr = "";

            string vGetDataStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select x.BuDate,(select ClassTxt from DBDICB where fkey = '輪胎異動單B     TIREFIXB        KIND' and ClassNo = y.Kind) as KindName, " + Environment.NewLine +
                      "       y.Car_ID, y.TireNo, y.Kind " + Environment.NewLine +
                      "  from TirefixA x, TirefixB y " + Environment.NewLine +
                      " where x.FixNo = y.FixNo " + Environment.NewLine +
                      "   and y.TireNo = '" + vTireNo_Selected + "' " + Environment.NewLine +
                      "   and y.Kind in ('1', '2', '3', '4') " + Environment.NewLine +
                      " order by x.BuDate, y.FixNoItem";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                tbResult.Columns.Add("BuDate", typeof(DateTime));
                tbResult.Columns.Add("Car_ID", typeof(String));
                tbResult.Columns.Add("KM_Sum", typeof(Double));
                tbResult.Columns.Add("CBus1KM", typeof(Double));
                tbResult.Columns.Add("CBus2KM", typeof(Double));
                tbResult.Columns.Add("CRentTraKM", typeof(Double));
                tbResult.Columns.Add("CRentAKM", typeof(Double));
                tbResult.Columns.Add("CRentBKM", typeof(Double));
                tbResult.Columns.Add("CBus3KM", typeof(Double));
                tbResult.Columns.Add("CBus4KM", typeof(Double));
                tbResult.Columns.Add("CBus5KM", typeof(Double));
                SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vKind = drTemp["Kind"].ToString().Trim();
                    vReNew = ((vKind == "2") || (vKind == "4"));
                    if (vReNew)
                    {
                        vKindState = ((vKind_Old == "1") || (vKind_Old == "3"));
                        if ((vKind_Old != "") && (vKindState != true))
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('輪胎異動順序有誤，無法顯示！')");
                            Response.Write("</" + "Script>");
                        }
                        else
                        {
                            vGetDataStr = "select BuDate, Car_ID, (CBus1KM + CBus2KM + CRentAKM + CRentBKM + CRentTraKM + CBus3KM + CBus4KM + CBus5KM) KM_Sum, " + Environment.NewLine +
                                          "       CBus1KM, CBus2KM, CRentTraKM, CRentAKM, CRentBKM, CBus3KM, CBus4KM, CBus5KM " + Environment.NewLine +
                                          "  from RSCarTotal " + Environment.NewLine +
                                          " where Car_ID = '" + drTemp["Car_ID"].ToString().Trim() + "' " + Environment.NewLine +
                                          "   and BuDate between '" + vDateB +
                                          "' and '" + DateTime.Parse(drTemp["BuDate"].ToString().Trim()).Year.ToString("D4") + "/" + DateTime.Parse(drTemp["BuDate"].ToString().Trim()).ToString("MM/dd") + "' " + Environment.NewLine +
                                          " order by BuDate, Car_ID ";
                            using (SqlConnection connResult = new SqlConnection(vConnStr))
                            {
                                SqlCommand cmdResult = new SqlCommand(vGetDataStr, connResult);
                                connResult.Open();
                                SqlDataReader drResult = cmdResult.ExecuteReader();
                                while (drResult.Read())
                                {
                                    DataRow rowResult = tbResult.NewRow();
                                    rowResult["BuDate"] = DateTime.Parse(drResult["BuDate"].ToString()).Year.ToString("D4") + "/" + DateTime.Parse(drResult["BuDate"].ToString()).ToString("MM/dd");
                                    rowResult["Car_ID"] = drResult["Car_ID"].ToString();
                                    rowResult["KM_Sum"] = drResult["KM_Sum"].ToString();
                                    rowResult["CBus1KM"] = drResult["CBus1KM"].ToString();
                                    rowResult["CBus2KM"] = drResult["CBus2KM"].ToString();
                                    rowResult["CRentTraKM"] = drResult["CRentTraKM"].ToString();
                                    rowResult["CRentAKM"] = drResult["CRentAKM"].ToString();
                                    rowResult["CRentBKM"] = drResult["CRentBKM"].ToString();
                                    rowResult["CBus3KM"] = drResult["CBus3KM"].ToString();
                                    rowResult["CBus4KM"] = drResult["CBus4KM"].ToString();
                                    rowResult["CBus5KM"] = drResult["CBus5KM"].ToString();
                                    tbResult.Rows.Add(rowResult);
                                }
                                drResult.Close();
                                cmdResult.Cancel();
                            }
                        }
                    }
                    else
                    {
                        vKindState = ((vKind_Old == "2") || (vKind_Old == "4"));
                        if ((vKind_Old != "") && (vKindState != true))
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('輪胎異動順序有誤，無法顯示！')");
                            Response.Write("</" + "Script>");
                        }
                        else
                        {
                            vDateB = DateTime.Parse(drTemp["BuDate"].ToString()).Year.ToString("D4") + "/" + DateTime.Parse(drTemp["BuDate"].ToString()).ToString("MM/dd");
                        }
                    }
                    vKind_Old = drTemp["Kind"].ToString().Trim();
                }
                drTemp.Close();
                cmdTemp.Cancel();
            }
            if (tbResult.Rows.Count > 0)
            {
                gridResultData.DataSourceID = "";
                gridResultData.DataSource = tbResult;
                gridResultData.DataBind();
                plResultData.Visible = true;
            }
        }

        protected void bbOutToExcel_Click(object sender, EventArgs e)
        {
            switch (ddlExcelVersion.SelectedValue)
            {
                case "OldExcel":
                    OutToOldExcel(gridTirePerfCalMain.SelectedRow.Cells[1].Text.Trim());
                    break;
                case "NewExcel":
                    OutToNewExcel(gridTirePerfCalMain.SelectedRow.Cells[1].Text.Trim());
                    break;
            }
        }
    }
}