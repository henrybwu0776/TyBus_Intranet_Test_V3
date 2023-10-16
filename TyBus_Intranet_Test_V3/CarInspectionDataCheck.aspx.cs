using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI;


namespace TyBus_Intranet_Test_V3
{
    public partial class CarInspectionDataCheck : Page
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

                    string vCheckDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCheckDateS_Search.ClientID;
                    string vCheckDateScript = "window.open('" + vCheckDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCheckDateS_Search.Attributes["onClick"] = vCheckDateScript;

                    vCheckDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCheckDateE_Search.ClientID;
                    vCheckDateScript = "window.open('" + vCheckDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCheckDateE_Search.Attributes["onClick"] = vCheckDateScript;

                    if (!IsPostBack)
                    {
                        eCheckDateS_Search.Text = PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C");
                        eCheckDateE_Search.Text = PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C");
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
        /// 產生查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            DateTime vCheckDate_S = (eCheckDateS_Search.Text.Trim() != "") ? DateTime.Parse(eCheckDateS_Search.Text.Trim()) : DateTime.Today;
            string vCheckDateStr_S = vCheckDate_S.Year.ToString("D4") + "/" + vCheckDate_S.Month.ToString("D2") + "/" + vCheckDate_S.Day.ToString("D2");
            DateTime vCheckDate_E = (eCheckDateE_Search.Text.Trim() != "") ? DateTime.Parse(eCheckDateE_Search.Text.Trim()) : DateTime.Today;
            string vCheckDateStr_E = vCheckDate_E.Year.ToString("D4") + "/" + vCheckDate_E.Month.ToString("D2") + "/" + vCheckDate_E.Day.ToString("D2");
            string vWStr_CarID = (eCarID_Search.Text.Trim() != "") ? "   and c.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and a.CompanyNo between '" + eDepNoS_Search.Text.Trim() + "' and '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? "   and a.CompanyNo = '" + eDepNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? "   and a.CompanyNo = '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CheckDate = ((eCheckDateS_Search.Text.Trim() != "") && (eCheckDateE_Search.Text.Trim() != "")) ? "   and c.CheckDate between '" + vCheckDateStr_S + "' and '" + vCheckDateStr_E + "' " + Environment.NewLine :
                                     ((eCheckDateS_Search.Text.Trim() != "") && (eCheckDateE_Search.Text.Trim() == "")) ? "   and c.CheckDate = '" + vCheckDateStr_S + "' " + Environment.NewLine :
                                     ((eCheckDateS_Search.Text.Trim() == "") && (eCheckDateE_Search.Text.Trim() != "")) ? "   and c.CheckDate = '" + vCheckDateStr_E + "' " + Environment.NewLine : "";
            string vResultStr = "select c.Car_ID, " + Environment.NewLine +
                                "       case when year(c.GDate) < 1911 then null else c.GDate end GDate, " + Environment.NewLine +
                                "       case when year(c.BespeakDate) < 1911 then null else c.BespeakDate end BespeakDate, " + Environment.NewLine +
                                "       case when year(c.CheckDate) < 1911 then null else c.CheckDate end CheckDate, " + Environment.NewLine +
                                "       case when year(a.NextEDate) < 1911 then null else a.NextEDate end NextEDate, " + Environment.NewLine +
                                "       case when year(a.NCheckTerm) < 1911 then null else a.NCheckTerm end NCheckTerm, " + Environment.NewLine +
                                "       (select [Name] from Department where DepNo = a.CompanyNo) DepName, " + Environment.NewLine +
                                "       (select [Name] from [Custom] where [Types] = 'S' and Code = c.StoreNo) StoreName " + Environment.NewLine +
                                "  from Car_InfoC c left join Car_InfoA a on a.Car_ID = c.Car_ID " + Environment.NewLine +
                                " where isnull(c.Car_ID, '') <> '' " + Environment.NewLine +
                                vWStr_CarID + vWStr_DepNo + vWStr_CheckDate +
                                " order by a.CompanyNo, c.GDate, c.Car_ID";
            return vResultStr;
        }

        /// <summary>
        /// 開啟資料
        /// </summary>
        private void OpenData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelStr = GetSelectStr();
            sdsShowData.SelectCommand = "";
            sdsShowData.SelectCommand = vSelStr;
            gridShowData.DataBind();
        }

        protected void eDepNoS_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNoS_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNoS_Search.Text = vDepNo_Temp.Trim();
            eDepNameS_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDepNoE_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNoE_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNoE_Search.Text = vDepNo_Temp.Trim();
            eDepNameE_Search.Text = vDepName_Temp.Trim();
        }

        /// <summary>
        /// 查詢資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbShowData_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
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
            DateTime vBuDate;

            string vSelectStr = GetSelectStr();
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
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet("檢驗資料");
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "牌照號碼" :
                                      (drExcel.GetName(i).ToUpper() == "GDATE") ? "領照日期" :
                                      (drExcel.GetName(i).ToUpper() == "BESPEAKDATE") ? "預定檢驗日" :
                                      (drExcel.GetName(i).ToUpper() == "CHECKDATE") ? "完成檢驗日" :
                                      (drExcel.GetName(i).ToUpper() == "NEXTEDATE") ? "下次檢驗日" :
                                      (drExcel.GetName(i).ToUpper() == "NCHECKTERM") ? "下次檢驗期限" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "STORENAME") ? "檢驗廠商" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i).ToUpper() == "GDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "BESPEAKDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "CHECKDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "NEXTEDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "NCHECKTERM")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                if (vBuDate.Year > 3822)
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 3822).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                                }
                                else if (vBuDate.Year > 1911)
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 1911).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
                                }
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
                            string vRecordCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "全部";
                            string vRecordDepNoStr = ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoS_Search.Text.Trim() + "~" + eDepNoE_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() != "") && (eDepNoE_Search.Text.Trim() == "")) ? eDepNoS_Search.Text.Trim() :
                                                     ((eDepNoS_Search.Text.Trim() == "") && (eDepNoE_Search.Text.Trim() != "")) ? eDepNoE_Search.Text.Trim() : "全部";
                            string vRecordDateStr = ((eCheckDateS_Search.Text.Trim() != "") && (eCheckDateE_Search.Text.Trim() != "")) ? eCheckDateS_Search.Text.Trim() + "~" + eCheckDateE_Search.Text.Trim() :
                                                    ((eCheckDateS_Search.Text.Trim() != "") && (eCheckDateE_Search.Text.Trim() == "")) ? eCheckDateS_Search.Text.Trim() :
                                                    ((eCheckDateS_Search.Text.Trim() == "") && (eCheckDateE_Search.Text.Trim() != "")) ? eCheckDateE_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_車輛檢驗資料查詢" + Environment.NewLine +
                                                 "CarInspectionDataCheck.aspx" + Environment.NewLine +
                                                 "車號：" + vRecordCarIDStr + Environment.NewLine +
                                                 "站別：" + vRecordDepNoStr + Environment.NewLine +
                                                 "最後檢驗日：" + vRecordDateStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            string vFileName = "車輛檢驗資料查詢";
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
        /// 關閉畫面回到主頁
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}