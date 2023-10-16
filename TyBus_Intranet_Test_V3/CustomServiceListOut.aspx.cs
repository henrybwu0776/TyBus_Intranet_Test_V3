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
    public partial class CustomServiceListOut : System.Web.UI.Page
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

                if (vLoginID != null)
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    string vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_S.ClientID;
                    string vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuildDate_S.Attributes["onClick"] = vTempDateScript;

                    vTempDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_E.ClientID;
                    vTempDateScript = "window.open('" + vTempDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuildDate_E.Attributes["onClick"] = vTempDateScript;

                    if (!IsPostBack)
                    {
                        eBuildDate_S.Text = "";
                        eBuildDate_E.Text = "";
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

        private string GetSelectStr()
        {
            string vBuildDate_S = (eBuildDate_S.Text.Trim() != "") ? PF.GetAD(eBuildDate_S.Text.Trim()) : "";
            string vBuildDate_E = (eBuildDate_E.Text.Trim() != "") ? PF.GetAD(eBuildDate_E.Text.Trim()) : "";
            string vWStr_BuildDate = ((eBuildDate_S.Text.Trim() != "") && (eBuildDate_E.Text.Trim() != "")) ? " where BuildDate between '" + vBuildDate_S + " 00:00:00' and '" + vBuildDate_E + " 23:59:59' " + Environment.NewLine :
                                     ((eBuildDate_S.Text.Trim() != "") && (eBuildDate_E.Text.Trim() == "")) ? " where BuildDate >= '" + vBuildDate_S + " 00:00:00' " + Environment.NewLine :
                                     ((eBuildDate_S.Text.Trim() == "") && (eBuildDate_E.Text.Trim() != "")) ? " where BuildDate <= '" + vBuildDate_E + " 23:59:59' " + Environment.NewLine : "";
            string vSelStr = "SELECT [ServiceNo] [客訴單號], [BuildDate] [開單日期], [BuildTime] [開單時間], " + Environment.NewLine +
                             "       (select [Name] from Employee where EMpNo = a.[BuildMan]) [開單人], " + Environment.NewLine +
                             "       [ServiceDate] [發生日期], (select ClassTxt from DBDICB where FKey = '客服專線記錄表  fmCustomService CaseSource' and CLassNo = a.CaseSource) [案件來源], " + Environment.NewLine +
                             "       (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) as 反映事項大類, " + Environment.NewLine +
                             "       (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) as  反映事項中類, " + Environment.NewLine +
                             "       (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) as 反映事項明細, " + Environment.NewLine +
                             "       [LinesNo] as [路線代號], [Car_ID] as [牌照號碼], [Driver] as [駕駛員編號], [DriverName] as [駕駛員], " + Environment.NewLine +
                             "       [BoardTime] as [上車時間], [BoardStation] as [上車站牌], [GetoffTime] as [下車時間], [GetoffStation] as [下車站牌], " + Environment.NewLine +
                             "       [ServiceNote] as [反映事項概述], [CivicName] as [民眾姓名], [CivicTelNo] as [民眾電話], [CivicCellPhone] as [民眾手機], " + Environment.NewLine +
                             "       [CivicAddress] as [民眾地址], [CivicEMail] as [民眾EMAIL], (select[Name] from Department where DepNo = a.[AthorityDepNo]) as [權責單位], " + Environment.NewLine +
                             "       [AthorityDepNote] as [權責單位說明], " + Environment.NewLine +
                             "       case when isnull([IsReplied], 0) = 1 then '已回應' else '尚未回應' end as [是否已回應], [Remark] as [回應概述], " + Environment.NewLine +
                             "       case when isnull([IsPending], 0) = 1 then '是' else '否' end as [是否分發待查], " + Environment.NewLine +
                             "       case when isnull([IsTrue], '') = '' then '待查證' " + Environment.NewLine +
                             "            else (select ClassTxt from DBDICB where FKey = '客服專線記錄表  fmCustomService IsTrue' and CLassNo = [IsTrue]) end [查證情況], " + Environment.NewLine +
                             "       [AssignDate] as [受理日期], [AssignMan] as [受理人], " + Environment.NewLine +
                             "       case when isnull([IsClosed], 0) = 1 then '已結案' else '尚未結案' end as [是否結案], " + Environment.NewLine +
                             "       [CloseDate] as [結案日期], (select [Name] from Employee where EMpNo = a.[CloseMan]) as [結案人] " + Environment.NewLine +
                             "  FROM CustomService a " + Environment.NewLine +
                             vWStr_BuildDate +
                             " order by ServiceType, ServiceNo";
            return vSelStr;
        }

        private string GetSelectStr_Excel2()
        {
            string vBuildDate_S = (eBuildDate_S.Text.Trim() != "") ? PF.GetAD(eBuildDate_S.Text.Trim()) : "";
            string vBuildDate_E = (eBuildDate_E.Text.Trim() != "") ? PF.GetAD(eBuildDate_E.Text.Trim()) : "";
            string vWStr_BuildDate = ((eBuildDate_S.Text.Trim() != "") && (eBuildDate_E.Text.Trim() != "")) ? "          and a.BuildDate between '" + vBuildDate_S + " 00:00:00' and '" + vBuildDate_E + " 23:59:59' " + Environment.NewLine :
                                     ((eBuildDate_S.Text.Trim() != "") && (eBuildDate_E.Text.Trim() == "")) ? "          and a.BuildDate >= '" + vBuildDate_S + " 00:00:00' " + Environment.NewLine :
                                     ((eBuildDate_S.Text.Trim() == "") && (eBuildDate_E.Text.Trim() != "")) ? "          and a.BuildDate <= '" + vBuildDate_E + " 23:59:59' " + Environment.NewLine : "";
            string vSelStr = "select (select ClassTxt from DBDICB where FKey = '客服專線記錄表  fmCustomService CaseSource' and CLassNo = ta.CaseSource) [案件來源], Count(ta.CaseSource) [數量], " + Environment.NewLine +
                             "       sum(IT01)[待查證], sum(IT02)[經查屬實], sum(IT03)[查無實據], sum(IT04)[無影像] " + Environment.NewLine +
                             "  from( " + Environment.NewLine +
                             "       select CaseSource, case when IsTrue = 'IT01' then 1 else 0 end IT01, " + Environment.NewLine +
                             "                          case when IsTrue = 'IT02' then 1 else 0 end IT02, " + Environment.NewLine +
                             "                          case when IsTrue = 'IT03' then 1 else 0 end IT03, " + Environment.NewLine +
                             "                          case when IsTrue = 'IT04' then 1 else 0 end IT04 " + Environment.NewLine +
                             "         from CustomService a " + Environment.NewLine +
                             "        where a.ServiceType = 'S04' " + Environment.NewLine +
                             "          and isnull(IsTrue, '') <> '' " + Environment.NewLine +
                             vWStr_BuildDate +
                             "  ) ta group by ta.CaseSource";
            return vSelStr;
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelectStr();
            string vFileName = "客服系統資料";
            int vLinesNo = 0;

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
                    DateTime vBuDate;

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


                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);

                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(drExcel.GetName(i));
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    vLinesNo++;
                    while (drExcel.Read())
                    {
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if (((drExcel.GetName(i) == "開單日期") ||
                                 (drExcel.GetName(i) == "受理日期") ||
                                 (drExcel.GetName(i) == "結案日期") ||
                                 (drExcel.GetName(i) == "發生日期")) && (drExcel[i].ToString() != ""))
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
                    }
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbExcel_2_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelectStr_Excel2();
            string vFileName = "各站客訴分類案件數量";
            int vLinesNo = 0;

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
                    //DateTime vBuDate;

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


                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);

                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(drExcel.GetName(i));
                        wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                    }
                    vLinesNo++;
                    while (drExcel.Read())
                    {
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            if (((drExcel.GetName(i) == "數量") ||
                                 (drExcel.GetName(i) == "待查證") ||
                                 (drExcel.GetName(i) == "經查屬實") ||
                                 (drExcel.GetName(i) == "查無實據") ||
                                 (drExcel.GetName(i) == "無影像")) && (drExcel[i].ToString() != ""))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
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
                    }
                }
            }
        }
    }
}