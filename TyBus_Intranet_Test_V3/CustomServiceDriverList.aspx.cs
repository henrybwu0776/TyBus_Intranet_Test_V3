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
    public partial class CustomServiceDriverList : System.Web.UI.Page
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

                string vTempDate_URL = "InputDate_ChineseYears.aspx?TextboxID=" + eServiceDate_Search_S.ClientID;
                string vTempDate_Script = "window.open('" + vTempDate_URL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eServiceDate_Search_S.Attributes["onClick"] = vTempDate_Script;

                vTempDate_URL = "InputDate_ChineseYears.aspx?TextboxID=" + eServiceDate_Search_E.ClientID;
                vTempDate_Script = "window.open('" + vTempDate_URL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eServiceDate_Search_E.Attributes["onClick"] = vTempDate_Script;

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
                        plReport.Visible = false;
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
            string vResultStr = "";
            string vWStr_ServiceType = "";
            if (rbServiceType_Search.SelectedValue == "")
            {
                vWStr_ServiceType = "                  and ServiceType in ('S04','S06')" + Environment.NewLine;
            }
            else
            {
                vWStr_ServiceType = "                  and ServiceType = '" + rbServiceType_Search.SelectedValue + "' " + Environment.NewLine;
            }
            string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? "   and a.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            /* 暫時先不過濾開單日期
            string vWStr_BuildDate = ((eBuildDate_Search_S.Text.Trim() != "") && (eBuildDate_Search_E.Text.Trim() != "")) ? "   and a.BuildDate between '" + (DateTime.Parse(eBuildDate_Search_S.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_Search_S.Text.Trim()).ToString("MM/dd") + "' and '" + (DateTime.Parse(eBuildDate_Search_E.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_Search_E.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eBuildDate_Search_S.Text.Trim() != "") && (eBuildDate_Search_E.Text.Trim() == "")) ? "   and a.BuildDate = '" + (DateTime.Parse(eBuildDate_Search_S.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_Search_S.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eBuildDate_Search_S.Text.Trim() == "") && (eBuildDate_Search_E.Text.Trim() != "")) ? "   and a.BuildDate = '" + (DateTime.Parse(eBuildDate_Search_E.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_Search_E.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";
            */
            string vWStr_ServiceDate = ((eServiceDate_Search_S.Text.Trim() != "") && (eServiceDate_Search_E.Text.Trim() != "")) ? "   and a.ServiceDate between '" + (DateTime.Parse(eServiceDate_Search_S.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_Search_S.Text.Trim()).ToString("MM/dd") + "' and '" + (DateTime.Parse(eServiceDate_Search_E.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_Search_E.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                       ((eServiceDate_Search_S.Text.Trim() != "") && (eServiceDate_Search_E.Text.Trim() == "")) ? "   and a.ServiceDate = '" + (DateTime.Parse(eServiceDate_Search_S.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_Search_S.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                       ((eServiceDate_Search_S.Text.Trim() == "") && (eServiceDate_Search_E.Text.Trim() != "")) ? "   and a.ServiceDate = '" + (DateTime.Parse(eServiceDate_Search_E.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_Search_E.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";
            vResultStr = "select * from (" + Environment.NewLine +
                         "               SELECT AthorityDepNo, AthorityDepNote, Driver, DriverName, BuildDate, ServiceDate, LinesNo, " + Environment.NewLine +
                         "                      ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = 1) AND (TypeLevel1 = a.ServiceType)) AS ServiceType_C, " + Environment.NewLine +
                         "                      ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = 2) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS ServiceTypeB_C, " + Environment.NewLine +
                         "                      ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = 3) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS ServiceTypeC_C " + Environment.NewLine +
                         "                 FROM CustomService AS a " + Environment.NewLine +
                         "                WHERE isnull(Driver, '') <> '' " + Environment.NewLine +
                         "                  AND isnull(IsTrue, '') = 'V' " + Environment.NewLine +
                         vWStr_ServiceType +
                         vWStr_Driver +
                         vWStr_ServiceDate +
                         "                union all " + Environment.NewLine +
                         "               SELECT AthorityDepNo, AthorityDepNote, Driver2 as Driver, DriverName2 as DriverName, BuildDate, ServiceDate, LinesNo, " + Environment.NewLine +
                         "                      ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = 1) AND (TypeLevel1 = a.ServiceType)) AS ServiceType_C, " + Environment.NewLine +
                         "                      ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = 2) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS ServiceTypeB_C, " + Environment.NewLine +
                         "                      ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = 3) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS ServiceTypeC_C " + Environment.NewLine +
                         "                 FROM CustomService AS a " + Environment.NewLine +
                         "                WHERE isnull(Driver2, '') <> '' " + Environment.NewLine +
                         "                  AND isnull(IsTrue, '') = 'V' " + Environment.NewLine +
                         vWStr_ServiceType +
                         vWStr_Driver +
                         vWStr_ServiceDate +
                         "                union all " + Environment.NewLine +
                         "               SELECT AthorityDepNo, AthorityDepNote, Driver, DriverName, BuildDate, ServiceDate, LinesNo2 as LinesNo, " + Environment.NewLine +
                         "                      ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = 1) AND (TypeLevel1 = a.ServiceType)) AS ServiceType_C, " + Environment.NewLine +
                         "                      ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = 2) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS ServiceTypeB_C, " + Environment.NewLine +
                         "                      ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = 3) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS ServiceTypeC_C " + Environment.NewLine +
                         "                 FROM CustomService AS a " + Environment.NewLine +
                         "                WHERE isnull(Driver, '') <> '' " + Environment.NewLine +
                         "                  AND isnull(IsTrue, '') = 'V' " + Environment.NewLine +
                         "                  AND isnull(LinesNo2, '') <> '' " + Environment.NewLine +
                         vWStr_ServiceType +
                         vWStr_Driver +
                         vWStr_ServiceDate +
                         "                union all " + Environment.NewLine +
                         "               SELECT AthorityDepNo, AthorityDepNote, Driver2 as Driver, DriverName2 as DriverName, BuildDate, ServiceDate, LinesNo2 as LinesNo, " + Environment.NewLine +
                         "                      ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = 1) AND (TypeLevel1 = a.ServiceType)) AS ServiceType_C, " + Environment.NewLine +
                         "                      ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = 2) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS ServiceTypeB_C, " + Environment.NewLine +
                         "                      ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = 3) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS ServiceTypeC_C " + Environment.NewLine +
                         "                 FROM CustomService AS a " + Environment.NewLine +
                         "                WHERE isnull(Driver2, '') <> '' " + Environment.NewLine +
                         "                  AND isnull(IsTrue, '') = 'V' " + Environment.NewLine +
                         "                  AND isnull(LinesNo2, '') <> '' " + Environment.NewLine +
                         vWStr_ServiceType +
                         vWStr_Driver +
                         vWStr_ServiceDate +
                         ") aa " + Environment.NewLine +
                         " order by aa.AthorityDepNo, aa.Driver, aa.ServiceDate, aa.ServiceType, aa.ServiceTypeB, aa.ServiceTypeC";
            return vResultStr;
        }

        /// <summary>
        /// 預覽報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPreview_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                string vSQLStr = GetSelectStr();
                DataTable dtPrint = new DataTable();
                SqlDataAdapter daPrint = new SqlDataAdapter(vSQLStr, connPrint);
                connPrint.Open();
                daPrint.Fill(dtPrint);
                ReportDataSource rdsPrint = new ReportDataSource("CustomServiceDriverList", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\CustomServiceDriverList.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.Refresh();
                plReport.Visible = true;
                plSearch.Visible = false;

                string vServiceTypeStr = rbServiceType_Search.SelectedItem.Text.Trim();
                string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                string vServiceDateStr = ((eServiceDate_Search_S.Text.Trim() != "") && (eServiceDate_Search_E.Text.Trim() != "")) ? "自 " + eServiceDate_Search_S.Text.Trim() + " 起至 " + eServiceDate_Search_E.Text.Trim() + " 止" :
                                         ((eServiceDate_Search_S.Text.Trim() != "") && (eServiceDate_Search_E.Text.Trim() == "")) ? eServiceDate_Search_S.Text.Trim() :
                                         ((eServiceDate_Search_S.Text.Trim() == "") && (eServiceDate_Search_E.Text.Trim() != "")) ? eServiceDate_Search_E.Text.Trim() : "";
                string vRecordNote = "預覽報表_駕駛客訴案件報表_" + rbGroupBy.SelectedItem.Text + Environment.NewLine +
                                     "CustomServiceDriverList.aspx" + Environment.NewLine +
                                     "客訴類別：" + vServiceTypeStr + Environment.NewLine +
                                     "駕駛員編號：" + vDriverStr + Environment.NewLine +
                                     "客訴日期：" + vServiceDateStr;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                string vSQLStr = GetSelectStr();
                SqlCommand cmdExcel = new SqlCommand(vSQLStr, connExcel);
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

                    string vHeaderText = "";
                    int vLinesNo = 0;
                    string vFileName = "駕駛客訴案件報表";
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "AthorityDepNo") ? "單位編號" :
                                      (drExcel.GetName(i) == "AthorityDepNote") ? "單位" :
                                      (drExcel.GetName(i) == "Driver") ? "員工編號" :
                                      (drExcel.GetName(i) == "DriverName") ? "姓名" :
                                      (drExcel.GetName(i) == "ServiceDate") ? "發生日期" :
                                      (drExcel.GetName(i) == "BuildDate") ? "開單日期" :
                                      (drExcel.GetName(i) == "ServiceType") ? "反映類別代碼" :
                                      (drExcel.GetName(i) == "ServiceType_C") ? "反映類別" :
                                      (drExcel.GetName(i) == "ServiceTypeB") ? "客訴類別代碼" :
                                      (drExcel.GetName(i) == "ServiceTypeB_C") ? "客訴類別" :
                                      (drExcel.GetName(i) == "ServiceTypeC") ? "客訴事項代碼" :
                                      (drExcel.GetName(i) == "ServiceTypeC_C") ? "客訴事項" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i) == "BuildDate") || (drExcel.GetName(i) == "ServiceDate")) && (drExcel[i].ToString().Trim() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(((vBuDate.Year) - 1911).ToString("D3") + "/" + vBuDate.ToString("MM/dd"));
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
                            string vServiceTypeStr = rbServiceType_Search.SelectedItem.Text.Trim();
                            string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vServiceDateStr = ((eServiceDate_Search_S.Text.Trim() != "") && (eServiceDate_Search_E.Text.Trim() != "")) ? "自 " + eServiceDate_Search_S.Text.Trim() + " 到 " + eServiceDate_Search_E.Text.Trim() :
                                                     ((eServiceDate_Search_S.Text.Trim() != "") && (eServiceDate_Search_E.Text.Trim() == "")) ? eServiceDate_Search_S.Text.Trim() :
                                                     ((eServiceDate_Search_S.Text.Trim() == "") && (eServiceDate_Search_E.Text.Trim() != "")) ? eServiceDate_Search_E.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_駕駛客訴案件報表_" + rbGroupBy.SelectedItem.Text + Environment.NewLine +
                                                 "CustomServiceDriverList.aspx" + Environment.NewLine +
                                                 "客訴類別：" + vServiceTypeStr + Environment.NewLine +
                                                 "駕駛員編號：" + vDriverStr + Environment.NewLine +
                                                 "客訴日期：" + vServiceDateStr;
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

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            rvPrint.LocalReport.DataSources.Clear();
            plReport.Visible = false;
            plSearch.Visible = true;
        }
    }
}