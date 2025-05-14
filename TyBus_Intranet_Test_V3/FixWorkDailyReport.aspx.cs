using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class FixWorkDailyReport : System.Web.UI.Page
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
        private string vTempStr = "";

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

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        GetListViewItems(eEmpNoList_Search);
                        plPrint.Visible = false;
                        plSearch.Visible = true;
                    }
                    string vFDateURL = "InputDate.aspx?TextboxID=" + eFixInDateS_Search.ClientID;
                    string vFDateScript = "window.open('" + vFDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eFixInDateS_Search.Attributes["onClick"] = vFDateScript; 
                    vFDateURL = "InputDate.aspx?TextboxID=" + eFixInDateE_Search.ClientID;
                    vFDateScript = "window.open('" + vFDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eFixInDateE_Search.Attributes["onClick"] = vFDateScript;
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

        private string GetSelectString()
        {
            string vResultStr = "";
            DateTime vFixDate_S;
            DateTime vFixDate_E;
            string vSelectedEmpNo = "";
            foreach (ListItem liTemp in eEmpNoList_Search.Items)
            {
                if ((vSelectedEmpNo.Trim() == "") && (liTemp.Selected == true))
                {
                    vSelectedEmpNo = "'" + liTemp.Value.Trim() + "'";
                }
                else if ((vSelectedEmpNo.Trim() != "") && (liTemp.Selected == true))
                {
                    vSelectedEmpNo = vSelectedEmpNo.Trim() + ", '" + liTemp.Value.Trim() + "'";
                }
            }
            string vWStr_EmpNo = (vSelectedEmpNo.Trim() != "") ? "   and (a.BuMan in (" + vSelectedEmpNo + ") or c.FixMan in (" + vSelectedEmpNo + ")) " + Environment.NewLine : "";
            string vWStr_FixDate = ((DateTime.TryParse(eFixInDateS_Search.Text.Trim(), out vFixDate_S) == true) && (DateTime.TryParse(eFixInDateE_Search.Text.Trim(), out vFixDate_E) == true)) ?
                                   "   and (a.BuDate between '" + vFixDate_S.ToString("yyyy/MM/dd") + "' and '" + vFixDate_E.ToString("yyyy/MM/dd") + "')" + Environment.NewLine :
                                   ((DateTime.TryParse(eFixInDateS_Search.Text.Trim(), out vFixDate_S) == true) && (DateTime.TryParse(eFixInDateE_Search.Text.Trim(), out vFixDate_E) == false)) ?
                                   "   and a.BuDate = '" + vFixDate_S.ToString("yyyy/MM/dd") + "' " + Environment.NewLine :
                                   ((DateTime.TryParse(eFixInDateS_Search.Text.Trim(), out vFixDate_S) == false) && (DateTime.TryParse(eFixInDateE_Search.Text.Trim(), out vFixDate_E) == true)) ?
                                   "   and a.BuDate = '" + vFixDate_E.ToString("yyyy/MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            vResultStr = "select a.FixDateIn, isnull(a.Car_ID, '') as Car_ID, isnull(a.Maintain1, '') as MainTain1, isnull(d.ClassTxt, '') as MainTain1_C, " + Environment.NewLine +
                         "       isnull(a.Maintain, '') as MainTain, isnull(d2.ClassTxt, '') as MainTain_C, isnull(a.[Service], '') as [Service], isnull(d3.ClassTxt, '') as Service_C, " + Environment.NewLine +
                         "       cast((convert(varchar(10), a.FixDateIn, 111) + ' ' + case when RTrim(LTrim(a.TotTime)) = ':' then '00:00' else a.TotTime END) as datetime) as FixInTime, " + Environment.NewLine +
                         "       cast((convert(varchar(10), a.FixDateOut, 111) + ' ' + case when RTrim(LTrim(a.EndTime)) = ':' then '00:00' else a.EndTime END) as datetime) as FixOutTime, " + Environment.NewLine +
                         "       a.Remark, isnull(a.BuMan, '') as BuMan, isnull(e.[Name], '') as BuMan_C, isnull(c.FixMan, '') as FixMan, isnull(e2.[Name], '') as FixMan_C, " + Environment.NewLine +
                         "       (cast(isnull(c.[Hours], 0.0) as varchar) + ':' + right('00' + cast(isnull(c.[Minute], 0.0) as varchar), 2)) as WorkHours " + Environment.NewLine +
                         "  from FixWorkA a left join FixWorkC c on c.WorkNo = a.WorkNo " + Environment.NewLine +
                         "                  left join DBDICB d on d.ClassNo = a.Maintain1 and d.FKey = '工作單A         FixworkA        MAINTAIN1' " + Environment.NewLine +
                         "                  left join DBDICB d2 on d2.ClassNo = a.Maintain1 and d2.FKey = '工作單A         FixworkA        MAINTAIN' " + Environment.NewLine +
                         "                  left join DBDICB d3 on d3.ClassNo = a.[Service] and d3.FKey = '工作單A         FixworkA        SERVICE' " + Environment.NewLine +
                         "                  left join Employee e on e.EmpNo = a.BuMan " + Environment.NewLine +
                         "                  left join Employee e2 on e2.EmpNo = c.FixMan " + Environment.NewLine +
                         " where isnull(a.WorkNo, '') <> '' " + Environment.NewLine +
                         vWStr_EmpNo +
                         vWStr_FixDate +
                         vWStr_DepNo +
                         " order by a.FixDateIn, a.WorkNo ";
            return vResultStr;
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo = "";
            string vDepName = "";
            vDepNo = eDepNo_Search.Text.Trim();
            vTempStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            vDepName = PF.GetValue(vConnStr, vTempStr, "Name");
            if (vDepName == "")
            {
                vDepName = eDepNo_Search.Text.Trim();
                vTempStr = "select top 1 DepNo from Department where [Name] like '" + vDepName + "%' ";
                vDepNo = PF.GetValue(vConnStr, vTempStr, "DepNo");
            }
            eDepNo_Search.Text = vDepNo.Trim();
            eDepName_Search.Text = vDepName.Trim();
        }

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelStr = GetSelectString();
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);
                if (dtPrint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("dsFixWorkDailyReportP", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\FixWorkDailyReportP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", PF.GetValue(vConnStr, "select [Name] from Custom where Types = 'O' and Code = 'A000'", "Name")));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "修車廠修車日報表"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plSearch.Visible = false;

                    //string vSelectedEmpNo = "";
                    string vRecordNote = "預覽報表_修車廠修車日報表" + Environment.NewLine +
                                         "FixWorkDailyReport.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('查無資料')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void bbExcel_Click(object sender, EventArgs e)
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

            HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

            HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
            csData_Int.DataFormat = format.GetFormat("###,##0");

            HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
            csData_Float.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csData_Float.DataFormat = format.GetFormat("###,##0.00");

            string vHeaderText = "";
            int vLinesNo = 0;
            string vSelStr = GetSelectString();
            string vSheetName = "";
            string vFixDateIn = "";
            DateTime vDate;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //有資料才開始執行
                    while (drExcel.Read())
                    {
                        vDate = DateTime.Parse(drExcel["FixDateIn"].ToString().Trim());
                        vFixDateIn = vDate.Year.ToString("D4") + "年" + vDate.Month.ToString("D2") + "月" + vDate.Day.ToString("D2") + "日";
                        if (vFixDateIn != vSheetName)
                        {
                            vLinesNo = 0;
                            vSheetName = vFixDateIn;
                            wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                            //標題列
                            wsExcel.CreateRow(vLinesNo);
                            for (int i = 0; i < drExcel.FieldCount; i++)
                            {
                                vHeaderText = (drExcel.GetName(i).ToUpper().Trim() == "FIXDATEIN") ? "進廠日期" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "CAR_ID") ? "車號" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "MAINTAIN1_C") ? "小修類別" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "MAINTAIN_C") ? "維修項目" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "SERVICE_C") ? "保養級別" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "FIXINTIME") ? "進廠時間" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "FIXOUTTIME") ? "出廠時間" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "REMARK") ? "工作記要" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "BUMAN_C") ? "開單人" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "FIXMAN_C") ? "承修人" :
                                              (drExcel.GetName(i).ToUpper().Trim() == "WORKHOURS") ? "工時" : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellValue(vHeaderText);
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csTitle;
                            }
                        }
                        wsExcel = (HSSFSheet)wbExcel.GetSheet(vSheetName);
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            if ((drExcel.GetName(i).ToUpper().Trim() == "FIXDATEIN")||
                                (drExcel.GetName(i).ToUpper().Trim() == "FIXINTIME") ||
                                (drExcel.GetName(i).ToUpper().Trim() == "FIXOUTTIME"))
                            {
                                vDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vDate.Year.ToString("D4") + "/" + vDate.ToString("MM/dd hh:mm:ss"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
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
                            string vCheckTearStr = "自 " + eFixInDateS_Search.Text.Trim() + " 起至 " + eFixInDateE_Search.Text.Trim();
                            string vCheckDepNo = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "不分站別";
                            string vSelectedEmpNo = "";
                            foreach (ListItem liTemp in eEmpNoList_Search.Items)
                            {
                                if ((vSelectedEmpNo.Trim() == "") && (liTemp.Selected == true))
                                {
                                    vSelectedEmpNo = liTemp.Value.Trim();
                                }
                                else if ((vSelectedEmpNo.Trim() != "") && (liTemp.Selected == true))
                                {
                                    vSelectedEmpNo = vSelectedEmpNo.Trim() + ", " + liTemp.Value.Trim();
                                }
                            }
                            string vCheckEmpNo = (vSelectedEmpNo != "") ? vSelectedEmpNo : "所有承修人";
                            string vRecordNote = "匯出檔案_修車廠修車日報表" + Environment.NewLine +
                                                 "FixWorkDailyReport.aspx" + Environment.NewLine +
                                                 "查詢年月：" + vCheckTearStr + Environment.NewLine +
                                                 "車輛站別：" + vCheckDepNo + Environment.NewLine +
                                                 "承修人員：" + vCheckEmpNo;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            string vFileName = "修車廠修車日報表";

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

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        private void GetListViewItems(ListBox lbTemp)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connListItem = new SqlConnection(vConnStr))
            {
                lbTemp.Items.Clear();
                vTempStr = "select EmpNo, (EmpNo + '_' + [Name]) as EmpName " + Environment.NewLine +
                           "  from Employee " + Environment.NewLine +
                           " where DepNo = '30' " + Environment.NewLine +
                           "   and isnull(Leaveday, '') = '' " + Environment.NewLine +
                           "   and [Type] = '10' " + Environment.NewLine +
                           " order by EmpNo ";
                SqlCommand cmdListItem = new SqlCommand(vTempStr, connListItem);
                connListItem.Open();
                SqlDataReader drListItem = cmdListItem.ExecuteReader();
                while (drListItem.Read())
                {
                    ListItem liTemp = new ListItem();
                    liTemp.Value = drListItem["EmpNo"].ToString().Trim();
                    liTemp.Text = drListItem["EmpName"].ToString().Trim();
                    lbTemp.Items.Add(liTemp);
                }
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plSearch.Visible = true;
            //rvPrint.Dispose();
        }
    }
}