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
    public partial class OfficialDocumentHistory : System.Web.UI.Page
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
                    if (!IsPostBack)
                    {
                        eDocDate_Start_Search.Text = PF.GetMonthFirstDay(vToday, "C");
                        eDocDate_End_Search.Text = PF.GetMonthLastDay(vToday, "C");
                    }

                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    string vBuildDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_Start_Search.ClientID;
                    string vBuildDateScript = "window.open('" + vBuildDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eDocDate_Start_Search.Attributes["onClick"] = vBuildDateScript;

                    vBuildDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_End_Search.ClientID;
                    vBuildDateScript = "window.open('" + vBuildDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eDocDate_End_Search.Attributes["onClick"] = vBuildDateScript;

                    vBuildDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eModifyDate_Start_Search.ClientID;
                    vBuildDateScript = "window.open('" + vBuildDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eModifyDate_Start_Search.Attributes["onClick"] = vBuildDateScript;

                    vBuildDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eModifyDate_End_Search.ClientID;
                    vBuildDateScript = "window.open('" + vBuildDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eModifyDate_End_Search.Attributes["onClick"] = vBuildDateScript;

                    DocDataBind();
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

        private string GetSelStr()
        {
            string vWStr_DocDep = ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() != "")) ?
                                  "   and a.DocDep between '" + eDocDep_Start_Search.Text.Trim() + "' and '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() == "")) ?
                                  "   and a.DocDep = '" + eDocDep_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eDocDep_Start_Search.Text.Trim() == "") && (eDocDep_End_Search.Text.Trim() != "")) ?
                                  "   and a.DocDep = '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vWStr_DocDate = ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() != "")) ?
                                   "   and a.DocDate between '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                   ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() == "")) ?
                                   "   and a.DocDate = '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                   ((eDocDate_Start_Search.Text.Trim() == "") && (eDocDate_End_Search.Text.Trim() != "")) ?
                                   "   and a.DocDate = '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";

            string vWStr_ModifyDate = ((eModifyDate_Start_Search.Text.Trim() != "") && (eModifyDate_End_Search.Text.Trim() != "")) ?
                                      "   and a.ModifyDate between '" + PF.TransDateString(eModifyDate_Start_Search.Text.Trim(), "B") + " 00:00:00' and '" + PF.TransDateString(eModifyDate_End_Search.Text.Trim(), "B") + " 23:59:59' " + Environment.NewLine :
                                      ((eModifyDate_Start_Search.Text.Trim() != "") && (eModifyDate_End_Search.Text.Trim() == "")) ?
                                      "   and a.ModifyDate between '" + PF.TransDateString(eModifyDate_Start_Search.Text.Trim(), "B") + " 00:00:00' and '" + PF.TransDateString(eModifyDate_Start_Search.Text.Trim(), "B") + " 23:59:59' " + Environment.NewLine :
                                      ((eModifyDate_Start_Search.Text.Trim() == "") && (eModifyDate_End_Search.Text.Trim() != "")) ?
                                      "   and a.ModifyDate between '" + PF.TransDateString(eModifyDate_End_Search.Text.Trim(), "B") + " 00:00:00' and '" + PF.TransDateString(eModifyDate_End_Search.Text.Trim(), "B") + " 23:59:59' " + Environment.NewLine : "";

            string vWStr_DocType = (eDocType_Search.Text.Trim() != "") ? "   and DocType = '" + eDocType_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vSelectStr = "select DocIndex, Items, DocDate, DocDep, " + Environment.NewLine +
                                "       (select [Name] from Department where Department.DepNo = a.DocDep) DocDepName, " + Environment.NewLine +
                                "       DocFirstWord, (select DocFirstCWord from DOCFirstWord where DOCFirstWord.FWNo = a.DocFirstWord) DocFirstWord_C, " + Environment.NewLine +
                                "       DocNo, DocSourceUnit, " + Environment.NewLine +
                                "       DocType, (select ClassTxt from DBDICB where FKey = '公文收發登記    OfficialDocumentDocType' and ClassNo = a.DocType) DocType_C, " + Environment.NewLine +
                                "       DocTitle, Undertaker, (select [Name] from Employee where EmpNo = a.Undertaker) UndertakerName, " + Environment.NewLine +
                                "       OutsideDocFirstWord, OutsideDocNo, AttacheMent, Implementation, " + Environment.NewLine +
                                "       BuildMan, (select [Name] from Employee where EmpNo = a.BuildMan) BuildManName, " + Environment.NewLine +
                                "       BuildDate, Remark, ModifyDate, ModifyMode, " + Environment.NewLine +
                                "       ModifyMan, (select [Name] from Employee where EmpNo = a.ModifyMan) ModifyManName, " + Environment.NewLine +
                                "       DocYears, IsHide " + Environment.NewLine +
                                "  from OfficialDocumentHistory a " + Environment.NewLine +
                                " where isnull(DocIndex, '') <> '' " + Environment.NewLine +
                                vWStr_DocDep + vWStr_DocType + vWStr_DocDate + vWStr_ModifyDate +
                                " order by DocDate DESC, DocIndex, Items";
            return vSelectStr;
        }

        private void DocDataBind()
        {
            string vSelectStr = GetSelStr();

            sdsOfficialDocumentHistoryList.SelectCommand = "";
            sdsOfficialDocumentHistoryList.SelectCommand = vSelectStr;
            gridOfficialDocumentHistoryList.DataBind();
        }

        protected void eDocDep_Start_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Search = eDocDep_Start_Search.Text.Trim();
            string vDepName_Search = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "select [Name] from Department where DepNo = '" + vDepNo_Search + "' ";
            vDepName_Search = PF.GetValue(vConnStr, vSelectStr, "Name");
            if (vDepName_Search == "")
            {
                vDepName_Search = vDepNo_Search;
                vSelectStr = "select DepNo from Department where [Name] = '" + vDepName_Search + "' ";
                vDepNo_Search = PF.GetValue(vConnStr, vSelectStr, "DepNo");
            }
            eDocDep_Start_Search.Text = vDepNo_Search;
            eDocDepName_Start_Search.Text = vDepName_Search;
        }

        protected void eDocDep_End_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Search = eDocDep_End_Search.Text.Trim();
            string vDepName_Search = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "select [Name] from Department where DepNo = '" + vDepNo_Search + "' ";
            vDepName_Search = PF.GetValue(vConnStr, vSelectStr, "Name");
            if (vDepName_Search == "")
            {
                vDepName_Search = vDepNo_Search;
                vSelectStr = "select DepNo from Department where [Name] = '" + vDepName_Search + "' ";
                vDepNo_Search = PF.GetValue(vConnStr, vSelectStr, "DepNo");
            }
            eDocDep_End_Search.Text = vDepNo_Search;
            eDocDepName_End_Search.Text = vDepName_Search;
        }

        protected void ddlDocType_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDocType_Search.Text = ddlDocType_Search.SelectedValue;
        }

        protected void eUndertaker_Start_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo_Search = eUndertaker_Start_Search.Text.Trim();
            string vEmpName_Search = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_Search + "' and LeaveDay is null '";
            vEmpName_Search = PF.GetValue(vConnStr, vSelectStr, "Name");
            if (vEmpName_Search == "")
            {
                vEmpName_Search = vEmpNo_Search;
                vSelectStr = "select EmpNo from Employee where [Name] = '" + vEmpName_Search + "' and LeaveDay is null";
                vEmpNo_Search = PF.GetValue(vConnStr, vSelectStr, "EmpNo");
            }
            eUndertaker_Start_Search.Text = vEmpNo_Search;
            eUndertakerName_Start_Search.Text = vEmpName_Search;
        }

        protected void eUndertaker_End_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo_Search = eUndertaker_End_Search.Text.Trim();
            string vEmpName_Search = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_Search + "' and LeaveDay is null '";
            vEmpName_Search = PF.GetValue(vConnStr, vSelectStr, "Name");
            if (vEmpName_Search == "")
            {
                vEmpName_Search = vEmpNo_Search;
                vSelectStr = "select EmpNo from Employee where [Name] = '" + vEmpName_Search + "' and LeaveDay is null";
                vEmpNo_Search = PF.GetValue(vConnStr, vSelectStr, "EmpNo");
            }
            eUndertaker_End_Search.Text = vEmpNo_Search;
            eUndertakerName_End_Search.Text = vEmpName_Search;
        }

        protected void ddlModifyMode_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eModifyMode_Search.Text = ddlModifyMode_Search.SelectedValue;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            DocDataBind();
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                string vSelStr = GetSelStr();
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    DateTime vTempDate;
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

                    string vHeaderText = ""; ;
                    string vFileName = "公文異動歷史記錄";
                    int vLinesNo = 0;
                    int vColIndex = 0;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "DOCDATE") ? "收發日期" :
                                      (drExcel.GetName(i).ToUpper() == "DOCDEP") ? "收發單位編號" :
                                      (drExcel.GetName(i).ToUpper() == "DOCDEPNAME") ? "收發單位" :
                                      (drExcel.GetName(i).ToUpper() == "DOCSOURCEUNIT") ? "來文機關" :
                                      (drExcel.GetName(i).ToUpper() == "DOCTYPE") ? "文別代號" :
                                      (drExcel.GetName(i).ToUpper() == "DOCTYPE_C") ? "文別" :
                                      (drExcel.GetName(i).ToUpper() == "DOCTITLE") ? "事由" :
                                      (drExcel.GetName(i).ToUpper() == "UNDERTAKER") ? "承辦人工號" :
                                      (drExcel.GetName(i).ToUpper() == "UNDERTAKERNAME") ? "承辦人" :
                                      (drExcel.GetName(i).ToUpper() == "REMARK") ? "說明" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYDATE") ? "異動日期" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYMODE") ? "異動類別" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYMAN") ? "異動人工號" :
                                      (drExcel.GetName(i).ToUpper() == "MODIFYMANNAME") ? "異動人" : "";
                        if (vHeaderText != "")
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(vColIndex);
                            wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellType(CellType.String);
                            wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellValue(vHeaderText);
                            wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csTitle;
                            vColIndex++;
                        }
                    }

                    vLinesNo++;
                    while (drExcel.Read())
                    {
                        wsExcel.CreateRow(vLinesNo);
                        vColIndex = 0;
                        for (int i = 0; i < drExcel.FieldCount; i++)
                        {
                            if ((drExcel.GetName(i).ToUpper() == "DOCDATE") || (drExcel.GetName(i).ToUpper() == "MODIFYDATE"))
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(vColIndex);
                                wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellType(CellType.String);
                                if (drExcel[i].ToString() != "")
                                {
                                    vTempDate = DateTime.Parse(drExcel[i].ToString());
                                    wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellValue(vTempDate.Year.ToString("D4") + "/" + vTempDate.ToString("MM/dd"));
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellValue("");
                                }
                                wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csData;
                                vColIndex++;
                            }
                            else if ((drExcel.GetName(i).ToUpper() == "DOCDEP") || (drExcel.GetName(i).ToUpper() == "DOCDEPNAME") ||
                                     (drExcel.GetName(i).ToUpper() == "DOCSOURCEUNIT") || (drExcel.GetName(i).ToUpper() == "DOCTYPE") ||
                                     (drExcel.GetName(i).ToUpper() == "DOCTYPE_C") || (drExcel.GetName(i).ToUpper() == "DOCTITLE") ||
                                     (drExcel.GetName(i).ToUpper() == "UNDERTAKER") || (drExcel.GetName(i).ToUpper() == "UNDERTAKERNAME") ||
                                     (drExcel.GetName(i).ToUpper() == "REMARK") || (drExcel.GetName(i).ToUpper() == "MODIFYDATE") ||
                                     (drExcel.GetName(i).ToUpper() == "MODIFYMODE") || (drExcel.GetName(i).ToUpper() == "MODIFYMAN") ||
                                     (drExcel.GetName(i).ToUpper() == "MODIFYMANNAME"))
                            {
                                wsExcel.GetRow(vLinesNo).CreateCell(vColIndex);
                                wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(vColIndex).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(vColIndex).CellStyle = csData;
                                vColIndex++;
                            }
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
                        //throw;
                    }
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}