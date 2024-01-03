using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
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
    public partial class EducationList : System.Web.UI.Page
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
                        GetDepartmentList();
                        plSearch.Visible = true;
                        plPrint.Visible = false;
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

        private void GetDepartmentList()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            eDepNo_Search.Items.Clear();
            using (SqlConnection connGetList = new SqlConnection(vConnStr))
            {
                string vListStr = "select DepNo, (DepNo + '_' + [Name]) as DepName " + Environment.NewLine +
                                  "  from Department " + Environment.NewLine +
                                  " where DepNo in (select distinct DepNo from Employee where isnull(Leaveday, '') = '') " + Environment.NewLine +
                                  "   and DepNo != '00' " + Environment.NewLine +
                                  " order by DepNo ";
                SqlCommand cmdGetList = new SqlCommand(vListStr, connGetList);
                connGetList.Open();
                SqlDataReader drGetList = cmdGetList.ExecuteReader();
                while (drGetList.Read())
                {
                    ListItem ltTemp = new ListItem(drGetList["DepName"].ToString().Trim(), drGetList["DepNo"].ToString().Trim());
                    eDepNo_Search.Items.Add(ltTemp);
                }
            }
        }

        private string GetSelectStr()
        {
            string vResultStr = "";
            string vWStr_DepNo = "";
            string vWStr_EmpType = "";

            for (int i = 0; i < eDepNo_Search.Items.Count; i++)
            {
                if (eDepNo_Search.Items[i].Selected == true)
                {
                    vWStr_DepNo = (vWStr_DepNo == "") ? "   and e.DepNo in ('" + eDepNo_Search.Items[i].Value.Trim() + "' " : vWStr_DepNo + ", '" + eDepNo_Search.Items[i].Value.Trim() + "' ";
                }
            }
            vWStr_DepNo = (vWStr_DepNo != "") ? vWStr_DepNo + ") " + Environment.NewLine : "";
            for (int i = 0; i < eEmpType_Search.Items.Count; i++)
            {
                if (eEmpType_Search.Items[i].Selected == true)
                {
                    vWStr_EmpType = (vWStr_EmpType == "") ? "   and e.[Type] in ('" + eEmpType_Search.Items[i].Value.Trim() + "' " : vWStr_EmpType + ", '" + eEmpType_Search.Items[i].Value.Trim() + "' ";
                }
            }
            vWStr_EmpType = (vWStr_EmpType != "") ? vWStr_EmpType + ") " + Environment.NewLine : "";

            vResultStr = "select d.[Name] DepName, t.ClassTxt TitleName, e.EmpNo, e.[Name] EmpName " + Environment.NewLine +
                         "  from Employee e left join Department d on d.DepNo = e.DepNo " + Environment.NewLine +
                         "                  left join DBDICB t on t.ClassNo = e.Title and t.FKey = '人事資料檔      EMPLOYEE        TITLE' " + Environment.NewLine +
                         " where isnull(e.Leaveday, '') = '' " + Environment.NewLine +
                         vWStr_DepNo +
                         vWStr_EmpType +
                         " order by e.DepNo, e.Title, e.EmpNo ";
            return vResultStr;
        }

        protected void bbPreview_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelectStr();
            string vReportName = "職業安全衛生教育訓練宣導";
            try
            {
                using (SqlConnection connPrint = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                    connPrint.Open();
                    DataTable dtPrint = new DataTable("EducationListP");
                    daPrint.Fill(dtPrint);
                    ReportDataSource rdsPrint = new ReportDataSource("dsEducationListP", dtPrint);
                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\EducationListP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                    rvPrint.LocalReport.Refresh();
                    plSearch.Visible = false;
                    plPrint.Visible = true;
                }
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
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
            string vSelectStr = GetSelectStr();

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                DataTable dtExcel = new DataTable();
                SqlDataAdapter daExcel = new SqlDataAdapter(vSelectStr, connExcel);
                connExcel.Open();
                daExcel.Fill(dtExcel);
                if (dtExcel.Rows.Count > 0)
                {
                    int vLinesNo = 0;
                    string vColName = "";
                    string vColHeadText = "";
                    string vFileName = "職業安全衛生教育訓練宣導";
                    string vSheetName = "";
                    string vDepName = "";
                    string vTitleName = "";
                    string vEmpNo = "";
                    string vEmpName = "";
                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vDepName = dtExcel.Rows[i]["DepName"].ToString().Trim();
                        vTitleName = dtExcel.Rows[i]["TitleName"].ToString().Trim();
                        vEmpNo = dtExcel.Rows[i]["EmpNo"].ToString().Trim();
                        vEmpName = dtExcel.Rows[i]["EmpName"].ToString().Trim();

                        if (vDepName != vSheetName)
                        {
                            //部門名稱有變，建立一個新的工作表
                            vSheetName = vDepName;
                            //開始準備轉入 EXCEL 的資料
                            wsExcel = (HSSFSheet)wbExcel.CreateSheet(vSheetName);
                            vLinesNo = 0;
                            //寫入標題列
                            wsExcel.CreateRow(vLinesNo);
                            for (int ColIndex = 0; ColIndex < dtExcel.Columns.Count; ColIndex++)
                            {
                                vColName = dtExcel.Columns[ColIndex].ColumnName.Trim();
                                vColHeadText = (vColName == "EmpNo") ? "工號" :
                                               (vColName == "EmpName") ? "姓名" :
                                               (vColName == "DepName") ? "現職單位" :
                                               (vColName == "TitleName") ? "職稱" : "";
                                wsExcel.GetRow(vLinesNo).CreateCell(ColIndex).SetCellValue(vColHeadText);
                                wsExcel.GetRow(vLinesNo).GetCell(ColIndex).CellStyle = csTitle;
                            }
                        }
                        wsExcel = (HSSFSheet)wbExcel.GetSheet(vSheetName);
                        vLinesNo++;
                        //寫入資料
                        wsExcel.CreateRow(vLinesNo);
                        for (int j = 0; j < dtExcel.Columns.Count; j++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(j).SetCellValue(dtExcel.Rows[i][j].ToString().Trim());
                            wsExcel.GetRow(vLinesNo).GetCell(j).CellStyle = csData;
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
                            string vEmpTypeStr = "";
                            string vDepNameStr = "";
                            for (int i = 0; i < eDepNo_Search.Items.Count; i++)
                            {
                                if (eDepNo_Search.Items[i].Selected == true)
                                {
                                    vDepNameStr = (vDepNameStr.Trim() == "") ? eDepNo_Search.Items[i].Text.Trim() : vDepNameStr + "、" + eDepNo_Search.Items[i].Text.Trim();
                                }
                            }
                            for (int i = 0; i < eEmpType_Search.Items.Count; i++)
                            {
                                if (eEmpType_Search.Items[i].Selected == true)
                                {
                                    vEmpTypeStr = (vEmpTypeStr.Trim() == "") ? eEmpType_Search.Items[i].Text.Trim() : vEmpTypeStr + "、" + eEmpType_Search.Items[i].Text.Trim();
                                }
                            }
                            string vRecordNote = "匯出檔案_職業安全衛生教育訓練宣導" + Environment.NewLine +
                                                 "EducationList.aspx" + Environment.NewLine +
                                                 "員工類別：" + vEmpTypeStr + Environment.NewLine +
                                                 "站別：" + vDepNameStr;
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

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCLoseReport_Click(object sender, EventArgs e)
        {
            plSearch.Visible = true;
            plPrint.Visible = false;
        }
    }
}