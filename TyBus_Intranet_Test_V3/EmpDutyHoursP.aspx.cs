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

namespace TyBus_Intranet_Test_V3
{
    public partial class EmpDutyHoursP : System.Web.UI.Page
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
                        plPrint.Visible = false;
                        eCalYear_S.Text = (DateTime.Today.Year - 1911).ToString();
                        eCalMonth_S.Text = "01";
                        eCalYear_E.Text = eCalYear_S.Text.Trim();
                        eCalMonth_E.Text = (DateTime.Today.Month > 1) ? (DateTime.Today.Month - 1).ToString("D2") : "01";
                    }
                    //eCalYear_S.Enabled = ((rbReportType.SelectedValue == "DutyHours") || (rbReportType.SelectedValue == "ESCHours"));
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
        /// 取回資料查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelStr()
        {
            string vSelStr = "";
            //string vCalYear = (eCalYear_S.Text.Trim() != "") ? eCalYear_S.Text.Trim() : (DateTime.Today.Year - 1911).ToString();
            string vCalYear_S = (eCalYear_S.Text.Trim() != "") ? eCalYear_S.Text.Trim() : "";
            string vCalYear_E = (eCalYear_E.Text.Trim() != "") ? eCalYear_E.Text.Trim() : "";
            string vCalMonth_S = (eCalMonth_S.Text.Trim() != "") ? Int32.Parse(eCalMonth_S.Text.Trim()).ToString("D2") : "";
            string vCalMonth_E = (eCalMonth_E.Text.Trim() != "") ? Int32.Parse(eCalMonth_E.Text.Trim()).ToString("D2") : "";
            string vCalYM_S = (vCalYear_S.Trim() != "") ? vCalYear_S.Trim() + vCalMonth_S.Trim() : "";
            string vCalYM_E = (vCalYear_E.Trim() != "") ? vCalYear_E.Trim() + vCalMonth_E.Trim() : "";
            if (Int32.Parse(vCalYM_S) > Int32.Parse(vCalYM_E))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('您輸入的起迄日期錯誤')");
                Response.Write("</" + "Script>");
                eCalYear_S.Focus();
            }
            else if ((vCalYM_S == "") && (vCalYM_E == ""))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('起迄日期不可以全都空白')");
                Response.Write("</" + "Script>");
                eCalYear_S.Focus();
            }
            else
            {
                string vWStr_DepNo = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? "           and e.DepNo between '" + eDepNo_S.Text.Trim() + "' and '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? "           and e.DepNo = '" + eDepNo_S.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? "           and e.depNo = '" + eDepNo_E.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_EmpNo = (eEmpNo_S.Text.Trim() != "") ? "           and e.EmpNo = '" + eEmpNo_S.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_DutyDate = ((vCalYM_S != "") && (vCalYM_E != "")) ? "           and left(e.DutyIndex, " + vCalYM_S.Length.ToString() + ") >= '" + vCalYM_S + "' and left(e.DutyIndex, " + vCalYM_E.Length.ToString() + ") <= '" + vCalYM_E + "' " :
                                        ((vCalYM_S != "") && (vCalYM_E == "")) ? "           and left(e.DutyIndex, " + vCalYM_S.Length.ToString() + ") = '" + vCalYM_S + "' " :
                                        ((vCalYM_S == "") && (vCalYM_E != "")) ? "           and left(e.DutyIndex, " + vCalYM_E.Length.ToString() + ") = '" + vCalYM_E + "' " : "";
                switch (rbReportType.SelectedValue)
                {
                    case "DutyHours":
                        vSelStr = "select a.DutyYear, a.DepName, a.EmpNo, a.EmpName, a.Title, " + Environment.NewLine +
                                  "       (select ClassTxt from DBDICB where FKey = '員工資料        EMPLOYEE        TITLE' and ClassNo = a.Title) Title_C, " + Environment.NewLine +
                                  "       a.DutyDateStart, a.StartTime, a.DutyDateEnd, a.EndTime, a.DutyHours, a.Remark " + Environment.NewLine +
                                  "  from (" + Environment.NewLine +
                                  "        select left(DutyIndex, 3) DutyYear, DepNo, DepName, EmpNo, EmpName, " + Environment.NewLine +
                                  "               (select Title from Employee where EmpNo = e.EmpNo) Title, " + Environment.NewLine +
                                  "               DutyDateStart, StartTime, DutyDateEnd, EndTime, DutyHours, Remark " + Environment.NewLine +
                                  "          from EmpDutyHours e " + Environment.NewLine +
                                  //"         where e.DutyType = 'T01' and left(e.DutyIndex, 3) = '" + vCalYear + "' " + Environment.NewLine +
                                  "         where e.DutyType = 'T01' and (e.IsAllowed = 1) " + Environment.NewLine +
                                  vWStr_DutyDate + vWStr_DepNo + vWStr_EmpNo +
                                  ") a" + Environment.NewLine +
                                  " order by a.DepNo, a.Title, a.EmpNo, a.DutyYear";
                        break;
                    case "ESCHours":
                        vSelStr = "select a.DutyYear, a.DepName, a.EmpNo, a.EmpName, a.Title, " + Environment.NewLine +
                                  "       (select ClassTxt from DBDICB where FKey = '員工資料        EMPLOYEE        TITLE' and ClassNo = a.Title) Title_C, " + Environment.NewLine +
                                  "       a.DutyDateStart, a.StartTime, a.DutyDateEnd, a.EndTime, a.ESCHours, a.Remark " + Environment.NewLine +
                                  "  from (" + Environment.NewLine +
                                  "        select left(DutyIndex, 3) DutyYear, DepNo, DepName, EmpNo, EmpName, " + Environment.NewLine +
                                  "               (select Title from Employee where EmpNo = e.EmpNo) Title, " + Environment.NewLine +
                                  "               DutyDateStart, StartTime, DutyDateEnd, EndTime, ESCHours, Remark " + Environment.NewLine +
                                  "          from EmpDutyHours e " + Environment.NewLine +
                                  //"         where e.DutyType = 'T02' and left(e.DutyIndex, 3) = '" + vCalYear + "' " + Environment.NewLine +
                                  "         where e.DutyType = 'T02' and (e.IsAllowed = 1) " + Environment.NewLine +
                                  vWStr_DutyDate + vWStr_DepNo + vWStr_EmpNo +
                                  ") a" + Environment.NewLine +
                                  " order by a.DepNo, a.Title, a.EmpNo, a.DutyYear";
                        break;
                    case "Overdue":
                        vSelStr = "select a.DutyYear, a.DepName, a.EmpNo, a.EmpName, a.Title, " + Environment.NewLine +
                                  "       (select ClassTxt from DBDICB where FKey = '員工資料        EMPLOYEE        TITLE' and ClassNo = a.Title) Title_C, " + Environment.NewLine +
                                  "       a.TotalUseableHours, a.Remark " + Environment.NewLine +
                                  "  from (" + Environment.NewLine +
                                  "        select left(DutyIndex, 3) DutyYear, DepNo, DepName, EmpNo, EmpName, " + Environment.NewLine +
                                  "               (select Title from Employee where EmpNo = e.EmpNo) Title, " + Environment.NewLine +
                                  "               sum(UsableHours) TotalUseableHours, cast(null as varchar) Remark " + Environment.NewLine +
                                  "          from EmpDutyHours e " + Environment.NewLine +
                                  "         where e.DutyType = 'T01' and (e.IsAllowed = 1) and left(e.DutyIndex, 3) < year(Getdate()) - 1912 " + Environment.NewLine +
                                  vWStr_DepNo + vWStr_EmpNo +
                                  "         group by left(e.DutyIndex, 3), e.DepName, e.EmpNo, e.EmpName, e.DepNo " + Environment.NewLine +
                                  ") a" + Environment.NewLine +
                                  " order by a.DepNo, a.Title, a.EmpNo, a.DutyYear";
                        break;
                    case "UseableHours":
                        vSelStr = "select a.DutyYear, a.DepName, a.EmpNo, a.EmpName, a.Title, " + Environment.NewLine +
                                  "       (select ClassTxt from DBDICB where FKey = '員工資料        EMPLOYEE        TITLE' and ClassNo = a.Title) Title_C, " + Environment.NewLine +
                                  "       a.TotalUseableHours, a.Remark " + Environment.NewLine +
                                  "  from (" + Environment.NewLine +
                                  "        select left(DutyIndex, 3) DutyYear, DepNo, DepName, EmpNo, EmpName, " + Environment.NewLine +
                                  "               (select Title from Employee where EmpNo = e.EmpNo) Title, " + Environment.NewLine +
                                  "               sum(UsableHours) TotalUseableHours, cast(null as varchar) Remark " + Environment.NewLine +
                                  "          from EmpDutyHours e " + Environment.NewLine +
                                  "         where e.DutyType = 'T01' and (e.IsAllowed = 1) and left(e.DutyIndex, 3) >= year(Getdate()) - 1912 " + Environment.NewLine +
                                  vWStr_DepNo + vWStr_EmpNo +
                                  "         group by left(e.DutyIndex, 3), e.DepName, e.EmpNo, EmpName, e.DepNo " + Environment.NewLine +
                                  ") a" + Environment.NewLine +
                                  " order by a.DepNo, a.Title, a.EmpNo, a.DutyYear";
                        break;
                }
            }
            return vSelStr;
        }

        protected void eDepNo_S_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDepNo_S.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_S.Text = vDepNo_Temp;
            eDepName_S.Text = vDepName_Temp;
        }

        protected void eDepNo_E_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDepNo_E.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_E.Text = vDepNo_Temp;
            eDepName_E.Text = vDepName_Temp;
        }

        protected void eEmpNo_S_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo = eEmpNo_S.Text.Trim();
            string vEmpName = "";
            string vSQLStr = "select [Name] from EMployee where EmpNo = '" + vEmpNo + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName + "' ";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eEmpNo_S.Text = vEmpNo;
            eEmpName_S.Text = vEmpName;
        }

        /// <summary>
        /// 預覽列印
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPreview_Click(object sender, EventArgs e)
        {
            string vSelectStr = GetSelStr();
            string vReportName = "EmpDutyHoursP_" + rbReportType.SelectedValue;
            string vCalYear_S = (eCalYear_S.Text.Trim() != "") ? eCalYear_S.Text.Trim() : (DateTime.Today.Year - 1911).ToString();
            string vCalYear_E = (eCalYear_E.Text.Trim() != "") ? eCalYear_E.Text.Trim() : (DateTime.Today.Year - 1911).ToString();
            string vCalMonth_S = (eCalMonth_S.Text.Trim() != "") ? Int32.Parse(eCalMonth_S.Text.Trim()).ToString("D2") : "01";
            string vCalMonth_E = (eCalMonth_E.Text.Trim() != "") ? Int32.Parse(eCalMonth_E.Text.Trim()).ToString("D2") : (DateTime.Today.Month - 1).ToString("D2");
            string vCalYear = ((vCalYear_S == vCalYear_E) && (vCalMonth_S == vCalMonth_E)) ? vCalYear_S + " 年 " + vCalMonth_S + " 月 " :
                              ((vCalYear_S == vCalYear_E) && (vCalMonth_S != vCalMonth_E)) ? vCalYear_S + " 年 " + vCalMonth_S + " 月至 " + vCalMonth_E + " 月" :
                              (vCalYear_S != vCalYear_E) ? vCalYear_S + " 年 " + vCalMonth_S + " 月至 " + vCalYear_E + " 年 " + vCalMonth_E + " 月" : "";
            vCalYear += rbReportType.SelectedItem.Text.Trim();

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable();
                daPrint.Fill(dtPrint);
                if (dtPrint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("EmpDutyHoursP", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\" + vReportName + ".rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CalYear", vCalYear));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", PF.GetValue(vConnStr, "select [Name] from Custom where Types = 'O' and Code = 'A000'", "Name")));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plSearch.Visible = false;

                    string vReportTypeStr = rbReportType.SelectedItem.Text.Trim();
                    string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_S.Text.Trim() + " 站至 " + eDepNo_E.Text.Trim() + " 站" :
                                       ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() + " 站" :
                                       ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() + " 站" : "全部";
                    string vRecordNote = "預覽報表_員工加班補休報表" + Environment.NewLine +
                                         "EmpDutyHoursP.aspx" + Environment.NewLine +
                                         "報表類別：" + vReportTypeStr + Environment.NewLine +
                                         "年度別：" + vCalYear + Environment.NewLine +
                                         "站別：" + vDepNoStr;
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

            string vSelectStr = GetSelStr();
            string vReportName = "EmpDutyHoursP_" + rbReportType.SelectedValue;
            string vCalYear_S = (eCalYear_S.Text.Trim() != "") ? eCalYear_S.Text.Trim() : (DateTime.Today.Year - 1911).ToString();
            string vCalYear_E = (eCalYear_E.Text.Trim() != "") ? eCalYear_E.Text.Trim() : (DateTime.Today.Year - 1911).ToString();
            string vCalMonth_S = (eCalMonth_S.Text.Trim() != "") ? Int32.Parse(eCalMonth_S.Text.Trim()).ToString("D2") : "01";
            string vCalMonth_E = (eCalMonth_E.Text.Trim() != "") ? Int32.Parse(eCalMonth_E.Text.Trim()).ToString("D2") : (DateTime.Today.Month - 1).ToString("D2");
            string vCalYear = ((vCalYear_S == vCalYear_E) && (vCalMonth_S == vCalMonth_E)) ? vCalYear_S + "年" + vCalMonth_S + "月" :
                              ((vCalYear_S == vCalYear_E) && (vCalMonth_S != vCalMonth_E)) ? vCalYear_S + "年" + vCalMonth_S + "月至" + vCalMonth_E + "月" :
                              (vCalYear_S != vCalYear_E) ? vCalYear_S + "年" + vCalMonth_S + "月至" + vCalYear_E + "年" + vCalMonth_E + "月" : "";
            vCalYear = (rbReportType.SelectedIndex <= 1) ? vCalYear + rbReportType.SelectedItem.Text.Trim() : rbReportType.SelectedItem.Text.Trim();

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
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vCalYear);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "DutyYear") ? "年度" :
                                      (drExcel.GetName(i) == "DepName") ? "單位別" :
                                      (drExcel.GetName(i) == "EmpNo") ? "人員工號" :
                                      (drExcel.GetName(i) == "EmpName") ? "姓名" :
                                      (drExcel.GetName(i) == "Title") ? "職稱代號" :
                                      (drExcel.GetName(i) == "Title_C") ? "職稱" :
                                      (drExcel.GetName(i) == "DutyDateStart") ? "開始日期" :
                                      (drExcel.GetName(i) == "StartTime") ? "開始時間" :
                                      (drExcel.GetName(i) == "DutyDateEnd") ? "結束日期" :
                                      (drExcel.GetName(i) == "EndTime") ? "結束時間" :
                                      (drExcel.GetName(i) == "DutyHours") ? "時數" :
                                      (drExcel.GetName(i) == "TotalUseableHours") ? "時數" :
                                      (drExcel.GetName(i) == "ESCHours") ? "時數" :
                                      (drExcel.GetName(i) == "Remark") ? "備註" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i) == "NCheckTerm") || (drExcel.GetName(i) == "NextEDate")) && (drExcel[i].ToString() != ""))
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
                            string vReportTypeStr = rbReportType.SelectedItem.Text.Trim();
                            string vDepNoStr = ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_S.Text.Trim() + " 站至 " + eDepNo_E.Text.Trim() + " 站" :
                                               ((eDepNo_S.Text.Trim() != "") && (eDepNo_E.Text.Trim() == "")) ? eDepNo_S.Text.Trim() + " 站" :
                                               ((eDepNo_S.Text.Trim() == "") && (eDepNo_E.Text.Trim() != "")) ? eDepNo_E.Text.Trim() + " 站" : "全部";
                            string vRecordNote = "預覽報表_員工加班補休報表" + Environment.NewLine +
                                                 "EmpDutyHoursP.aspx" + Environment.NewLine +
                                                 "報表類別：" + vReportTypeStr + Environment.NewLine +
                                                 "年度別：" + vCalYear + Environment.NewLine +
                                                 "站別：" + vDepNoStr;
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            string vFileName = "車輛檢驗到期查核";
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vCalYear + ".xls"));
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
            plPrint.Visible = false;
            plSearch.Visible = true;
        }

        protected void rbReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (rbReportType.SelectedIndex)
            {
                case 0:
                case 1:
                    eCalYear_S.Enabled = true;
                    lbSplit_1.Visible = true;
                    eCalYear_E.Visible = true;
                    lbSplit_2.Visible = true;
                    eCalMonth_S.Visible = true;
                    lbSplit_3.Visible = true;
                    eCalMonth_E.Visible = true;
                    lbSplit_4.Visible = true;
                    break;
                case 2:
                case 3:
                    eCalYear_S.Enabled = false;
                    lbSplit_1.Visible = true;
                    eCalYear_E.Visible = false;
                    lbSplit_2.Visible = false;
                    eCalMonth_S.Visible = false;
                    lbSplit_3.Visible = false;
                    eCalMonth_E.Visible = false;
                    lbSplit_4.Visible = false;
                    break;
            }
        }
    }
}