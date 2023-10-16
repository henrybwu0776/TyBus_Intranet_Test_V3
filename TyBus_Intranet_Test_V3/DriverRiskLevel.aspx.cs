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
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverRiskLevel : System.Web.UI.Page
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
                    if (!IsPostBack)
                    {
                        plPrint.Visible = false;
                        plShowData.Visible = true;
                        plSearch.Visible = true;
                        plShowData_Detail.Visible = false;
                        plShowData_Detail2.Visible = false;
                        cbGetFullYear_Search.Checked = true;
                        eRiskYearS_Search.Text = (DateTime.Today.AddMonths(-1).Year - 1911).ToString();
                        eRiskYearE_Search.Text = (DateTime.Today.AddMonths(-1).Year - 1911).ToString();
                        eRiskMonthS_Search.SelectedIndex = 0;
                        eRiskMonthE_Search.SelectedIndex = 11;
                        eRiskMonthS_Search.Enabled = (cbGetFullYear_Search.Checked == false);
                        eRiskYearE_Search.Enabled = (cbGetFullYear_Search.Checked == false);
                        eRiskMonthE_Search.Enabled = (cbGetFullYear_Search.Checked == false);
                    }
                    else
                    {
                        OpenData();
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

        private string GetSearchStr()
        {
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   AND d.IDCardNo in (select IDCardNo from Employee where DepNo = '" + eDepNo_Search.Text.Trim() + "') " + Environment.NewLine : "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and d.IDCardNo = (select IDCardNo from Employee where EmpNo = '" + eEmpNo_Search.Text.Trim() + "') " + Environment.NewLine : "";
            string vWStr_CalYM = "";
            if (cbGetFullYear_Search.Checked)
            {
                vWStr_CalYM = (eRiskYearS_Search.Text.Trim() != "") ? "   and d.CalYear = '" + eRiskYearS_Search.Text.Trim() + "' " + Environment.NewLine : "";
            }
            else
            {
                vWStr_CalYM = ((eRiskYearS_Search.Text.Trim() != "") && (eRiskYearE_Search.Text.Trim() != "")) ?
                              "   and (d.CalYM between '" + eRiskYearS_Search.Text.Trim() + Int32.Parse(eRiskMonthS_Search.SelectedValue).ToString("D2") + "' and '" + eRiskYearE_Search.Text.Trim() + Int32.Parse(eRiskMonthE_Search.SelectedValue).ToString("D2") + "')" + Environment.NewLine :
                              ((eRiskYearS_Search.Text.Trim() != "") && (eRiskYearE_Search.Text.Trim() == "")) ?
                              "   and (d.CalYM between '" + eRiskYearS_Search.Text.Trim() + Int32.Parse(eRiskMonthS_Search.SelectedValue).ToString("D2") + "' and '" + eRiskYearS_Search.Text.Trim() + Int32.Parse(eRiskMonthE_Search.SelectedValue).ToString("D2") + "')" + Environment.NewLine :
                              ((eRiskYearS_Search.Text.Trim() == "") && (eRiskYearE_Search.Text.Trim() != "")) ?
                              "   and (d.CalYM between '" + eRiskYearE_Search.Text.Trim() + Int32.Parse(eRiskMonthS_Search.SelectedValue).ToString("D2") + "' and '" + eRiskYearE_Search.Text.Trim() + Int32.Parse(eRiskMonthE_Search.SelectedValue).ToString("D2") + "')" + Environment.NewLine : "";
            }
            string vResultStr = "select distinct IDCardNo, cast('' as varchar) EmpNo, cast('' as varchar) EmpName, cast('' as varchar) DepNo, cast('' as varchar) DepName " + Environment.NewLine +
                                "  from DriverRiskLevel d " + Environment.NewLine +
                                " where isnull(IDCardNo, '') <> '' " + Environment.NewLine +
                                vWStr_CalYM + vWStr_DepNo + vWStr_EmpNo +
                                " order by DepNo, IDCardNo";
            return vResultStr;
        }

        private void OpenData()
        {
            string vSearechStr = GetSearchStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DataTable dtShowData = new DataTable();
            using (SqlConnection connShowData = new SqlConnection(vConnStr))
            {
                string vIDCardNo = "";
                string[] vStrStore;
                string vTempStr = "";
                string vResultStr = "";

                SqlDataAdapter daShowData = new SqlDataAdapter(vSearechStr, connShowData);
                connShowData.Open();
                daShowData.Fill(dtShowData);
                for (int i = 0; i < dtShowData.Rows.Count; i++)
                {
                    vIDCardNo = dtShowData.Rows[i]["IDCardNo"].ToString().Trim();
                    vTempStr = "select top 1 EmpNo + ',' + [Name] + ',' + DepNo + ',' + (select [Name] from Department where DepNo = e.DepNo) ResultStr " + Environment.NewLine +
                               "  from Employee e " + Environment.NewLine +
                               " where IDCardNo = '" + vIDCardNo + "' " + Environment.NewLine +
                               " order by AssumeDay DESC";
                    vResultStr = PF.GetValue(vConnStr, vTempStr, "ResultStr");
                    if (vResultStr != "")
                    {
                        vStrStore = vResultStr.Split(',');
                        dtShowData.Rows[i]["EmpNo"] = vStrStore[0].Trim();
                        dtShowData.Rows[i]["EmpName"] = vStrStore[1].Trim();
                        dtShowData.Rows[i]["DepNo"] = vStrStore[2].Trim();
                        dtShowData.Rows[i]["DepName"] = vStrStore[3].Trim();
                    }
                }
                gridDriverRiskLevelList.DataSourceID = "";
                gridDriverRiskLevelList.DataSource = dtShowData;
                gridDriverRiskLevelList.DataBind();
            }
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Search = eDepNo_Search.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vTempStr = "select [Name] from Department where DepNo = '" + vDepNo_Search + "' ";
            string vDepName_Search = PF.GetValue(vConnStr, vTempStr, "Name");
            if (vDepName_Search.Trim() == "")
            {
                vDepName_Search = vDepNo_Search.Trim();
                vTempStr = "select top 1 DepNo from Department where [Name] = '" + vDepName_Search.Trim() + "' order by DepNo DESC ";
                vDepNo_Search = PF.GetValue(vConnStr, vTempStr, "DepNo");
            }
            eDepNo_Search.Text = vDepNo_Search.Trim();
            eDepName_Search.Text = vDepName_Search.Trim();
        }

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo_Search = eEmpNo_Search.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vTempStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_Search + "' ";
            string vEmpName_Search = PF.GetValue(vConnStr, vTempStr, "Name");
            if (vEmpName_Search.Trim() == "")
            {
                vEmpName_Search = vEmpNo_Search.Trim();
                vTempStr = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Search.Trim() + "' order by EmpNo DESC ";
                vEmpNo_Search = PF.GetValue(vConnStr, vTempStr, "EmpNo");
            }
            eEmpNo_Search.Text = vEmpNo_Search.Trim();
            eEmpName_Search.Text = vEmpName_Search.Trim();
        }


        protected void bbGetNameList_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vTempStr = "";
            string vCalYM = (DateTime.Today.AddMonths(-1).Month < 9) ? (DateTime.Today.AddMonths(-1).Year - 1).ToString() : DateTime.Today.AddMonths(-1).Year.ToString();
            string vLeaveDay = vCalYM + "/09/01";

            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                vTempStr = "select EmpNo, IDCardNo " + Environment.NewLine +
                           "  from Employee " + Environment.NewLine +
                           " where (isnull(LeaveDay, '') = '' or LeaveDay > '" + vLeaveDay + "') " + Environment.NewLine +
                           "   and isnull(DepNo, '00') <> '00' " + Environment.NewLine +
                           "   and left(DepNo, 1) not in ('Z', 'S') " + Environment.NewLine +
                           " order by DepNo ";
                SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    //查詢有結果才進行
                    string vTempIDCardNo = "";
                    string vTempEmpNo = "";
                    string vTargetEmpNo = "";
                    string vSQLStr = "";
                    string vIndexNo = "";

                    while (drTemp.Read())
                    {
                        vTempIDCardNo = drTemp["IDCardNo"].ToString().Trim();
                        vTempEmpNo = drTemp["EmpNo"].ToString().Trim();
                        vIndexNo = vCalYM.Trim() + vTempIDCardNo.Trim();
                        vTempStr = "select EmpNo from DriverRiskLevel where IndexNo = '" + vIndexNo + "' ";
                        vTargetEmpNo = PF.GetValue(vConnStr, vTempStr, "EmpNo");
                        if (vTargetEmpNo == "")
                        {
                            vSQLStr = "insert into DriverRiskLevel (IndexNo, CalYM, IDCardNo, EmpNo)" + Environment.NewLine +
                                      "values ('" + vIndexNo + "', '" + vCalYM + "', '" + vTempIDCardNo + "', '" + vTempEmpNo + "')";
                            PF.ExecSQL(vConnStr, vSQLStr);
                        }
                        else if (vTargetEmpNo != vTempEmpNo)
                        {
                            vSQLStr = "Update DriverRiskLevel set EmpNo = '" + vTempEmpNo + "' where IndexNo = '" + vIndexNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr);
                        }
                    }
                }
                drTemp.Close();
                cmdTemp.Cancel();
            }
            /* 計算每月工時和上班天數
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                vTempStr = "select EmpNo from DriverRiskLevel where CalYM = '" + vCalYM + "' order by IndexNo";
                SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    //查詢有結果才執行
                    string vRiskDateFieldName = RiskMonthIndex[DateTime.Today.AddMonths(-1).Month - 1].Trim();
                    string vRiskDate_S = PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "B");
                    string vRiskDate_E = PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "B");
                    string vWorkHoursFieldName = vRiskDateFieldName.Replace("RiskDate", "WorkHours");
                    string vWorkDaysFieldName = vRiskDateFieldName.Replace("RiskDate", "WorkDays");
                    string vReturnStr = "";
                    string[] vStrStore;
                    int vWorkMins = 0;
                    int vWorkHours = 0;
                    string vWorkHoursStr = "";

                    while (drTemp.Read())
                    {
                        vTempStr = "select cast(count(AssignNo) as varchar) + ',' + cast(sum(WorkHR * 60) + sum(WorkMin) as varchar) as ReturnStr " + Environment.NewLine +
                                   "  from RunSheetA " + Environment.NewLine +
                                   " where BuDate between '" + vRiskDate_S + "' and '" + vRiskDate_E + "' " + Environment.NewLine +
                                   "   and Driver = '" + drTemp["EmpNo"].ToString().Trim() + "' " + Environment.NewLine +
                                   " group by Driver ";
                        vReturnStr = PF.GetValue(vConnStr, vTempStr, "ReturnStr");
                        if ((vReturnStr.Trim() != "") && (vReturnStr.IndexOf(':') >= 0))
                        {
                            vStrStore = vReturnStr.Split(',');
                            vWorkMins = Int32.Parse(vStrStore[1]);
                            vWorkHours = vWorkMins / 60;
                            vWorkHoursStr = vWorkHours.ToString() + ":" + (vWorkMins - (vWorkHours * 60)).ToString();
                            vTempStr = "update DriverRiskLevel " + Environment.NewLine +
                                       "   set " + vWorkDaysFieldName + " = " + vStrStore[0] + ", " + vWorkHoursFieldName + " = " + vWorkHoursStr + Environment.NewLine +
                                       " where CalYM = '" + vCalYM + "' " + Environment.NewLine +
                                       "   and EmpNo = '" + drTemp["EmpNo"].ToString().Trim() + "' ";
                            PF.ExecSQL(vConnStr, vTempStr);
                        }
                    }
                }
                drTemp.Close();
                cmdTemp.Cancel();
            }*/
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 取得匯出 EXCEL 用的 SQL 語法
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr_Excel()
        {
            string vCalYM_S = "";
            string vCalYM_E = "";
            string vWStr_DepNo = "";
            string vWStr_EmpNo = "";
            string vResultStr = "";
            int vMonthCount = 0;
            int vStartYear = 0;
            int vCurrentYear = 0;
            int vStartMonth = 0;
            int vCurrentMonth = 0;
            if (cbGetFullYear_Search.Checked)
            {
                vMonthCount = 12;
                vStartYear = Int32.Parse(eRiskYearS_Search.Text.Trim());
                vStartMonth = 1;
                vCalYM_S = eRiskYearS_Search.Text.Trim() + "01";
                vCalYM_E = eRiskYearS_Search.Text.Trim() + "12";
            }
            else
            {
                vMonthCount = (Int32.Parse(eRiskYearE_Search.Text.Trim()) - Int32.Parse(eRiskYearS_Search.Text.Trim())) * 12 +
                              (Int32.Parse(eRiskMonthE_Search.SelectedValue) - Int32.Parse(eRiskMonthS_Search.SelectedValue)) + 1;
                vStartYear = Int32.Parse(eRiskYearS_Search.Text.Trim());
                vStartMonth = Int32.Parse(eRiskMonthS_Search.SelectedValue);
                vCalYM_S = eRiskYearS_Search.Text.Trim() + Int32.Parse(eRiskMonthS_Search.SelectedValue).ToString("D2");
                vCalYM_E = eRiskYearE_Search.Text.Trim() + Int32.Parse(eRiskMonthE_Search.SelectedValue).ToString("D2");

            }
            vResultStr = "select (select top 1 EmpNo from Employee where IDCardNo = z.IDCardNo order by AssumeDay DESC) EmpNo, " + Environment.NewLine +
                         "       (select top 1 [Name] from Employee where IDCardNo = z.IDCardNo order by AssumeDay DESC) EmpName, " + Environment.NewLine +
                         "       (select top 1 DepNo from Employee where IDCardNo = z.IDCardNo order by AssumeDay DESC) DepNo, " + Environment.NewLine +
                         "       (select [Name] from Department where DepNo = (select top 1 DepNo from Employee where IDCardNo = z.IDCardNo order by AssumeDay DESC)) DepName, " + Environment.NewLine;
            for (int i = 0; i < vMonthCount; i++)
            {
                vCurrentMonth = (vStartMonth + i) - ((vStartMonth + i - 1) / 12) * 12;
                vCurrentYear = (vStartMonth + i > 12) ? vStartYear + ((vStartMonth + i) / 12) : vStartYear;
                vResultStr += "       cast(null as datetime) as [" + vCurrentYear.ToString() + "/" + vCurrentMonth.ToString("D2") + "_RiskDate], " +
                              " cast(null as varchar) as [" + vCurrentYear.ToString() + "/" + vCurrentMonth.ToString("D2") + "_DoctorLevel], " +
                              " cast(null as varchar) as [" + vCurrentYear.ToString() + "/" + vCurrentMonth.ToString("D2") + "_CompanyLevel], " +
                              " cast(null as varchar) as [" + vCurrentYear.ToString() + "/" + vCurrentMonth.ToString("D2") + "_WorkHours], " +
                              " cast(null as int) as [" + vCurrentYear.ToString() + "/" + vCurrentMonth.ToString("D2") + "_WorkDays], " + Environment.NewLine;
            }
            vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "           and (select top 1 DepNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "           and (select top 1 EmpNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) = '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            vResultStr += "       z.IDCardNo " + Environment.NewLine +
                          "  from (" + Environment.NewLine +
                          "        select distinct IDCardNo from DriverRiskLevel d " + Environment.NewLine +
                          "         where d.CalYM >= '" + vCalYM_S + "' and d.CalYM <= '" + vCalYM_E + "' " + Environment.NewLine +
                          vWStr_DepNo + vWStr_EmpNo +
                          "       ) z " + Environment.NewLine +
                          " order by (select top 1 DepNo from Employee where IDCardNo = z.IDCardNo order by AssumeDay DESC) ";
            return vResultStr;
        }

        protected void bbExportExcel_Click(object sender, EventArgs e)
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

            string vSQLStr = GetSelectStr_Excel();
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                DataTable dtExcel = new DataTable();
                SqlDataAdapter daExcel = new SqlDataAdapter(vSQLStr, connExcel);
                connExcel.Open();
                daExcel.Fill(dtExcel);
                if (dtExcel.Rows.Count > 0)
                {
                    string vIDCardNo_Data = "";
                    string vEmpNo_Data = "";
                    string vCalYM_S = "";
                    string vCalYM_E = "";
                    string vCalDate_S = "";
                    string vCalDate_E = "";
                    string vCurrentCalYear = "";
                    string vCurrentCalYM = "";
                    string vColName_RiskDate = "";
                    string vColName_DoctorLevel = "";
                    string vColName_CompanyLevel = "";
                    string vColName_WorkHours = "";
                    string vColName_WorkDays = "";
                    string vLastDoctorLevel = "";

                    //根據是否勾選 "產生全年度報表" 來決定資料取回起迄年月
                    if (cbGetFullYear_Search.Checked)
                    {
                        vCalYM_S = eRiskYearS_Search.Text.Trim() + "01";
                        vCalYM_E = eRiskYearS_Search.Text.Trim() + "12";
                        vCalDate_S = (Int32.Parse(eRiskYearS_Search.Text.Trim()) < 1911) ?
                                     (Int32.Parse(eRiskYearS_Search.Text.Trim()) + 1911).ToString() + "/01/01" :
                                     eRiskYearS_Search.Text.Trim() + "01/01";
                        vCalDate_E = (Int32.Parse(eRiskYearS_Search.Text.Trim()) < 1911) ?
                                     (Int32.Parse(eRiskYearS_Search.Text.Trim()) + 1911).ToString() + "/12/31" :
                                     eRiskYearS_Search.Text.Trim() + "12/31";
                    }
                    else
                    {
                        vCalYM_S = eRiskYearS_Search.Text.Trim() + Int32.Parse(eRiskMonthS_Search.Text.Trim()).ToString("D2");
                        vCalYM_E = eRiskYearE_Search.Text.Trim() + Int32.Parse(eRiskMonthE_Search.Text.Trim()).ToString("D2");
                        vCalDate_S = (Int32.Parse(eRiskYearS_Search.Text.Trim()) < 1911) ?
                                     (Int32.Parse(eRiskYearS_Search.Text.Trim()) + 1911).ToString() + "/" + eRiskMonthS_Search.Text.Trim() + "/01" :
                                     eRiskYearS_Search.Text.Trim() + "/" + eRiskMonthS_Search.Text.Trim() + "/01";
                        vCalDate_E = PF.GetMonthLastDay(DateTime.Parse(vCalDate_S), "B");
                    }

                    if (cbIsIgnoreLevel2_Search.Checked)
                    {
                        //有勾選 "評級2不列入報表"
                        for (int i = dtExcel.Rows.Count - 1; i >= 0; i--)
                        {
                            vIDCardNo_Data = dtExcel.Rows[i]["IDCardNo"].ToString().Trim();
                            vSQLStr = "select top 1 DoctorLevel " + Environment.NewLine +
                                      "  from DriverRiskLevel " + Environment.NewLine +
                                      " where IDCardNo = '" + vIDCardNo_Data + "' " + Environment.NewLine +
                                      "   and CalYM between '" + vCalYM_S + "' and '" + vCalYM_E + "' " + Environment.NewLine +
                                      " order by CalYM DESC";
                            vLastDoctorLevel = PF.GetValue(vConnStr, vSQLStr, "DoctorLevel");
                            if (vLastDoctorLevel == "2")
                            {
                                dtExcel.Rows[i].Delete();
                            }
                        }
                        dtExcel.AcceptChanges();
                    }

                    for (int i = 0; i < dtExcel.Rows.Count; i++)
                    {
                        vIDCardNo_Data = dtExcel.Rows[i]["IDCardNo"].ToString().Trim();
                        vEmpNo_Data = dtExcel.Rows[i]["EmpNo"].ToString().Trim();
                        //vSQLStr = "select CalYear, CalYM, RiskDate, DoctorLevel, CompanyLevel, WorkHours, WorkDays " + Environment.NewLine +
                        //          "  from DriverRiskLevel " + Environment.NewLine +
                        //          " where IDCardNo = '" + vIDCardNo_Data + "' " + Environment.NewLine +
                        //          "   and CalYM >= '" + vCalYM_S + "' and CalYM <= '" + vCalYM_E + "' " + Environment.NewLine +
                        //          " order by CalYM ";
                        //==========================================================================================================================
                        //vSQLStr = "select CalYear, CalYM, RiskDate, DoctorLevel, CompanyLevel, " + Environment.NewLine +
                        //          "       (select round(cast(sum(WorkHR * 60) + sum(WorkMin) as float) / 60.0, 2) from RunSheetA where (cast(Year(BuDate) - 1911 as varchar) + right('00' + cast(Month(BuDate) as varchar), 2)) = dr.CalYM and Driver = '" + vEmpNo_Data + "') WorkHours, " + Environment.NewLine +
                        //          "       (select count(AssignNo) from RunSheetA where (cast(Year(BuDate) - 1911 as varchar) + right('00' + cast(Month(BuDate) as varchar), 2)) = dr.CalYM and Driver = '" + vEmpNo_Data + "') WorkDays " + Environment.NewLine +
                        //          "  from DriverRiskLevel dr " + Environment.NewLine +
                        //          " where IDCardNo = '" + vIDCardNo_Data + "' " + Environment.NewLine +
                        //          "   and CalYM >= '" + vCalYM_S + "' and CalYM <= '" + vCalYM_E + "' " + Environment.NewLine +
                        //          " order by CalYM ";
                        //2022.07.13 換個寫法，不然全年度報表慢到靠北
                        vSQLStr = "select CalYear, CalYM, RiskDate, DoctorLevel, CompanyLevel, " + Environment.NewLine +
                                  "       (round(cast(sum(ra.WorkHR * 60) + sum(WorkMin) as float) / 60.0, 2)) WorkHours, count(AssignNo) WorkDays  " + Environment.NewLine +
                                  //"  from DriverRiskLevel dr left join RunSheetA ra " + Environment.NewLine +
                                  //"       on (cast(Year(ra.BuDate) - 1911 as varchar) + right('00' + cast(Month(ra.BuDate) as varchar), 2)) = dr.CalYM and ra.Driver = '" + vEmpNo_Data + "' " + Environment.NewLine +
                                  //2023.07.25 再換一個寫法，同時在二個資料表各別建立索引，加快查詢速度
                                  "  from DriverRiskLevel dr left join " + Environment.NewLine +
                                  "       (select(cast(Year(BuDate) - 1911 as varchar) + right('00' + cast(Month(BuDate) as varchar), 2)) SheetYM, *from RunSheetA) ra " + Environment.NewLine +
                                  "       on ra.SheetYM = dr.CalYM and ra.Driver = '" + vEmpNo_Data + "'" + Environment.NewLine +
                                  " where IDCardNo = '" + vIDCardNo_Data + "' " + Environment.NewLine +
                                  "   and CalYM >= '" + vCalYM_S + "' and CalYM <= '" + vCalYM_E + "' " + Environment.NewLine +
                                  " group by CalYear, CalYM, RiskDate, DoctorLevel, CompanyLevel " + Environment.NewLine +
                                  " order by CalYM ";
                        using (SqlConnection connRiskData = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdRiskData = new SqlCommand(vSQLStr, connRiskData);
                            connRiskData.Open();
                            SqlDataReader drRiskData = cmdRiskData.ExecuteReader();
                            if (drRiskData.HasRows)
                            {
                                while (drRiskData.Read())
                                {
                                    vCurrentCalYear = drRiskData["CalYear"].ToString().Trim();
                                    vCurrentCalYM = drRiskData["CalYM"].ToString().Trim();
                                    vColName_RiskDate = vCurrentCalYM.Replace(vCurrentCalYear, vCurrentCalYear + "/") + "_RiskDate";
                                    vColName_DoctorLevel = vCurrentCalYM.Replace(vCurrentCalYear, vCurrentCalYear + "/") + "_DoctorLevel";
                                    vColName_CompanyLevel = vCurrentCalYM.Replace(vCurrentCalYear, vCurrentCalYear + "/") + "_CompanyLevel";
                                    vColName_WorkHours = vCurrentCalYM.Replace(vCurrentCalYear, vCurrentCalYear + "/") + "_WorkHours";
                                    vColName_WorkDays = vCurrentCalYM.Replace(vCurrentCalYear, vCurrentCalYear + "/") + "_WorkDays";
                                    dtExcel.Rows[i][vColName_RiskDate] = drRiskData["RiskDate"];
                                    dtExcel.Rows[i][vColName_DoctorLevel] = drRiskData["DoctorLevel"];
                                    dtExcel.Rows[i][vColName_CompanyLevel] = drRiskData["CompanyLevel"];
                                    dtExcel.Rows[i][vColName_WorkHours] = drRiskData["WorkHours"];
                                    dtExcel.Rows[i][vColName_WorkDays] = drRiskData["WorkDays"];
                                }
                            }
                            drRiskData.Close();
                            cmdRiskData.Cancel();
                        }
                    }
                    int vLinesNo = 0;
                    string vColName = "";
                    string vColHeadText = "";
                    string vTempStr = "";
                    DateTime vRiskDate;
                    string vFileName = (cbGetFullYear_Search.Checked) ?
                                       eRiskYearS_Search.Text.Trim() + "年度職醫風險評估清冊" :
                                       eRiskYearS_Search.Text.Trim() + "年" + Int32.Parse(eRiskMonthS_Search.SelectedValue).ToString("D2") + "月至" +
                                       eRiskYearE_Search.Text.Trim() + "年" + Int32.Parse(eRiskMonthE_Search.Text.Trim()).ToString("D2") + "月職醫風險評估清冊";
                    //開始準備轉入 EXCEL 的資料
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int ColIndex = 0; ColIndex < dtExcel.Columns.Count; ColIndex++)
                    {
                        vColName = dtExcel.Columns[ColIndex].ColumnName.Trim();
                        vColHeadText = (vColName.IndexOf("RiskDate") > 0) ? vColName.Replace("RiskDate", "評鑑日期") :
                                       (vColName.IndexOf("DoctorLevel") > 0) ? vColName.Replace("DoctorLevel", "職醫評等") :
                                       (vColName.IndexOf("CompanyLevel") > 0) ? vColName.Replace("CompanyLevel", "內部評等") :
                                       (vColName.IndexOf("WorkHours") > 0) ? vColName.Replace("WorkHours", "工時") :
                                       (vColName.IndexOf("WorkDays") > 0) ? vColName.Replace("WorkDays", "天數") :
                                       (vColName == "EmpNo") ? "工號" :
                                       (vColName == "EmpName") ? "姓名" :
                                       (vColName == "DepNo") ? "" :
                                       (vColName == "DepName") ? "現職單位" :
                                       (vColName == "IDCardNo") ? "身分證字號" : "";
                        vColHeadText = vColHeadText.Replace("/", "年");
                        vColHeadText = vColHeadText.Replace("_", "月");
                        wsExcel.GetRow(vLinesNo).CreateCell(ColIndex).SetCellValue(vColHeadText);
                        wsExcel.GetRow(vLinesNo).GetCell(ColIndex).CellStyle = csTitle;
                    }
                    //寫入資料
                    for (int DataIndex = 0; DataIndex < dtExcel.Rows.Count; DataIndex++)
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int FieldCount = 0; FieldCount < dtExcel.Columns.Count; FieldCount++)
                        {
                            vColName = dtExcel.Columns[FieldCount].ColumnName.Trim();
                            vTempStr = dtExcel.Rows[DataIndex][FieldCount].ToString().Trim();
                            if (vTempStr != "")
                            {
                                if (vColName.IndexOf("RiskDate") > 0)
                                {
                                    vRiskDate = DateTime.Parse(vTempStr);
                                    wsExcel.GetRow(vLinesNo).CreateCell(FieldCount).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(FieldCount).SetCellValue((vRiskDate.Year - 1911).ToString("D3") + "/" + vRiskDate.ToString("MM/dd"));
                                    wsExcel.GetRow(vLinesNo).GetCell(FieldCount).CellStyle = csData;

                                }
                                else if (vColName.IndexOf("WorkDays") > 0)
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
                            string vDepNoStr = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "全部";
                            string vEmpNoStr = (eEmpNo_Search.Text.Trim() != "") ? eEmpNo_Search.Text.Trim() : "全部";
                            string vRiskYMStr = (cbGetFullYear_Search.Checked) ? "取得 " + eRiskYearS_Search.Text.Trim() + " 全年度資料" :
                                                "自 " + eRiskYearS_Search.Text.Trim() + " 年 " + eRiskMonthS_Search.Text.Trim() + "起至 " + eRiskYearE_Search.Text.Trim() + " 年 " + eRiskMonthE_Search.Text.Trim() + "止 ";
                            string vRecordNote = "匯出檔案_員工風險評估等級" + Environment.NewLine +
                                                 "DriverRiskLevel.aspx" + Environment.NewLine +
                                                 "評估時期：" + vRiskYMStr + Environment.NewLine +
                                                 "員工工號：" + vEmpNoStr + Environment.NewLine +
                                                 "站別：" + vDepNoStr;
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

        protected void bbImportExcel_Click(object sender, EventArgs e)
        {
            string vUploadPath = Request.PhysicalApplicationPath + @"App_Data\";
            string vUploadFileName = vUploadPath + fuExcel.FileName;
            string vExtName = "";
            string vTempStr = "";

            string vIndexNo_Temp = "";
            string vCalYear_Excel = "";
            string vCalYM_Excel = "";
            string vEmpNo_Excel = "";
            string vIDCardNo_Temp = "";
            string vIDCardNo_ERP = "";
            string vRiskDateStr_Excel = "";
            DateTime vRiskDate_Excel;
            string vRiskDateStr_S = "";
            string vRiskDateStr_E = "";
            string vRiskDateStr_Temp = "";
            string vCompanyRiskLevel = "";
            string vDoctorRiskLevel = "";
            string vReturnStr = "";
            string[] vStrStore;
            int vWorkMins = 0;
            int vWorkHours = 0;
            string vWorkDaysStr = "";
            string vWorkHoursStr = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if (fuExcel.FileName != "")
            {
                vExtName = Path.GetExtension(fuExcel.FileName);
                fuExcel.SaveAs(vUploadFileName);
                switch (vExtName)
                {
                    case ".xls":
                        //Excel 97-2003
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0);
                        for (int i = sheetExcel_H.FirstRowNum; i <= sheetExcel_H.LastRowNum; i++)
                        {
                            HSSFRow vRowTemp_H = (HSSFRow)sheetExcel_H.GetRow(i);
                            if (vRowTemp_H != null)
                            {
                                vEmpNo_Excel = vRowTemp_H.Cells[1].ToString().Trim();
                                if ((vEmpNo_Excel != "") && (vEmpNo_Excel != "員編"))
                                {
                                    vTempStr = "select IDCardNo from Employee where EmpNo = '" + vEmpNo_Excel + "' ";
                                    vIDCardNo_ERP = PF.GetValue(vConnStr, vTempStr, "IDCardNo");
                                    vTempStr = vRowTemp_H.Cells[5].ToString().Trim();
                                    vRiskDateStr_Excel = PF.TransNoneSymbolDateStrToDate(vTempStr, "C");
                                    vRiskDate_Excel = DateTime.Parse(vRiskDateStr_Excel);
                                    vRiskDateStr_S = PF.GetMonthFirstDay(vRiskDate_Excel, "B");
                                    vRiskDateStr_E = PF.GetMonthLastDay(vRiskDate_Excel, "B");
                                    vRiskDateStr_Temp = "'" + vRiskDate_Excel.Year.ToString() + "/" + vRiskDate_Excel.ToString("MM/dd") + "'";
                                    vCalYear_Excel = (vRiskDate_Excel.Year - 1911).ToString();
                                    vCalYM_Excel = vCalYear_Excel + vRiskDate_Excel.Month.ToString("D2");
                                    vCompanyRiskLevel = (vRowTemp_H.Cells[7].ToString().Trim() != "") ? "'" + vRowTemp_H.Cells[7].ToString().Trim() + "'" : "NULL";
                                    //vDoctorRiskLevel = (vRowTemp_H.Cells[8].ToString().Trim() != "") ? "'" + vRowTemp_H.Cells[8].ToString().Trim() + "'" : "NULL";
                                    if (vRowTemp_H.Cells[8].ToString().Trim() != "")
                                    {
                                        vDoctorRiskLevel = vRowTemp_H.Cells[8].ToString().Trim();
                                    }
                                    else
                                    {
                                        vTempStr = "select top 1 DoctorLevel " + Environment.NewLine +
                                                   "  from DriverRiskLevel " + Environment.NewLine +
                                                   " where CalYM < '" + vCalYM_Excel + "' " + Environment.NewLine +
                                                   "   and IDCardNo = '" + vIDCardNo_ERP + "' " + Environment.NewLine +
                                                   "   and isnull(DoctorLevel, '') <> '' " + Environment.NewLine +
                                                   " order by CalYM DESC ";
                                        vDoctorRiskLevel = PF.GetValue(vConnStr, vTempStr, "DoctorLevel");
                                        if (vDoctorRiskLevel.Trim() != "")
                                        {
                                            vRiskDateStr_Temp = "NULL";
                                        }
                                    }
                                    vDoctorRiskLevel = (vDoctorRiskLevel != "") ? "'" + vDoctorRiskLevel + "'" : "NULL";
                                    vIndexNo_Temp = vCalYM_Excel.Trim() + vIDCardNo_ERP.Trim();
                                    vTempStr = "select cast(count(AssignNo) as varchar) + ',' + cast(sum(WorkHR * 60) + sum(WorkMin) as varchar) as ReturnStr " + Environment.NewLine +
                                               "  from RunSheetA " + Environment.NewLine +
                                               " where BuDate between '" + vRiskDateStr_S + "' and '" + vRiskDateStr_E + "' " + Environment.NewLine +
                                               "   and Driver = '" + vEmpNo_Excel + "' " + Environment.NewLine +
                                               " group by Driver ";
                                    vReturnStr = PF.GetValue(vConnStr, vTempStr, "ReturnStr");
                                    if ((vReturnStr.Trim() != "") && (vReturnStr.IndexOf(',') >= 0))
                                    {
                                        vStrStore = vReturnStr.Split(',');
                                        vWorkDaysStr = (vStrStore[0] != "") ? vStrStore[0] : "NULL";
                                        if (vStrStore[1] != "")
                                        {
                                            vWorkMins = Int32.Parse(vStrStore[1]);
                                            vWorkHours = vWorkMins / 60;
                                            vWorkHoursStr = "'" + vWorkHours.ToString() + ":" + (vWorkMins - (vWorkHours * 60)).ToString("D2") + "'";
                                        }
                                        else
                                        {
                                            vWorkHoursStr = "NULL";
                                        }
                                    }
                                    else
                                    {
                                        vWorkHoursStr = "NULL";
                                        vWorkDaysStr = "NULL";
                                    }
                                    vTempStr = "select IDCardNo from DriverRiskLevel where IndexNo = '" + vIndexNo_Temp + "' ";
                                    vIDCardNo_Temp = PF.GetValue(vConnStr, vTempStr, "IDCardNo");
                                    if (vIDCardNo_Temp == "")
                                    {
                                        vTempStr = "insert into DriverRiskLevel " + Environment.NewLine +
                                                   "       (IndexNo, CalYear, CalYM, IDCardNo, RiskDate, CompanyLevel, DoctorLevel, WorkDays, WorkHours, BuMan, BuDate)" + Environment.NewLine +
                                                   "values ('" + vIndexNo_Temp + "', '" + vCalYear_Excel + "', '" + vCalYM_Excel + "', " + Environment.NewLine +
                                                   "        '" + vIDCardNo_ERP + "', " + vRiskDateStr_Temp + ", " + Environment.NewLine +
                                                   "        " + vCompanyRiskLevel + ", " + vDoctorRiskLevel + ", " + Environment.NewLine +
                                                   "        " + vWorkDaysStr + ", " + vWorkHoursStr + ", " + Environment.NewLine +
                                                   "        '" + vLoginID + "', GetDate())";
                                    }
                                    else
                                    {
                                        vTempStr = "update DriverRiskLevel set RiskDate = " + vRiskDateStr_Temp + ", " + Environment.NewLine +
                                                   "                           CompanyLevel = " + vCompanyRiskLevel + ", " + Environment.NewLine +
                                                   "                           DoctorLevel = " + vDoctorRiskLevel + ", " + Environment.NewLine +
                                                   "                           WorkDays = " + vWorkDaysStr + ", " + Environment.NewLine +
                                                   "                           WorkHours = " + vWorkHoursStr + ", " + Environment.NewLine +
                                                   "                           ModifyMan = '" + vLoginID + "', " + Environment.NewLine +
                                                   "                           ModifyDate = GetDate() " + Environment.NewLine +
                                                   " where IndexNo = '" + vIndexNo_Temp + "' ";
                                    }
                                    if (vTempStr != "")
                                    {
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                }
                            }
                        }
                        break;
                    case ".xlsx":
                        //Excel 2010--
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0);
                        for (int i = sheetExcel_X.FirstRowNum; i <= sheetExcel_X.LastRowNum; i++)
                        {
                            XSSFRow vRowTemp_X = (XSSFRow)sheetExcel_X.GetRow(i);
                            if (vRowTemp_X != null)
                            {
                                vEmpNo_Excel = vRowTemp_X.Cells[1].ToString().Trim();
                                if ((vEmpNo_Excel != "") && (vEmpNo_Excel != "員編"))
                                {
                                    vTempStr = "select IDCardNo from Employee where EmpNo = '" + vEmpNo_Excel + "' ";
                                    vIDCardNo_ERP = PF.GetValue(vConnStr, vTempStr, "IDCardNo");
                                    vTempStr = vRowTemp_X.Cells[5].ToString().Trim();
                                    vRiskDateStr_Excel = PF.TransNoneSymbolDateStrToDate(vTempStr, "C");
                                    vRiskDate_Excel = DateTime.Parse(vRiskDateStr_Excel);
                                    vRiskDateStr_S = PF.GetMonthFirstDay(vRiskDate_Excel, "B");
                                    vRiskDateStr_E = PF.GetMonthLastDay(vRiskDate_Excel, "B");
                                    //vRiskDateStr_Temp = vRiskDate_Excel.Year.ToString() + "/" + vRiskDate_Excel.ToString("MM/dd");
                                    vRiskDateStr_Temp = "'" + vRiskDate_Excel.Year.ToString() + "/" + vRiskDate_Excel.ToString("MM/dd") + "'";
                                    vCalYear_Excel = (vRiskDate_Excel.Year - 1911).ToString();
                                    vCalYM_Excel = vCalYear_Excel + vRiskDate_Excel.Month.ToString("D2");
                                    vCompanyRiskLevel = (vRowTemp_X.Cells[7].ToString().Trim() != "") ? "'" + vRowTemp_X.Cells[7].ToString().Trim() + "'" : "NULL";
                                    //vDoctorRiskLevel = (vRowTemp_X.Cells[8].ToString().Trim() != "") ? "'" + vRowTemp_X.Cells[8].ToString().Trim() + "'" : "NULL";
                                    if (vRowTemp_X.Cells[8].ToString().Trim() != "")
                                    {
                                        vDoctorRiskLevel = vRowTemp_X.Cells[8].ToString().Trim();
                                    }
                                    else
                                    {
                                        vTempStr = "select top 1 DoctorLevel " + Environment.NewLine +
                                                   "  from DriverRiskLevel " + Environment.NewLine +
                                                   " where CalYM < '" + vCalYM_Excel + "' " + Environment.NewLine +
                                                   "   and IDCardNo = '" + vIDCardNo_ERP + "' " + Environment.NewLine +
                                                   "   and isnull(DoctorLevel, '') <> '' " + Environment.NewLine +
                                                   " order by CalYM DESC ";
                                        vDoctorRiskLevel = PF.GetValue(vConnStr, vTempStr, "DoctorLevel");
                                        if (vDoctorRiskLevel.Trim() != "")
                                        {
                                            vRiskDateStr_Temp = "NULL";
                                        }
                                    }
                                    vDoctorRiskLevel = (vDoctorRiskLevel != "") ? "'" + vDoctorRiskLevel + "'" : "NULL";
                                    vIndexNo_Temp = vCalYM_Excel.Trim() + vIDCardNo_ERP.Trim();
                                    vTempStr = "select cast(count(AssignNo) as varchar) + ',' + cast(sum(WorkHR * 60) + sum(WorkMin) as varchar) as ReturnStr " + Environment.NewLine +
                                               "  from RunSheetA " + Environment.NewLine +
                                               " where BuDate between '" + vRiskDateStr_S + "' and '" + vRiskDateStr_E + "' " + Environment.NewLine +
                                               "   and Driver = '" + vEmpNo_Excel + "' " + Environment.NewLine +
                                               " group by Driver ";
                                    vReturnStr = PF.GetValue(vConnStr, vTempStr, "ReturnStr");
                                    if ((vReturnStr.Trim() != "") && (vReturnStr.IndexOf(',') >= 0))
                                    {
                                        vStrStore = vReturnStr.Split(',');
                                        vWorkDaysStr = (vStrStore[0] != "") ? vStrStore[0] : "NULL";
                                        if (vStrStore[1] != "")
                                        {
                                            vWorkMins = Int32.Parse(vStrStore[1]);
                                            vWorkHours = vWorkMins / 60;
                                            vWorkHoursStr = "'" + vWorkHours.ToString() + ":" + (vWorkMins - (vWorkHours * 60)).ToString("D2") + "'";
                                        }
                                        else
                                        {
                                            vWorkHoursStr = "NULL";
                                        }
                                    }
                                    else
                                    {
                                        vWorkHoursStr = "NULL";
                                        vWorkDaysStr = "NULL";
                                    }
                                    vTempStr = "select IDCardNo from DriverRiskLevel where IndexNo = '" + vIndexNo_Temp + "' ";
                                    vIDCardNo_Temp = PF.GetValue(vConnStr, vTempStr, "IDCardNo");
                                    if (vIDCardNo_Temp == "")
                                    {
                                        vTempStr = "insert into DriverRiskLevel " + Environment.NewLine +
                                                   "       (IndexNo, CalYear, CalYM, IDCardNo, RiskDate, CompanyLevel, DoctorLevel, WorkDays, WorkHours, BuMan, BuDate)" + Environment.NewLine +
                                                   "values ('" + vIndexNo_Temp + "', '" + vCalYear_Excel + "', '" + vCalYM_Excel + "', " + Environment.NewLine +
                                                   "        '" + vIDCardNo_ERP + "', " + vRiskDateStr_Temp + ", " + Environment.NewLine +
                                                   "        " + vCompanyRiskLevel + ", " + vDoctorRiskLevel + ", " + Environment.NewLine +
                                                   "        " + vWorkDaysStr + ", " + vWorkHoursStr + ", " + Environment.NewLine +
                                                   "        '" + vLoginID + "', GetDate())";
                                    }
                                    else
                                    {
                                        vTempStr = "update DriverRiskLevel set RiskDate = " + vRiskDateStr_Temp + ", " + Environment.NewLine +
                                                   "                           CompanyLevel = " + vCompanyRiskLevel + ", " + Environment.NewLine +
                                                   "                           DoctorLevel = " + vDoctorRiskLevel + ", " + Environment.NewLine +
                                                   "                           WorkDays = " + vWorkDaysStr + ", " + Environment.NewLine +
                                                   "                           WorkHours = " + vWorkHoursStr + ", " + Environment.NewLine +
                                                   "                           ModifyMan = '" + vLoginID + "', " + Environment.NewLine +
                                                   "                           ModifyDate = GetDate() " + Environment.NewLine +
                                                   " where IndexNo = '" + vIndexNo_Temp + "' ";
                                    }
                                    if (vTempStr != "")
                                    {
                                        PF.ExecSQL(vConnStr, vTempStr);
                                    }
                                }
                            }
                        }
                        break;
                }
                //EXCEL 的資料轉入完畢之後再把 EXCEL 檔案月份前一個月有資料但是當月沒有列入評估人員的評鑑等級複製到當月
                if (vCalYM_Excel.Trim() != "")
                {
                    vTempStr = "select Distinct IDCardNo " + Environment.NewLine +
                               "  from DriverRiskLevel " + Environment.NewLine +
                               " where CalYM < '" + vCalYM_Excel + "' " + Environment.NewLine +
                               "   and IDCardNo not in (" + Environment.NewLine +
                               "                        select IDCardNo from DriverRiskLevel where CalYM = '" + vCalYM_Excel + "' " + Environment.NewLine +
                               "                        )" + Environment.NewLine +
                               " order by IDCardNo ";
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        if (drTemp.HasRows)
                        {
                            vRiskDateStr_S = (Int32.Parse(vCalYM_Excel.Substring(0, 3)) + 1911).ToString() + "/" + (Int32.Parse(vCalYM_Excel.Substring(3, 2))).ToString("D2") + "/01";
                            vRiskDateStr_E = PF.GetMonthLastDay(DateTime.Parse(Int32.Parse(vCalYM_Excel.Substring(0, 3)).ToString() + "/" + (Int32.Parse(vCalYM_Excel.Substring(3, 2))).ToString("D2") + "/01"), "B");
                            while (drTemp.Read())
                            {
                                try
                                {
                                    //取回人員工號
                                    vTempStr = "select EmpNo " + Environment.NewLine +
                                               "  from Employee " + Environment.NewLine +
                                               " where AssumeDay <= '" + vRiskDateStr_E + "' " + Environment.NewLine +
                                               "   and isnull(LeaveDay, '9999/12/31') >= '" + vRiskDateStr_S + "' " + Environment.NewLine +
                                               "   and IDCardNo = '" + drTemp[0].ToString().Trim() + "' ";
                                    vEmpNo_Excel = PF.GetValue(vConnStr, vTempStr, "EmpNo");
                                    //計算人員當月工時
                                    vTempStr = "select cast(count(AssignNo) as varchar) + ',' + cast(sum(WorkHR * 60) + sum(WorkMin) as varchar) as ReturnStr " + Environment.NewLine +
                                               "  from RunSheetA " + Environment.NewLine +
                                               " where BuDate between '" + vRiskDateStr_S + "' and '" + vRiskDateStr_E + "' " + Environment.NewLine +
                                               "   and Driver = '" + vEmpNo_Excel + "' " + Environment.NewLine +
                                               " group by Driver ";
                                    vReturnStr = PF.GetValue(vConnStr, vTempStr, "ReturnStr");
                                    if ((vReturnStr.Trim() != "") && (vReturnStr.IndexOf(',') >= 0))
                                    {
                                        vStrStore = vReturnStr.Split(',');
                                        vWorkDaysStr = (vStrStore[0] != "") ? vStrStore[0] : "NULL";
                                        if (vStrStore[1] != "")
                                        {
                                            vWorkMins = Int32.Parse(vStrStore[1]);
                                            vWorkHours = vWorkMins / 60;
                                            vWorkHoursStr = "'" + vWorkHours.ToString() + ":" + (vWorkMins - (vWorkHours * 60)).ToString("D2") + "'";
                                        }
                                        else
                                        {
                                            vWorkHoursStr = "NULL";
                                        }
                                    }
                                    else
                                    {
                                        vWorkHoursStr = "NULL";
                                        vWorkDaysStr = "NULL";
                                    }
                                    //取回前一月份的評估等級
                                    vTempStr = "select top 1 (DoctorLevel + ',' + CompanyLevel) as ResultStr " + Environment.NewLine +
                                               "  from DriverRiskLevel " + Environment.NewLine +
                                               " where IDCardNo = '" + drTemp[0].ToString().Trim() + "' " + Environment.NewLine +
                                               "   and CalYM < '" + vCalYM_Excel + "' " + Environment.NewLine +
                                               " order by CalYM DESC";
                                    vReturnStr = PF.GetValue(vConnStr, vTempStr, "ResultStr");
                                    if ((vReturnStr.Trim() != "") && (vReturnStr.IndexOf(',') >= 0))
                                    {
                                        vStrStore = vReturnStr.Split(',');
                                        vDoctorRiskLevel = (vStrStore[0] != "") ? "'" + vStrStore[0] + "'" : "NULL";
                                        vCompanyRiskLevel = (vStrStore[1] != "") ? "'" + vStrStore[1] + "'" : "NULL";
                                        vIndexNo_Temp = vCalYM_Excel + drTemp[0].ToString().Trim();

                                        vTempStr = "insert into DriverRiskLevel " + Environment.NewLine +
                                                   "(IndexNo, CalYear, CalYM, IDCardNo, DoctorLevel, CompanyLevel, WorkHours, WorkDays, BuMan, BuDate)" + Environment.NewLine +
                                                   "values" + Environment.NewLine +
                                                   "('" + vIndexNo_Temp + "', '" + vCalYear_Excel + "', '" + vCalYM_Excel + "', '" + drTemp[0].ToString().Trim() + "', " + Environment.NewLine +
                                                   " " + vDoctorRiskLevel + ", " + vCompanyRiskLevel + ", " + Environment.NewLine +
                                                   " " + vWorkHoursStr + ", " + vWorkDaysStr + ", " + Environment.NewLine +
                                                   " '" + vLoginID + "', GetDate())";
                                        PF.ExecSQL(vConnStr, vTempStr);
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
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇要匯入的 EXCEL 檔')");
                Response.Write("</" + "Script>");
                fuExcel.Focus();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCLoseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plShowData.Visible = true;
            plSearch.Visible = true;
        }

        protected void fvDriverRiskLevelDetail_DataBound(object sender, EventArgs e)
        {
            string vCaseDateURL = "";
            string vCaseDateScript = "";
            string vTempStr = "";
            string vResultStr = "";
            string[] vStrStore;

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvDriverRiskLevelDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    Label eIndexNo_Edit = (Label)fvDriverRiskLevelDetail.FindControl("eIndexNo_Edit");
                    if (eIndexNo_Edit != null)
                    {
                        TextBox eRiskDate_Edit = (TextBox)fvDriverRiskLevelDetail.FindControl("eRiskDate_Edit");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eRiskDate_Edit.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eRiskDate_Edit.Attributes["onClick"] = vCaseDateScript;
                    }
                    break;
                case FormViewMode.Insert:
                    Label eIndexNo_INS = (Label)fvDriverRiskLevelDetail.FindControl("eIndexNo_INS");
                    if (eIndexNo_INS != null)
                    {
                        string vIDCardNo_INS = gridDriverRiskLevelList.SelectedRow.Cells[1].Text.Trim();
                        Label eIDCardNo_INS = (Label)fvDriverRiskLevelDetail.FindControl("eIDCardNo_INS");
                        Label eEmpNo_INS = (Label)fvDriverRiskLevelDetail.FindControl("eEmpNo_INS");
                        Label eEmpName_INS = (Label)fvDriverRiskLevelDetail.FindControl("eEmpName_INS");
                        Label eDepNo_INS = (Label)fvDriverRiskLevelDetail.FindControl("eDepNo_INS");
                        Label eDepName_INS = (Label)fvDriverRiskLevelDetail.FindControl("eDepName_INS");
                        TextBox eRiskDate_INS = (TextBox)fvDriverRiskLevelDetail.FindControl("eRiskDate_INS");
                        vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eRiskDate_INS.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eRiskDate_INS.Attributes["onClick"] = vCaseDateScript;
                        vTempStr = "select top 1 (EmpNo + ',' + [Name] + ',' + DepNo + ',' + (select [Name] from Department where DepNo = e.DepNo)) ResultStr " + Environment.NewLine +
                                   "  from Employee e " + Environment.NewLine +
                                   " where IDCardNo = '" + vIDCardNo_INS + "' " + Environment.NewLine +
                                   " order by AssumeDay DESC ";
                        vResultStr = PF.GetValue(vConnStr, vTempStr, "ResultStr");
                        vStrStore = vResultStr.Split(',');
                        eEmpNo_INS.Text = vStrStore[0];
                        eEmpName_INS.Text = vStrStore[1];
                        eDepNo_INS.Text = vStrStore[2];
                        eDepName_INS.Text = vStrStore[3];
                    }
                    break;
                default:
                    break;
            }
        }

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eIndexNo_Edit = (Label)fvDriverRiskLevelDetail.FindControl("eIndexNo_Edit");
            if ((eIndexNo_Edit != null) && (eIndexNo_Edit.Text.Trim() != ""))
            {
                try
                {
                    //取回異動檔的序號
                    string vIndexNo_Edit = eIndexNo_Edit.Text.Trim();
                    string vIndexNoItem_Edit = "";
                    string vMaxItems_Old = "";
                    string vItems_Edit = "";
                    string vTempStr = "select MAX(Items) MaxItem from DriverRiskLevel_History where IndexNo = '" + vIndexNo_Edit + "' ";
                    vMaxItems_Old = PF.GetValue(vConnStr, vTempStr, "MaxItem");
                    vItems_Edit = (vMaxItems_Old != "") ? (Int32.Parse(vMaxItems_Old) + 1).ToString("D4") : "0001";
                    vIndexNoItem_Edit = vIndexNo_Edit + vItems_Edit;
                    //先複製一份到異動檔
                    vTempStr = "INSERT INTO DriverRiskLevel_History " + Environment.NewLine +
                               "       (IndexNoItem, Items, ModifyMode, IndexNo, CalYear, CalYM, IDCardNo, " + Environment.NewLine +
                               "        RiskDate, DoctorLevel, CompanyLevel, WorkHours, WorkDays, " + Environment.NewLine +
                               "        Remark, BuMan, BuDate, ModifyMan, ModifyDate) " + Environment.NewLine +
                               "select '" + vIndexNoItem_Edit + "', '" + vItems_Edit + "', 'Edit', IndexNo, CalYear, CalYM, IDCardNo, " + Environment.NewLine +
                               "       RiskDate, DoctorLevel, CompanyLevel, WorkHours, WorkDays, " + Environment.NewLine +
                               "       Remark, BuMan, BuDate, ModifyMan, ModifyDate " + Environment.NewLine +
                               "  from DriverRiskLevel " + Environment.NewLine +
                               " where IndexNo = '" + vIndexNo_Edit + "' ";
                    PF.ExecSQL(vConnStr, vTempStr);
                    //開始實際上更新資料
                    TextBox eRiskDate_Edit = (TextBox)fvDriverRiskLevelDetail.FindControl("eRiskDate_Edit");
                    TextBox eDoctorLevel_Edit = (TextBox)fvDriverRiskLevelDetail.FindControl("eDoctorLevel_Edit");
                    TextBox eCompanyLevel_Edit = (TextBox)fvDriverRiskLevelDetail.FindControl("eCompanyLevel_Edit");
                    TextBox eWorkHours_Edit = (TextBox)fvDriverRiskLevelDetail.FindControl("eWorkHours_Edit");
                    TextBox eWorkDays_Edit = (TextBox)fvDriverRiskLevelDetail.FindControl("eWorkDays_Edit");
                    TextBox eRemark_Edit = (TextBox)fvDriverRiskLevelDetail.FindControl("eRemark_Edit");

                    vTempStr = "UPDATE DriverRiskLevel " + Environment.NewLine +
                               "   SET RiskDate = @RiskDate, DoctorLevel = @DoctorLevel, CompanyLevel = @CompanyLevel, " + Environment.NewLine +
                               "       WorkHours = @WorkHours, WorkDays = @WorkDays, " + Environment.NewLine +
                               "       Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate " + Environment.NewLine +
                               " WHERE (IndexNo = @IndexNo)";
                    sdsDriverRiskLevelDetail.UpdateParameters.Clear();
                    sdsDriverRiskLevelDetail.UpdateCommand = vTempStr;
                    sdsDriverRiskLevelDetail.UpdateParameters.Add(new Parameter("RiskDate", DbType.Date, (eRiskDate_Edit.Text.Trim() != "") ? DateTime.Parse(eRiskDate_Edit.Text.Trim()).ToShortDateString() : String.Empty));
                    sdsDriverRiskLevelDetail.UpdateParameters.Add(new Parameter("DoctorLevel", DbType.String, (eDoctorLevel_Edit.Text.Trim() != "") ? eDoctorLevel_Edit.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.UpdateParameters.Add(new Parameter("CompanyLevel", DbType.String, (eCompanyLevel_Edit.Text.Trim() != "") ? eCompanyLevel_Edit.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.UpdateParameters.Add(new Parameter("WorkHours", DbType.String, (eWorkHours_Edit.Text.Trim() != "") ? eWorkHours_Edit.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.UpdateParameters.Add(new Parameter("WorkDays", DbType.Int32, (eWorkDays_Edit.Text.Trim() != "") ? eWorkDays_Edit.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsDriverRiskLevelDetail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.Date, DateTime.Today.ToShortDateString()));
                    sdsDriverRiskLevelDetail.Update();

                    gridDriverRiskLevelList.DataBind();
                    fvDriverRiskLevelDetail.DataBind();
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

        /// <summary>
        /// 確定新增單筆資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eIndexNo_INS = (Label)fvDriverRiskLevelDetail.FindControl("eIndexNo_INS");
            if (eIndexNo_INS != null)
            {
                try
                {
                    Label eIDCardNo_INS = (Label)fvDriverRiskLevelDetail.FindControl("eIDCardNo_INS");
                    TextBox eRiskDate_INS = (TextBox)fvDriverRiskLevelDetail.FindControl("eRiskDate_INS");
                    TextBox eDoctorLevel_INS = (TextBox)fvDriverRiskLevelDetail.FindControl("eDoctorLevel_INS");
                    TextBox eCompanyLevel_INS = (TextBox)fvDriverRiskLevelDetail.FindControl("eCompanyLevel_INS");
                    TextBox eWorkHours_INS = (TextBox)fvDriverRiskLevelDetail.FindControl("eWorkHours_INS");
                    TextBox eWorkDays_INS = (TextBox)fvDriverRiskLevelDetail.FindControl("eWorkDays_INS");
                    TextBox eRemark_INS = (TextBox)fvDriverRiskLevelDetail.FindControl("eRemark_INS");

                    DateTime vRiskDate_INS = DateTime.Parse(eRiskDate_INS.Text.Trim());
                    string vIndexNo_INS = "";
                    string vCalYear_INS = (eRiskDate_INS.Text.Trim() != "") ? (vRiskDate_INS.Year - 1911).ToString() : (DateTime.Today.AddMonths(-1).Year - 1911).ToString();
                    string vCalYM_INS = (eRiskDate_INS.Text.Trim() != "") ? vCalYear_INS + vRiskDate_INS.Month.ToString("D2") : DateTime.Today.AddMonths(-1).Month.ToString("D2");
                    vIndexNo_INS = vCalYM_INS.Trim() + eIDCardNo_INS.Text.Trim();
                    string vTempStr = "INSERT INTO DriverRiskLevel " + Environment.NewLine +
                                      "       (IndexNo, CalYear, CalYM, IDCardNo, RiskDate, DoctorLevel, CompanyLevel, " + Environment.NewLine +
                                      "        WorkHours, WorkDays, Remark, BuMan, BuDate) " + Environment.NewLine +
                                      "values (@IndexNo, @CalYear, @CalYM, @IDCardNo, @RiskDate, @DoctorLevel, @CompanyLevel, " + Environment.NewLine +
                                      "        @WorkHours, @WorkDays, @Remark, @BuMan, @BuDate) ";
                    sdsDriverRiskLevelDetail.InsertCommand = vTempStr;
                    sdsDriverRiskLevelDetail.InsertParameters.Clear();
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_INS));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("CalYear", DbType.String, vCalYear_INS));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("CalYM", DbType.String, vCalYM_INS));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("RiskDate", DbType.Date, (eRiskDate_INS.Text.Trim() != "") ? vRiskDate_INS.ToShortDateString() : String.Empty));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("DoctorLevel", DbType.String, (eDoctorLevel_INS.Text.Trim() != "") ? eDoctorLevel_INS.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("CompanyLevel", DbType.String, (eCompanyLevel_INS.Text.Trim() != "") ? eCompanyLevel_INS.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("WorkHours", DbType.String, (eWorkHours_INS.Text.Trim() != "") ? eWorkHours_INS.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("WorkDays", DbType.Int32, (eWorkDays_INS.Text.Trim() != "") ? eWorkDays_INS.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark_INS.Text.Trim() != "") ? eRemark_INS.Text.Trim() : String.Empty));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsDriverRiskLevelDetail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.Today.ToShortDateString()));
                    sdsDriverRiskLevelDetail.Insert();

                    gridDriverRiskLevelDetail.DataBind();
                    fvDriverRiskLevelDetail.DataBind();
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

        /// <summary>
        /// 刪除單筆資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eIndexNo_List = (Label)fvDriverRiskLevelDetail.FindControl("eIndexNo_List");
            if ((eIndexNo_List != null) && (eIndexNo_List.Text.Trim() != ""))
            {
                try
                {
                    //取回異動檔的序號
                    string vIndexNo_List = eIndexNo_List.Text.Trim();
                    string vIndexNoItem_List = "";
                    string vMaxItems_Old = "";
                    string vItems_List = "";
                    string vTempStr = "select MAX(Items) MaxItem from DriverRiskLevel_History where IndexNo = '" + vIndexNo_List + "' ";
                    vMaxItems_Old = PF.GetValue(vConnStr, vTempStr, "MaxItem");
                    vItems_List = (vMaxItems_Old != "") ? (Int32.Parse(vMaxItems_Old) + 1).ToString("D4") : "0001";
                    vIndexNoItem_List = vIndexNo_List + vItems_List;
                    //先複製一份到異動檔
                    vTempStr = "INSERT INTO DriverRiskLevel_History " + Environment.NewLine +
                               "       (IndexNoItem, Items, ModifyMode, IndexNo, CalYear, CalYM, IDCardNo, " + Environment.NewLine +
                               "        RiskDate, DoctorLevel, CompanyLevel, WorkHours, WorkDays, " + Environment.NewLine +
                               "        Remark, BuMan, BuDate, ModifyMan, ModifyDate) " + Environment.NewLine +
                               "select '" + vIndexNoItem_List + "', '" + vItems_List + "', 'DEL', IndexNo, CalYear, CalYM, IDCardNo, " + Environment.NewLine +
                               "       RiskDate, DoctorLevel, CompanyLevel, WorkHours, WorkDays, " + Environment.NewLine +
                               "       Remark, BuMan, BuDate, ModifyMan, ModifyDate " + Environment.NewLine +
                               "  from DriverRiskLevel " + Environment.NewLine +
                               " where IndexNo = '" + vIndexNo_List + "' ";
                    PF.ExecSQL(vConnStr, vTempStr);
                    //實際刪除資料
                    sdsDriverRiskLevelDetail.DeleteParameters.Clear();
                    sdsDriverRiskLevelDetail.DeleteCommand = "delete DriverRiskLevel where IndexNo = @IndexNo_List ";
                    sdsDriverRiskLevelDetail.DeleteParameters.Add(new Parameter("IndexNo", DbType.String, vIndexNo_List));
                    sdsDriverRiskLevelDetail.Delete();
                    gridDriverRiskLevelList.DataBind();
                    fvDriverRiskLevelDetail.DataBind();
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

        /// <summary>
        /// 勾選是否取得全年度資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void cbGetFullYear_Search_CheckedChanged(object sender, EventArgs e)
        {
            eRiskMonthS_Search.Enabled = (cbGetFullYear_Search.Checked == false);
            eRiskYearE_Search.Enabled = (cbGetFullYear_Search.Checked == false);
            eRiskMonthE_Search.Enabled = (cbGetFullYear_Search.Checked == false);
        }

        protected void gridDriverRiskLevelList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridDriverRiskLevelList.PageIndex = e.NewPageIndex;
            gridDriverRiskLevelList.DataBind();
        }

        protected void gridDriverRiskLevelList_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vSelectedIDCardNo = gridDriverRiskLevelList.SelectedRow.Cells[1].Text.Trim();
            plShowData_Detail.Visible = (vSelectedIDCardNo != "");
            plShowData_Detail2.Visible = (vSelectedIDCardNo != "");
            if (vSelectedIDCardNo != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSelectEmpNo = PF.GetValue(vConnStr, "select top 1 EmpNo from Employee where IDCardNo = '" + vSelectedIDCardNo + "' order by Assumeday DESC", "EmpNo");
                string vCalDate_S = (cbGetFullYear_Search.Checked) ?
                                    (Int32.Parse(eRiskYearS_Search.Text.Trim()) + 1911).ToString() + "/01/01" :
                                    (Int32.Parse(eRiskYearS_Search.Text.Trim()) + 1911).ToString() + "/" + Int32.Parse(eRiskMonthS_Search.SelectedValue).ToString("D2") + "/01";
                string vCalDate_E = (cbGetFullYear_Search.Checked) ?
                                    (Int32.Parse(eRiskYearS_Search.Text.Trim()) + 1911).ToString() + "/12/31" :
                                    PF.GetMonthLastDay(DateTime.Parse((Int32.Parse(eRiskYearE_Search.Text.Trim())).ToString() + "/" + Int32.Parse(eRiskMonthE_Search.SelectedValue).ToString("D2") + "/01"), "B");
                string vWStr_CalYM = (cbGetFullYear_Search.Checked) ?
                                     "   and CalYear = '" + eRiskYearS_Search.Text.Trim() + "' " :
                                     "   and CalYM >= '" + eRiskYearS_Search.Text.Trim() + Int32.Parse(eRiskMonthS_Search.SelectedValue).ToString("D2") + "' and CalYM <= '" + eRiskYearE_Search.Text.Trim() + Int32.Parse(eRiskMonthE_Search.SelectedValue).ToString("D2") + "' ";
                /* 2023.04.17 修正，時數和工作天數按月取合計，而不是取全年合計
                string vSelStr = "select IndexNo, CalYear, CalYM, IDCardNo, " + Environment.NewLine +
                                 "       (select top 1 EmpNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) EmpNo, " + Environment.NewLine +
                                 "       (select top 1 [Name] from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) EmpName, " + Environment.NewLine +
                                 "       (select top 1 DepNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) DepNo, " + Environment.NewLine +
                                 "       (select [Name] from Department where DepNo = (select top 1 DepNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC)) DepName, " + Environment.NewLine +
                                 //"       RiskDate, DoctorLevel, CompanyLevel, WorkHours, WorkDays, Remark " + Environment.NewLine +
                                 "       RiskDate, DoctorLevel, CompanyLevel, " + Environment.NewLine +
                                 "       (select round(cast(sum(WorkHR * 60) + sum(WorkMin) as float) / 60.0, 2) from RunSheetA where BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and Driver = '" + vSelectEmpNo + "') WorkHours, " + Environment.NewLine +
                                 "       (select count(AssignNo) from RunSheetA where BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and Driver = '" + vSelectEmpNo + "') WorkDays " + Environment.NewLine +
                                 "  from DriverRiskLevel d " + Environment.NewLine +
                                 " where d.IDCardNo = '" + vSelectedIDCardNo + "' " + Environment.NewLine +
                                 vWStr_CalYM +
                                 " order by CalYM ";
                //*/
                string vSelStr = "select IndexNo, CalYear, d.CalYM, IDCardNo, " + Environment.NewLine +
                                 "       (select top 1 EmpNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) EmpNo, " + Environment.NewLine +
                                 "       (select top 1 [Name] from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) EmpName, " + Environment.NewLine +
                                 "       (select top 1 DepNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) DepNo, " + Environment.NewLine +
                                 "       (select [Name] from Department where DepNo = (select top 1 DepNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC)) DepName, " + Environment.NewLine +
                                 "       RiskDate, DoctorLevel, CompanyLevel, bb.WorkHours, bb.WorkDays " + Environment.NewLine +
                                 "  from DriverRiskLevel d join " + Environment.NewLine +
                                 "       ( " + Environment.NewLine +
                                 "        select (cast((year(BuDate) - 1911) as varchar) + right('00' + cast(month(BuDate) as varchar), 2)) as CalYM, Driver, round(cast(sum(WorkHR * 60) + sum(WorkMin) as float) / 60.0, 2) as WorkHours, count(AssignNo) as WorkDays " + Environment.NewLine +
                                 "          from RunSheetA " + Environment.NewLine +
                                 "         where Driver = '" + vSelectEmpNo + "' " + Environment.NewLine +
                                 "           and BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' " + Environment.NewLine +
                                 "         group by year(BuDate), Month(BuDate), Driver " + Environment.NewLine +
                                 "       ) bb on bb.CalYM = d.CalYM" + Environment.NewLine +
                                 " where d.IDCardNo = '" + vSelectedIDCardNo + "' " + Environment.NewLine +
                                 vWStr_CalYM +
                                 " order by d.CalYM ";
                //===========================================================================================================================================
                sdsDriverRiskLevelList.SelectCommand = vSelStr;
                gridDriverRiskLevelList.DataBind();
            }
        }

        protected void gridDriverRiskLevelDetail_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridDriverRiskLevelDetail.PageIndex = e.NewPageIndex;
            gridDriverRiskLevelDetail.DataBind();
        }

        protected void gridDriverRiskLevelDetail_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vIndexNo = gridDriverRiskLevelDetail.SelectedRow.Cells[1].Text.Trim();
            if (vIndexNo != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vEmpNo_Data = gridDriverRiskLevelDetail.SelectedRow.Cells[5].Text.Trim();
                string vCalDate_S = (cbGetFullYear_Search.Checked) ?
                                    (Int32.Parse(eRiskYearS_Search.Text.Trim()) + 1911).ToString() + "/01/01" :
                                    (Int32.Parse(eRiskYearS_Search.Text.Trim()) + 1911).ToString() + "/" + Int32.Parse(eRiskMonthS_Search.SelectedValue).ToString("D2") + "/01";
                string vCalDate_E = (cbGetFullYear_Search.Checked) ?
                                    (Int32.Parse(eRiskYearS_Search.Text.Trim()) + 1911).ToString() + "/12/31" :
                                    PF.GetMonthLastDay(DateTime.Parse((Int32.Parse(eRiskYearE_Search.Text.Trim())).ToString() + "/" + Int32.Parse(eRiskMonthE_Search.SelectedValue).ToString("D2") + "/01"), "B");
                string vWStr_CalYM = (cbGetFullYear_Search.Checked) ?
                                     "   and CalYear = '" + eRiskYearS_Search.Text.Trim() + "' " :
                                     "   and CalYM >= '" + eRiskYearS_Search.Text.Trim() + Int32.Parse(eRiskMonthS_Search.SelectedValue).ToString("D2") + "' and CalYM <= '" + eRiskYearE_Search.Text.Trim() + Int32.Parse(eRiskMonthE_Search.SelectedValue).ToString("D2") + "' ";
                /* 2023.04.17 修正，時數和工作天數按月取合計，而不是取全年合計
                string vSelStr = "select IndexNo, CalYear, CalYM, IDCardNo, " + Environment.NewLine +
                                 "       (select top 1 EmpNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) EmpNo, " + Environment.NewLine +
                                 "       (select top 1 [Name] from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) EmpName, " + Environment.NewLine +
                                 "       (select top 1 DepNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) DepNo, " + Environment.NewLine +
                                 "       (select [Name] from Department where DepNo = (select top 1 DepNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC)) DepName, " + Environment.NewLine +
                                 //"       RiskDate, DoctorLevel, CompanyLevel, WorkHours, WorkDays, Remark " + Environment.NewLine +
                                 "       RiskDate, DoctorLevel, CompanyLevel, Remark, " + Environment.NewLine +
                                 "       (select round(cast(sum(WorkHR * 60) + sum(WorkMin) as float) / 60.0, 2) from RunSheetA where BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and Driver = '" + vEmpNo_Data + "') WorkHours, " + Environment.NewLine +
                                 "       (select count(AssignNo) from RunSheetA where BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' and Driver = '" + vEmpNo_Data + "') WorkDays, " + Environment.NewLine +
                                 "       BuMan, (select [Name] from Employee where (EmpNo = d.BuMan)) AS BuManName, BuDate, " + Environment.NewLine +
                                 "       ModifyMan, (select [Name] from Employee where (EmpNo = d.ModifyMan)) AS ModifyManName, ModifyDate " + Environment.NewLine +
                                 "  from DriverRiskLevel d " + Environment.NewLine +
                                 " where d.IndexNo = '" + vIndexNo + "' " + Environment.NewLine +
                                 vWStr_CalYM +
                                 " order by CalYM ";
                //*/
                string vSelStr = "select IndexNo, CalYear, d.CalYM, IDCardNo, " + Environment.NewLine +
                                 "       (select top 1 EmpNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) EmpNo, " + Environment.NewLine +
                                 "       (select top 1 [Name] from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) EmpName, " + Environment.NewLine +
                                 "       (select top 1 DepNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC) DepNo, " + Environment.NewLine +
                                 "       (select [Name] from Department where DepNo = (select top 1 DepNo from Employee where IDCardNo = d.IDCardNo order by AssumeDay DESC)) DepName, " + Environment.NewLine +
                                 "       RiskDate, DoctorLevel, CompanyLevel, Remark, bb.WorkHours, bb.WorkDays, " + Environment.NewLine +
                                 "       BuMan, (select [Name] from Employee where (EmpNo = d.BuMan)) AS BuManName, BuDate, " + Environment.NewLine +
                                 "       ModifyMan, (select [Name] from Employee where (EmpNo = d.ModifyMan)) AS ModifyManName, ModifyDate " + Environment.NewLine +
                                 "  from DriverRiskLevel d join " + Environment.NewLine +
                                 "       ( " + Environment.NewLine +
                                 "        select (cast((year(BuDate) - 1911) as varchar) + right('00' + cast(month(BuDate) as varchar), 2)) as CalYM, Driver, round(cast(sum(WorkHR * 60) + sum(WorkMin) as float) / 60.0, 2) as WorkHours, count(AssignNo) as WorkDays " + Environment.NewLine +
                                 "          from RunSheetA " + Environment.NewLine +
                                 "         where Driver = '" + vEmpNo_Data + "' " + Environment.NewLine +
                                 "           and BuDate between '" + vCalDate_S + "' and '" + vCalDate_E + "' " + Environment.NewLine +
                                 "         group by year(BuDate), Month(BuDate), Driver " + Environment.NewLine +
                                 "       ) bb on bb.CalYM = d.CalYM" + Environment.NewLine +
                                 " where d.IndexNo = '" + vIndexNo + "' " + Environment.NewLine +
                                 vWStr_CalYM +
                                 " order by d.CalYM ";
                //===========================================================================================================================================
                sdsDriverRiskLevelDetail.SelectCommand = vSelStr;
                sdsDriverRiskLevelDetail.SelectParameters.Clear();
                fvDriverRiskLevelDetail.DataBind();
            }
        }
    }
}