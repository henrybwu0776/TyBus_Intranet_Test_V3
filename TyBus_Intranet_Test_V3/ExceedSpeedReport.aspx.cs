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
using System.Threading;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ExceedSpeedReport : System.Web.UI.Page
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
                Thread.CurrentThread.CurrentCulture = Cal;

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
                        plShowData.Visible = false;
                        plPrint.Visible = false;
                        eCaseDateYear_Search.Text = (DateTime.Today.Year - 1911).ToString().Trim();
                        eCaseDateMonth_Search.Text = DateTime.Today.Month.ToString().Trim();
                        ddlMonthStep_Search.SelectedIndex = 0;
                        DateTime vCalDate_Temp = DateTime.Parse(eCaseDateYear_Search.Text.Trim() + "/" + eCaseDateMonth_Search.Text.Trim() + "/01");
                        DateTime vCaseDateS_Temp = DateTime.Parse(PF.GetMonthFirstDay(vCalDate_Temp, "C"));
                        DateTime vCaseDateE_Temp = vCaseDateS_Temp.AddDays(9);
                        eCaseDateS_Search.Text = (vCaseDateS_Temp.Year - 1911).ToString() + "/" + vCaseDateS_Temp.ToString("MM/dd");
                        eCaseDateE_Search.Text = (vCaseDateE_Temp.Year - 1911).ToString() + "/" + vCaseDateE_Temp.ToString("MM/dd");
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

        /// <summary>
        /// 回傳 SQL 查詢語法
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            string vReturnStr = "";
            if ((eCaseDateYear_Search.Text.Trim() != "") && (eCaseDateMonth_Search.Text.Trim() != ""))
            {
                string vCalYM = eCaseDateYear_Search.Text.Trim() + Int32.Parse(eCaseDateMonth_Search.Text.Trim()).ToString("D2");
                string vWStr_CalYM = "   and e.CaseYM = '" + vCalYM.Trim() + "' ";
                string vMonthStep = ddlMonthStep_Search.SelectedValue;
                string vWStr_MonthStep = "   and e.MonthStep = '" + vMonthStep.Trim() + "' ";
                vReturnStr = "SELECT CaseNo, CaseYM, MonthStep, CaseDateS, CaseDateE, " + Environment.NewLine +
                             "       DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = e.DepNo)) AS DepName, " + Environment.NewLine +
                             "       Car_ID, Driver, DriverName, Exception, AbnormalValue, Attachment, Remark, " + Environment.NewLine +
                             "       Inspector, (select [Name] from Employee where EmpNo = e.Inspector) InspectorName, " + Environment.NewLine +
                             "       BuDate, BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = e.BuMan)) AS BuManName, " + Environment.NewLine +
                             "       ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = e.ModifyMan)) AS ModifyManName " + Environment.NewLine +
                             "  FROM ExceedSpeedReport AS e " + Environment.NewLine +
                             " WHERE (1 = 1) " + Environment.NewLine +
                             vWStr_CalYM + vWStr_MonthStep +
                             " order by e.DepNo, e.Car_ID";
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先輸入要查詢檢查表的年月及旬別')");
                Response.Write("</" + "Script>");
                eCaseDateS_Search.Focus();
            }
            return vReturnStr;
        }

        /// <summary>
        /// 開啟 SQL 資料庫
        /// </summary>
        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            sdsExceedSpeedReport_List.SelectCommand = vSelStr;
            gridExceedSpeedReport_List.DataSourceID = "sdsExceedSpeedReport_List";
            gridExceedSpeedReport_List.DataBind();
        }

        protected void ddlMonthStep_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eMonthStep_Search.Text = ddlMonthStep_Search.SelectedValue.Trim();
            DateTime vCalDate_Temp = DateTime.Parse(eCaseDateYear_Search.Text.Trim() + "/" + eCaseDateMonth_Search.Text.Trim() + "/01");
            DateTime vCaseDateS_Temp = DateTime.Today;
            DateTime vCaseDateE_Temp = DateTime.Today;
            switch (ddlMonthStep_Search.SelectedValue)
            {
                case "10": //上旬
                    vCaseDateS_Temp = DateTime.Parse(PF.GetMonthFirstDay(vCalDate_Temp, "C"));
                    vCaseDateE_Temp = vCaseDateS_Temp.AddDays(9);
                    break;
                case "11": //中旬
                    vCaseDateS_Temp = DateTime.Parse(PF.GetMonthFirstDay(vCalDate_Temp, "C")).AddDays(10);
                    vCaseDateE_Temp = vCaseDateS_Temp.AddDays(9);
                    break;
                case "12": //下旬
                    vCaseDateS_Temp = DateTime.Parse(PF.GetMonthFirstDay(vCalDate_Temp, "C")).AddDays(20);
                    vCaseDateE_Temp = DateTime.Parse(PF.GetMonthLastDay(vCalDate_Temp, "C"));
                    break;
            }
            eCaseDateS_Search.Text = (vCaseDateS_Temp.Year - 1911).ToString() + "/" + vCaseDateS_Temp.ToString("MM/dd");
            eCaseDateE_Search.Text = (vCaseDateE_Temp.Year - 1911).ToString() + "/" + vCaseDateE_Temp.ToString("MM/dd");
        }

        /// <summary>
        /// 產生指定旬別的空白檢查表 (38輛遊覽車每輛一筆資料)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCreate_Search_Click(object sender, EventArgs e)
        {
            if ((eCaseDateYear_Search.Text.Trim() != "") && (eCaseDateMonth_Search.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCalYM = eCaseDateYear_Search.Text.Trim() + eCaseDateMonth_Search.Text.Trim().PadLeft(2, '0');
                string vMonthStep = ddlMonthStep_Search.SelectedValue;
                DateTime vCaseDateS = DateTime.Parse(eCaseDateS_Search.Text.Trim());
                string vCaseDateS_Str = vCaseDateS.Year.ToString() + "/" + vCaseDateS.ToString("MM/dd");
                DateTime vCaseDateE = DateTime.Parse(eCaseDateE_Search.Text.Trim());
                string vCaseDateE_Str = vCaseDateE.Year.ToString() + "/" + vCaseDateE.ToString("MM/dd");
                string vSQLStr_Temp = "insert into ExceedSpeedReport(CaseNo, CaseYM, MonthStep, CaseDateS, CaseDateE, DepNo, Car_ID, Driver, DriverName, Inspector, BuDate, BuMan)" + Environment.NewLine +
                                      "select '" + vCalYM + "' + '" + vMonthStep + "' + CompanyNo + right('0000' + cast(ROW_NUMBER() OVER (ORDER BY CompanyNo, Car_ID) as varchar), 4) as CaseNo, " + Environment.NewLine +
                                      "       '" + vCalYM + "', '" + vMonthStep + "', '" + vCaseDateS_Str + "', '" + vCaseDateE_Str + "', CompanyNo, Car_ID, " + Environment.NewLine +
                                      "       Driver, (select [Name] from Employee where EmpNo = c.Driver) DriverName, '" + vLoginID + "', " + Environment.NewLine +
                                      "       GetDate(), '" + vLoginID + "' " + Environment.NewLine +
                                      "  from Car_InfoA c " + Environment.NewLine +
                                      " where c.Point = 'C1' " + Environment.NewLine +
                                      "   and c.Tran_Type = '1' " + Environment.NewLine +
                                      " order by CompanyNo, Car_ID";
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    OpenData();
                    int vGridRCount = gridExceedSpeedReport_List.Rows.Count;
                    plShowData.Visible = (vGridRCount > 0);
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先輸入要產生檢查表的年月及旬別')");
                Response.Write("</" + "Script>");
                eCaseDateS_Search.Focus();
            }
        }

        /// <summary>
        /// 查詢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSearch_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelStr = GetSelectStr();
            OpenData();
            int vGridRCount = gridExceedSpeedReport_List.Rows.Count;
            plShowData.Visible = (vGridRCount > 0);
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 預覽報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_Click(object sender, EventArgs e)
        {
            string vSelStr = GetSelectStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintPoint = new SqlDataAdapter(vSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintPoint = new DataTable("ExceedSpeedReportP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("ExceedSpeedReportP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\ExceedSpeedReportP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "各單位遊覽車超速檢查表"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;

                    string vCaseDateStr = eCaseDateYear_Search.Text.Trim() + " 年 " + eCaseDateMonth_Search.Text.Trim() + " 月 " + ddlMonthStep_Search.SelectedItem.Text.Trim();
                    string vRecordNote = "預覽報表_各單位遊覽車超速檢查表" + Environment.NewLine +
                                         "ExceedSpeedReport.aspx" + Environment.NewLine +
                                         "檢查日：" + vCaseDateStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        protected void fvExceedSpeedReport_Detail_DataBound(object sender, EventArgs e)
        {
            switch (fvExceedSpeedReport_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    DropDownList ddlException_List = (DropDownList)fvExceedSpeedReport_Detail.FindControl("ddlException_List");
                    Label eException_List = (Label)fvExceedSpeedReport_Detail.FindControl("eException_List");
                    if (ddlException_List != null)
                    {
                        ddlException_List.SelectedIndex = (eException_List.Text.Trim() == "有") ? 1 : 0;
                    }
                    break;
                case FormViewMode.Edit:
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    DropDownList ddlException_Edit = (DropDownList)fvExceedSpeedReport_Detail.FindControl("ddlException_Edit");
                    Label eException_Edit = (Label)fvExceedSpeedReport_Detail.FindControl("eException_Edit");
                    if (ddlException_Edit != null)
                    {
                        ddlException_Edit.SelectedIndex = (eException_Edit.Text.Trim() == "有") ? 1 : 0;
                    }
                    Label eModifyDate_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eModifyDate_Edit");
                    if (eModifyDate_Temp != null)
                    {
                        eModifyDate_Temp.Text = PF.GetChinsesDate(DateTime.Today);
                        Label eModifyMan_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eModifyMan_Edit");
                        Label eModifyManName_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eModifyManName_Edit");
                        eModifyMan_Temp.Text = vLoginID;
                        eModifyManName_Temp.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    }
                    break;
                case FormViewMode.Insert:
                    DropDownList ddlException_INS = (DropDownList)fvExceedSpeedReport_Detail.FindControl("ddlException_INS");
                    Label eException_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eException_INS");
                    if (ddlException_INS != null)
                    {
                        ddlException_INS.SelectedIndex = 0;
                        eException_INS.Text = "無";
                    }
                    Label eBudate_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eBuDate_INS");
                    if (eBudate_INS != null)
                    {
                        eBudate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                        Label eCaseDateS_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eCaseDateS_INS");
                        eCaseDateS_INS.Text = eCaseDateS_Search.Text.Trim();
                        Label eCaseDateE_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eCaseDateE_INS");
                        eCaseDateE_INS.Text = eCaseDateE_Search.Text.Trim();
                        TextBox eInspector_INS = (TextBox)fvExceedSpeedReport_Detail.FindControl("eInspector_INS");
                        eInspector_INS.Text = vLoginID;
                        Label eBuMan_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eBuMan_INS");
                        eBuMan_INS.Text = vLoginID;
                        Label eBuManName_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eBuManName_INS");
                        eBuManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                        Label eCaseYM_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eCaseYM_INS");
                        eCaseYM_INS.Text = eCaseDateYear_Search.Text.Trim() + eCaseDateMonth_Search.Text.Trim().PadLeft(2, '0');
                        Label eMonthStep_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eMonthStep_INS");
                        eMonthStep_INS.Text = eMonthStep_Search.Text.Trim();
                    }
                    break;
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox vDepNo_Temp = (TextBox)fvExceedSpeedReport_Detail.FindControl("eDepNo_Edit");
            Label vDepName_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eDepName_Edit");
            string vDepNoStr = vDepNo_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNoStr.Trim() + "' ";
            string vDepNameStr = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepNameStr == "")
            {
                vDepNameStr = vDepNoStr.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepNameStr.Trim() + "' ";
                vDepNoStr = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            vDepNo_Temp.Text = vDepNoStr.Trim();
            vDepName_Temp.Text = vDepNameStr.Trim();
        }

        protected void eDriver_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Temp = (TextBox)fvExceedSpeedReport_Detail.FindControl("eDriver_Edit");
            Label eDriverName_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eDriverName_Edit");
            string vDriver = eDriver_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver + "' ";
            string vDriverName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriver;
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Temp.Text = vDriver.Trim();
            eDriverName_Temp.Text = vDriverName.Trim();
        }

        protected void ddlException_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlException_Temp = (DropDownList)fvExceedSpeedReport_Detail.FindControl("ddlException_Edit");
            Label eException_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eException_Edit");
            eException_Temp.Text = ddlException_Temp.SelectedValue.Trim();
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox vDepNo_Temp = (TextBox)fvExceedSpeedReport_Detail.FindControl("eDepNo_INS");
            Label vDepName_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eDepName_INS");
            string vDepNoStr = vDepNo_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNoStr.Trim() + "' ";
            string vDepNameStr = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepNameStr == "")
            {
                vDepNameStr = vDepNoStr.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepNameStr.Trim() + "' ";
                vDepNoStr = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            vDepNo_Temp.Text = vDepNoStr.Trim();
            vDepName_Temp.Text = vDepNameStr.Trim();
        }

        protected void eDriver_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Temp = (TextBox)fvExceedSpeedReport_Detail.FindControl("eDriver_INS");
            Label eDriverName_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eDriverName_INS");
            string vDriver = eDriver_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver + "' ";
            string vDriverName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriver;
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Temp.Text = vDriver.Trim();
            eDriverName_Temp.Text = vDriverName.Trim();
        }

        protected void ddlException_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlException_Temp = (DropDownList)fvExceedSpeedReport_Detail.FindControl("ddlException_INS");
            Label eException_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eException_INS");
            eException_Temp.Text = ddlException_Temp.SelectedValue.Trim();
        }

        /// <summary>
        /// 刪除資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eCaseNo_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eCaseNo_List");
            if ((eCaseNo_Temp != null) && (eCaseNo_Temp.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCaseNo_Temp = eCaseNo_Temp.Text.Trim();
                string vSQLStr_Temp = "select isnull(MAX(Items), 0) MaxIndex from ExceedSpeedReport_History where CaseNo = '" + vCaseNo_Temp + "' ";
                int vMaxIndex = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex")) + 1;
                vSQLStr_Temp = "insert into ExceedSpeedReport_History" + Environment.NewLine +
                               "       (CaseNoItem, Items, ModifyMode, CaseNo, CaseYM, MonthStep, CaseDateS, CaseDateE, DepNo, Car_ID, " + Environment.NewLine +
                               "        Driver, DriverName, Exception, AbnormalValue, Attachment, Remark, Inspector, BuDate, BuMan, ModifyDate, ModifyMan)" + Environment.NewLine +
                               "select '" + vCaseNo_Temp + vMaxIndex.ToString("D4") + "', '" + vMaxIndex.ToString("D4") + "', 'DEL', CaseNo, CaseYM, MonthStep, " + Environment.NewLine +
                               "       CaseDateS, CaseDateE, DepNo, Car_ID, Driver, DriverName, Exception, AbnormalValue, Attachment, Remark, Inspector, " + Environment.NewLine +
                               "       BuDate, BuMan, GetDate(), '" + vLoginID + "' " + Environment.NewLine +
                               "  from ExceedSpeedReport " + Environment.NewLine +
                               " where CaseNo = '" + vCaseNo_Temp.Trim() + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    sdsExceedSpeedReport_Detail.DeleteParameters.Clear();
                    sdsExceedSpeedReport_Detail.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp));
                    sdsExceedSpeedReport_Detail.Delete();
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

        protected void sdsExceedSpeedReport_Detail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridExceedSpeedReport_List.DataBind();
            }
        }

        protected void sdsExceedSpeedReport_Detail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridExceedSpeedReport_List.DataBind();
            }
        }

        protected void sdsExceedSpeedReport_Detail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eCaseYM_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eCaseYM_INS");
            string vCaseYM_Str = eCaseYM_INS.Text.Trim();
            Label eMonthStep_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eMonthStep_INS");
            string vMonthStep_Str = (eMonthStep_INS.Text.Trim() != "") ? eMonthStep_INS.Text.Trim() : eMonthStep_Search.Text.Trim();
            TextBox eDepNo_INS = (TextBox)fvExceedSpeedReport_Detail.FindControl("eDepNo_INS");
            string vDepNo_Str = eDepNo_INS.Text.Trim();
            string vCaseNo_Temp = vCaseYM_Str + vMonthStep_Str + vDepNo_Str;
            string vSQLStr_Temp = "select max(CaseNo) MaxCaseNo from ExceedSpeedReport where CaseNo like '" + vCaseNo_Temp + "' ";
            string vIndex_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxCaseNo").Replace(vCaseNo_Temp.Trim(), "");
            int vIndex_Temp = (vIndex_Str != "") ? Int32.Parse(vIndex_Str) + 1 : 1;
            string vCaseNo = vCaseNo_Temp + vIndex_Temp.ToString("D4");
            Label eCaseDateS_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eCaseDateS_INS");
            DateTime vCaseDateS = DateTime.Parse(eCaseDateS_INS.Text.Trim());
            Label eCaseDateE_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eCaseDAteE_INS");
            DateTime vCaseDateE = DateTime.Parse(eCaseDateE_INS.Text.Trim());
            Label eBuDate_INS = (Label)fvExceedSpeedReport_Detail.FindControl("eBuDate_INS");
            DateTime vBuDate_Temp = (eBuDate_INS.Text.Trim() != "") ? DateTime.Parse(eBuDate_INS.Text.Trim()) : DateTime.Today;

            e.Command.Parameters["@CaseNo"].Value = vCaseNo;
            e.Command.Parameters["@CAseYM"].Value = vCaseYM_Str;
            e.Command.Parameters["@MonthStep"].Value = vMonthStep_Str;
            e.Command.Parameters["@CaseDateS"].Value = vCaseDateS;
            e.Command.Parameters["@CaseDateE"].Value = vCaseDateE;
            e.Command.Parameters["@BuDate"].Value = vBuDate_Temp;
        }

        protected void sdsExceedSpeedReport_Detail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridExceedSpeedReport_List.DataBind();
            }
        }

        protected void sdsExceedSpeedReport_Detail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            Label eCaseNo_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eCaseNo_Edit");
            if ((eCaseNo_Temp != null) && (eCaseNo_Temp.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCaseNo_Temp = eCaseNo_Temp.Text.Trim();
                string vSQLStr_Temp = "select isnull(MAX(Items), 0) MaxIndex from ExceedSpeedReport_History where CaseNo = '" + vCaseNo_Temp + "' ";
                int vMaxIndex = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex")) + 1;
                vSQLStr_Temp = "insert into ExceedSpeedReport_History" + Environment.NewLine +
                               "       (CaseNoItem, Items, ModifyMode, CaseNo, CaseYM, MonthStep, CaseDateS, CaseDateE, DepNo, Car_ID, " + Environment.NewLine +
                               "        Driver, DriverName, Exception, AbnormalValue, Attachment, Remark, Inspector, BuDate, BuMan, ModifyDate, ModifyMan)" + Environment.NewLine +
                               "select '" + vCaseNo_Temp + vMaxIndex.ToString("D4") + "', '" + vMaxIndex.ToString("D4") + "', 'EDIT', CaseNo, CaseYM, MonthStep, " + Environment.NewLine +
                               "       CaseDateS, CaseDateE, DepNo, Car_ID, Driver, DriverName, Exception, AbnormalValue, Attachment, Remark, " + Environment.NewLine +
                               "       Inspector, BuDate, BuMan, GetDate(), '" + vLoginID + "' " + Environment.NewLine +
                               "  from ExceedSpeedReport " + Environment.NewLine +
                               " where CaseNo = '" + vCaseNo_Temp.Trim() + "' ";
                Label eModifyDate_Temp = (Label)fvExceedSpeedReport_Detail.FindControl("eModifyDate_Edit");
                DateTime vModifyDate_Temp = (eModifyDate_Temp.Text.Trim() != "") ? DateTime.Parse(eModifyDate_Temp.Text.Trim()) : DateTime.Today;
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    e.Command.Parameters["@ModifyDate"].Value = vModifyDate_Temp;
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

        /// <summary>
        /// 結束報表預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
        }

        /// <summary>
        /// 匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

            string vHeaderText = "";
            int vLinesNo = 0;
            string vSelStr = GetSelectStr();
            string vFileName = "遊覽車超速檢查清冊";
            DateTime vBuDate;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                SqlCommand cmdExcel = new SqlCommand(vSelStr, connExcel);
                connExcel.Open();
                SqlDataReader drExcel = cmdExcel.ExecuteReader();
                if (drExcel.HasRows)
                {
                    //查詢結果有資料的時候才執行
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    vLinesNo = 0;
                    wsExcel.CreateRow(vLinesNo);
                    for (int i = 0; i < drExcel.FieldCount; i++)
                    {
                        vHeaderText = (drExcel.GetName(i) == "CaseNo") ? "序號" :
                                      (drExcel.GetName(i) == "CaseYM") ? "查核年月" :
                                      (drExcel.GetName(i) == "MonthStep") ? "旬別" :
                                      (drExcel.GetName(i) == "CaseDateS") ? "開始日" :
                                      (drExcel.GetName(i) == "CaseDateE") ? "結束日" :
                                      (drExcel.GetName(i) == "DepNo") ? "" :
                                      (drExcel.GetName(i) == "DepName") ? "站別" :
                                      (drExcel.GetName(i) == "Car_ID") ? "車號" :
                                      (drExcel.GetName(i) == "Driver") ? "" :
                                      (drExcel.GetName(i) == "DriverName") ? "駕駛員姓名" :
                                      (drExcel.GetName(i) == "Exception") ? "異常項目" :
                                      (drExcel.GetName(i) == "AbnormalValue") ? "異常數值" :
                                      (drExcel.GetName(i) == "Attachment") ? "附件編號" :
                                      (drExcel.GetName(i) == "Remark") ? "備註" :
                                      (drExcel.GetName(i) == "Inspector") ? "" :
                                      (drExcel.GetName(i) == "InspectorName") ? "查核人" :
                                      (drExcel.GetName(i) == "BuDate") ? "建檔日期" :
                                      (drExcel.GetName(i) == "BuMan") ? "" :
                                      (drExcel.GetName(i) == "BuManName") ? "建檔人" :
                                      (drExcel.GetName(i) == "ModifyDate") ? "異動日期" :
                                      (drExcel.GetName(i) == "ModifyMan") ? "" :
                                      (drExcel.GetName(i) == "ModifyManName") ? "異動人" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i) == "CaseDateS") ||
                                 (drExcel.GetName(i) == "CaseDateE") ||
                                 (drExcel.GetName(i) == "BuDate") ||
                                 (drExcel.GetName(i) == "ModifyDate")) && (drExcel[i].ToString() != ""))
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
                            string vCaseDateStr = eCaseDateYear_Search.Text.Trim() + " 年 " + eCaseDateMonth_Search.Text.Trim() + " 月 " + ddlMonthStep_Search.SelectedItem.Text.Trim();
                            string vRecordNote = "匯出檔案_各單位遊覽車超速檢查表" + Environment.NewLine +
                                                 "ExceedSpeedReport.aspx" + Environment.NewLine +
                                                 "檢查日：" + vCaseDateStr;
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
    }
}