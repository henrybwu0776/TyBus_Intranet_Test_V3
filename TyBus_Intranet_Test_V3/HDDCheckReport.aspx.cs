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
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class HDDCheckReport : System.Web.UI.Page
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
                    //事件日期
                    string vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_S_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    //查核日期
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCheckDate_S_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCheckDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCheckDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCheckDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        plSearch.Visible = true;
                        plPrint.Visible = false;
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

        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            gridHDDCheckreport_List.DataSourceID = "";
            sdsHDDCheckReport_List.SelectCommand = vSelStr;
            gridHDDCheckreport_List.DataSource = sdsHDDCheckReport_List;
            gridHDDCheckreport_List.DataBind();
        }

        private string GetSelectStr()
        {
            //事件日期
            DateTime eCaseDate_S_Temp = (eCaseDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_S_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_S_Temp = eCaseDate_S_Temp.Year.ToString() + "/" + eCaseDate_S_Temp.ToString("MM/dd");
            DateTime eCaseDate_E_Temp = (eCaseDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_E_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_E_Temp = eCaseDate_E_Temp.Year.ToString() + "/" + eCaseDate_E_Temp.ToString("MM/dd");
            string vWStr_CaseDate = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and h.CaseDate between '" + vCaseDate_S_Temp.Trim() + "' and '" + vCaseDate_E_Temp.Trim() + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? "   and h.CaseDate = '" + vCaseDate_S_Temp.Trim() + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and h.CaseDate = '" + vCaseDate_E_Temp.Trim() + "' " + Environment.NewLine : "";
            //查核日期
            DateTime eCheckDate_S_Temp = (eCheckDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eCheckDate_S_Search.Text.Trim()) : DateTime.Today;
            string vCheckdate_S_Temp = eCheckDate_S_Temp.Year.ToString() + "/" + eCheckDate_S_Temp.ToString("MM/dd");
            DateTime eCheckDate_E_Temp = (eCheckDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eCheckDate_E_Search.Text.Trim()) : DateTime.Today;
            string vCheckDate_E_Temp = eCheckDate_E_Temp.Year.ToString() + "/" + eCheckDate_E_Temp.ToString("MM/dd");
            string vWStr_CheckDate = ((eCheckDate_S_Search.Text.Trim() != "") && (eCheckDate_E_Search.Text.Trim() != "")) ? "   and h.CheckDate between '" + vCheckdate_S_Temp.Trim() + "' and '" + vCheckDate_E_Temp.Trim() + "' " + Environment.NewLine :
                                     ((eCheckDate_S_Search.Text.Trim() != "") && (eCheckDate_E_Search.Text.Trim() == "")) ? "   and h.CheckDate = '" + vCheckdate_S_Temp.Trim() + "' " + Environment.NewLine :
                                     ((eCheckDate_S_Search.Text.Trim() == "") && (eCheckDate_E_Search.Text.Trim() != "")) ? "   and h.CheckDate = '" + vCheckDate_E_Temp.Trim() + "' " + Environment.NewLine : "";
            //站別
            string vWStr_DepNo = ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   and h.DepNo between '" + eDepNo_S_Search.Text.Trim() + "' and '" + eDepNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? "   and h.DepNo = '" + eDepNo_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNo_S_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   and h.DepNo = '" + eDepNo_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            //駕駛員
            string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? "   and h.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            //車號
            string vWStr_CarID = (eCarID_Search.Text.Trim() != "") ? "   and h.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            //查核人
            string vWStr_Inspector = (eInspector_Search.Text.Trim() != "") ? "   and h.Inspector = '" + eInspector_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vReturnStr = "SELECT CaseNo, " + Environment.NewLine +
                                "       (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = h.DepNo)) AS DepName, " + Environment.NewLine +
                                "       Driver, DriverName, CaseDate, Car_ID, CheckNote, Remark, " + Environment.NewLine +
                                "       (SELECT NAME FROM EMPLOYEE WHERE(EMPNO = h.Inspector)) AS InspectorName, " + Environment.NewLine +
                                "       BuDate, (SELECT NAME FROM EMPLOYEE WHERE(EMPNO = h.BuMan)) AS BuManName, " + Environment.NewLine +
                                "       ModifyDate, (SELECT NAME FROM EMPLOYEE WHERE(EMPNO = h.ModifyMan)) AS ModifyManName, " + Environment.NewLine +
                                "       IsPassed, AdviceReportNo, CheckDate " + Environment.NewLine +
                                "  FROM HDDCheckReport AS h " + Environment.NewLine +
                                " WHERE (1 = 1)" + Environment.NewLine +
                                vWStr_CarID + vWStr_CaseDate + vWStr_CheckDate + vWStr_DepNo + vWStr_Driver + vWStr_Inspector +
                                " ORDER BY CaseNo DESC ";
            return vReturnStr;
        }

        protected void eDepNo_S_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_S_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_S_Search.Text = vDepNo_Temp.Trim();
            eDepName_S_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDepNo_E_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_E_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_E_Search.Text = vDepNo_Temp.Trim();
            eDepName_E_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDriver_Temp = eDriver_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp.Trim() + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp.Trim() == "")
            {
                vDriverName_Temp = vDriver_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Search.Text = vDriver_Temp.Trim();
            eDriverName_Search.Text = vDriverName_Temp.Trim();
        }

        protected void eInspector_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vInspector_Temp = eInspector_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vInspector_Temp.Trim() + "' ";
            string vInspectorName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vInspectorName_Temp.Trim() == "")
            {
                vInspectorName_Temp = vInspector_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vInspectorName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vInspector_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eInspector_Search.Text = vInspector_Temp.Trim();
            eInspectorName_Search.Text = vInspectorName_Temp.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 列印工作報告
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
                DataTable dtPrintPoint = new DataTable("HDDCheckReportP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("HDDCheckReportP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\HDDCheckReportP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "查核硬碟工作報告"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;

                    string vCaseDateStr = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "自 " + eCaseDate_S_Search.Text.Trim() + " 起至 " + eCaseDate_E_Search.Text.Trim() + " 止" :
                                          ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? eCaseDate_S_Search.Text.Trim() :
                                          ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? eCaseDate_E_Search.Text.Trim() : "";
                    string vDepNoStr = ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? "自 " + eDepNo_S_Search.Text.Trim() + " 站起至 " + eDepNo_E_Search.Text.Trim() + " 站止" :
                                       ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? eDepNo_S_Search.Text.Trim() + " 站" :
                                       ((eDepNo_S_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? eDepNo_E_Search.Text.Trim() + " 站" : "全部";
                    string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                    string vCheckDateStr = ((eCheckDate_S_Search.Text.Trim() != "") && (eCheckDate_E_Search.Text.Trim() != "")) ? "自 " + eCheckDate_S_Search.Text.Trim() + " 起至 " + eCheckDate_E_Search.Text.Trim() + " 止" :
                                           ((eCheckDate_S_Search.Text.Trim() != "") && (eCheckDate_E_Search.Text.Trim() == "")) ? eCheckDate_S_Search.Text.Trim() :
                                           ((eCheckDate_S_Search.Text.Trim() == "") && (eCheckDate_E_Search.Text.Trim() != "")) ? eCheckDate_E_Search.Text.Trim() : "";
                    string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "全部";
                    string vInspectorNoStr = (eInspector_Search.Text.Trim() != "") ? eInspector_Search.Text.Trim() : "";
                    string vRecordNote = "預覽報表_查核硬碟工作報告" + Environment.NewLine +
                                         "HDDCheckReport.aspx" + Environment.NewLine +
                                         "事件日期：" + vCaseDateStr + Environment.NewLine +
                                         "站別：" + vDepNoStr + Environment.NewLine +
                                         "駕駛員編號：" + vDriverStr + Environment.NewLine +
                                         "查核日期：" + vCheckDateStr + Environment.NewLine +
                                         "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                         "查核人：" + vInspectorNoStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        protected void gridHDDCheckreport_List_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridHDDCheckreport_List.PageIndex = e.NewPageIndex;
            gridHDDCheckreport_List.DataBind();
        }

        /// <summary>
        /// 資料模板狀態變更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void fvHDDCheckReport_Detail_DataBound(object sender, EventArgs e)
        {
            string vDateURL_Temp = "";
            string vDateScript_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvHDDCheckReport_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    CheckBox cbIsPassed_List = (CheckBox)fvHDDCheckReport_Detail.FindControl("cbIsPassed_List");
                    Label eIsPassed_List = (Label)fvHDDCheckReport_Detail.FindControl("eIsPassed_List");
                    if (cbIsPassed_List != null)
                    {
                        cbIsPassed_List.Checked = (eIsPassed_List.Text.Trim() == "True");

                        Button bbEdit_List = (Button)fvHDDCheckReport_Detail.FindControl("bbEdit_List");
                        Button bbDel_List = (Button)fvHDDCheckReport_Detail.FindControl("bbDel_List");
                        Button bbPassToAdvice_List = (Button)fvHDDCheckReport_Detail.FindControl("bbPassToAdvice_List");
                        bbEdit_List.Visible = !(eIsPassed_List.Text.Trim() == "True");
                        bbDel_List.Visible = !(eIsPassed_List.Text.Trim() == "True");
                        bbPassToAdvice_List.Visible = !(eIsPassed_List.Text.Trim() == "True");
                    }
                    break;
                case FormViewMode.Edit:
                    TextBox eCheckDate_Edit = (TextBox)fvHDDCheckReport_Detail.FindControl("eCheckDate_Edit");
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCheckDate_Edit.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCheckDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                    TextBox eCaseDate_Edit = (TextBox)fvHDDCheckReport_Detail.FindControl("eCaseDate_Edit");
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Edit.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                    CheckBox cbIsPassed_Edit = (CheckBox)fvHDDCheckReport_Detail.FindControl("cbIsPassed_Edit");
                    Label eIsPassed_Edit = (Label)fvHDDCheckReport_Detail.FindControl("eIsPassed_Edit");
                    if (cbIsPassed_Edit != null)
                    {
                        cbIsPassed_Edit.Checked = (eIsPassed_Edit.Text.Trim() == "True");
                    }

                    Label eModifyDate_Edit = (Label)fvHDDCheckReport_Detail.FindControl("eModifyDate_Edit");
                    eModifyDate_Edit.Text = PF.GetChinsesDate(DateTime.Today);
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eModifyDate_Edit.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eModifyDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                    Label eModifyMan_Edit = (Label)fvHDDCheckReport_Detail.FindControl("eModifyMan_Edit");
                    eModifyMan_Edit.Text = vLoginID;

                    Label eModifyManName_Edit = (Label)fvHDDCheckReport_Detail.FindControl("eModifyManName_Edit");
                    eModifyManName_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    break;
                case FormViewMode.Insert:
                    TextBox eCheckDate_INS = (TextBox)fvHDDCheckReport_Detail.FindControl("eCheckDate_INS");
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCheckDate_INS.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCheckDate_INS.Attributes["onClick"] = vDateScript_Temp;

                    TextBox eCaseDate_INS = (TextBox)fvHDDCheckReport_Detail.FindControl("eCaseDate_INS");
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_INS.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_INS.Attributes["onClick"] = vDateScript_Temp;

                    CheckBox cbIsPassed_INS = (CheckBox)fvHDDCheckReport_Detail.FindControl("cbIsPassed_INS");
                    Label eIsPassed_INS = (Label)fvHDDCheckReport_Detail.FindControl("eIsPassed_INS");
                    if (cbIsPassed_INS != null)
                    {
                        cbIsPassed_INS.Checked = (eIsPassed_INS.Text.Trim() == "True");
                    }

                    Label eBuDate_INS = (Label)fvHDDCheckReport_Detail.FindControl("eBuDate_INS");
                    eBuDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDate_INS.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_INS.Attributes["onClick"] = vDateScript_Temp;

                    Label eBuMan_INS = (Label)fvHDDCheckReport_Detail.FindControl("eBuMan_INS");
                    eBuMan_INS.Text = vLoginID;

                    Label eBuManName_INS = (Label)fvHDDCheckReport_Detail.FindControl("eBuManName_INS");
                    eBuManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    break;
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Edit = (TextBox)fvHDDCheckReport_Detail.FindControl("eDepNo_Edit");
            Label eDepName_Edit = (Label)fvHDDCheckReport_Detail.FindControl("eDepName_Edit");
            string vDepNo_Temp = eDepNo_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Edit.Text = vDepNo_Temp.Trim();
            eDepName_Edit.Text = vDepName_Temp.Trim();
        }

        protected void eDriver_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Edit = (TextBox)fvHDDCheckReport_Detail.FindControl("eDriver_Edit");
            Label eDriverName_Edit = (Label)fvHDDCheckReport_Detail.FindControl("eDriverName_Edit");
            string vDriver_Temp = eDriver_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp.Trim() + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp == "")
            {
                vDriverName_Temp = vDriver_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Edit.Text = vDriver_Temp.Trim();
            eDriverName_Edit.Text = vDriverName_Temp.Trim();
        }

        protected void eInspector_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eInspector_Edit = (TextBox)fvHDDCheckReport_Detail.FindControl("eInspector_Edit");
            Label eInspectorName_Edit = (Label)fvHDDCheckReport_Detail.FindControl("eInspectorName_Edit");
            string vInspector_Temp = eInspector_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vInspector_Temp.Trim() + "' ";
            string vInspectorName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vInspectorName_Temp.Trim() == "")
            {
                vInspectorName_Temp = vInspector_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from EMployee where [Name] = '" + vInspectorName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vInspector_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eInspector_Edit.Text = vInspector_Temp.Trim();
            eInspectorName_Edit.Text = vInspectorName_Temp.Trim();
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_INS = (TextBox)fvHDDCheckReport_Detail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvHDDCheckReport_Detail.FindControl("eDepName_INS");
            string vDepNo_Temp = eDepNo_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_INS.Text = vDepNo_Temp.Trim();
            eDepName_INS.Text = vDepName_Temp.Trim();
        }

        protected void eDriver_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_INS = (TextBox)fvHDDCheckReport_Detail.FindControl("eDriver_INS");
            Label eDriverName_INS = (Label)fvHDDCheckReport_Detail.FindControl("eDriverName_INS");
            string vDriver_Temp = eDriver_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp.Trim() + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp == "")
            {
                vDriverName_Temp = vDriver_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_INS.Text = vDriver_Temp.Trim();
            eDriverName_INS.Text = vDriverName_Temp.Trim();
        }

        protected void eInspector_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eInspector_INS = (TextBox)fvHDDCheckReport_Detail.FindControl("eInspector_INS");
            Label eInspectorName_INS = (Label)fvHDDCheckReport_Detail.FindControl("eInspectorName_INS");
            string vInspector_Temp = eInspector_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vInspector_Temp.Trim() + "' ";
            string vInspectorName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vInspectorName_Temp.Trim() == "")
            {
                vInspectorName_Temp = vInspector_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from EMployee where [Name] = '" + vInspectorName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vInspector_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eInspector_INS.Text = vInspector_Temp.Trim();
            eInspectorName_INS.Text = vInspectorName_Temp.Trim();
        }

        /// <summary>
        /// 刪除查核單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDel_List_Click(object sender, EventArgs e)
        {
            Label eCaseNo_List = (Label)fvHDDCheckReport_Detail.FindControl("eCaseNo_List");
            string vCaseNo_Temp = eCaseNo_List.Text.Trim();
            CheckBox eIsPassed_List = (CheckBox)fvHDDCheckReport_Detail.FindControl("cbIsPassed_List");
            if (eIsPassed_List != null)
            {
                if (eIsPassed_List.Checked)
                {
                    if (vCaseNo_Temp.Trim() != "")
                    {
                        string vSQLStr_Temp = "select isnull(MAX(Items), '0') MaxItems from HDDCheckReport_History where CaseNo = '" + vCaseNo_Temp + "' ";
                        int vMaxItems = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems")) + 1;
                        string vNewCaseNoItems = vCaseNo_Temp + vMaxItems.ToString("D4");
                        vSQLStr_Temp = "insert into HDDCheckReport_History (CaseNoItem, Items, ModifyMode, CaseNo, DepNo, " + Environment.NewLine +
                                       "                                    Driver, DriverName, CaseDate, Car_ID, CheckDate, CheckNote, Remark, " + Environment.NewLine +
                                       "                                    Inspector, BuDate, BuMan, ModifyDate, ModifyMan, IsPassed, AdviceReportNo)" + Environment.NewLine +
                                       "select '" + vNewCaseNoItems.Trim() + "', '" + vMaxItems.ToString("D4") + "', 'DEL', CaseNo, DepNo, " + Environment.NewLine +
                                       "       Driver, DriverName, CaseDate, Car_ID, CheckDate, CheckNote, Remark, Inspector, " + Environment.NewLine +
                                       "       BuDate, BuMan, ModifyDate, ModifyMan, IsPassed, AdviceReportNo " + Environment.NewLine +
                                       "  from HDDCheckReport " + Environment.NewLine +
                                       " where CaseNo = '" + vCaseNo_Temp + "' ";
                        try
                        {
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            sdsHDDCheckReport_Detail.DeleteParameters.Clear();
                            sdsHDDCheckReport_Detail.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                            sdsHDDCheckReport_Detail.Delete();

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
                        Response.Write("alert('請先選擇要刪除的查核單')");
                        Response.Write("</" + "Script>");
                    }
                }
                else
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('已開勸導單，不可刪除！')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// 產生勸導單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPassToAdvice_List_Click(object sender, EventArgs e)
        {
            Label eCaseNo_List = (Label)fvHDDCheckReport_Detail.FindControl("eCaseNo_List");
            string vCaseNo_Temp = eCaseNo_List.Text.Trim();
            string vExecStr = "select AdviceReportNo from HDDCheckReport where CaseNo = '" + vCaseNo_Temp.Trim() + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            if ((vCaseNo_Temp != "") && (PF.GetValue(vConnStr, vExecStr, "AdviceReportNo") == ""))
            {
                Label eCaseDate_List = (Label)fvHDDCheckReport_Detail.FindControl("eCaseDate_List");
                DateTime vCaseDate_Temp = (eCaseDate_List.Text.Trim() != "") ? DateTime.Parse(eCaseDate_List.Text.Trim()) : DateTime.Today;
                Label eDriver_List = (Label)fvHDDCheckReport_Detail.FindControl("eDriver_List");
                string vDriver_Temp = eDriver_List.Text.Trim();
                string vCaseYear = (vCaseDate_Temp.Year - 1911).ToString();
                string vSQLStr_Temp = "select MAX(CaseNo) MAXCaseNo from AdviceReport where CaseNo like '" + vCaseYear + "%' ";
                string vMaxCaseNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MAXCaseNo");
                int vCaseIndex = (vMaxCaseNo != "") ? Int32.Parse(vMaxCaseNo.Replace(vCaseYear, "")) + 1 : 1;
                string vAdviceCaseNo = vCaseYear + vCaseIndex.ToString("D4");
                vSQLStr_Temp = "select ([Name] + ',' + DepNo + ',' + Title) ReturnData from Employee where EmpNo = '" + vDriver_Temp + "' ";
                string vReturnDataStr = PF.GetValue(vConnStr, vSQLStr_Temp, "ReturnData");
                string[] vReturnData = vReturnDataStr.Split(',');
                string vEmpName = vReturnData[0].ToString();
                string vDepNo = vReturnData[1].ToString();
                string vTitle = vReturnData[2].ToString();
                vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vPosition = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                vExecStr = "insert into AdviceReport(CaseNo, CaseDate, EmpNo, EmpName, DepNo, Title, CaseTime, Position, BuMan, BuDate, SourceType) " + Environment.NewLine +
                           " values('" + vAdviceCaseNo + "', '" + vCaseDate_Temp.Year.ToString() + "/" + vCaseDate_Temp.ToString("MM/dd") + "', '" +
                                         vDriver_Temp + "', '" + vEmpName + "', '" + vDepNo + "', '" + vTitle + "', NULL, '" + vPosition + "', '" +
                                         vLoginID + "', '" + DateTime.Today.Year.ToString() + "/" + DateTime.Today.ToString("MM/dd") + "', 'H')";
                try
                {
                    PF.ExecSQL(vConnStr, vExecStr);
                    vExecStr = "update HDDCheckReport set IsPassed = 1, AdviceReportNo = '" + vAdviceCaseNo + "' where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vExecStr);
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
                Response.Write("alert('請先選擇查核硬碟工作報告項目！')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 關閉預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
        }

        protected void sdsHDDCheckReport_Detail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridHDDCheckreport_List.DataBind();
            }
        }

        protected void sdsHDDCheckReport_Detail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridHDDCheckreport_List.DataBind();
            }
        }

        /// <summary>
        /// 插入資料庫資料前的處理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void sdsHDDCheckReport_Detail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eCaseDate_Temp = (TextBox)fvHDDCheckReport_Detail.FindControl("eCaseDate_INS");
            DateTime vCaseDate_Temp = (eCaseDate_Temp.Text.Trim() != "") ? DateTime.Parse(eCaseDate_Temp.Text.Trim()) : DateTime.Today;
            TextBox eCheckDate_Temp = (TextBox)fvHDDCheckReport_Detail.FindControl("eCheckDate_INS");
            DateTime vCheckDate_Temp = (eCheckDate_Temp.Text.Trim() != "") ? DateTime.Parse(eCheckDate_Temp.Text.Trim()) : DateTime.Today;
            Label eBuDate_Temp = (Label)fvHDDCheckReport_Detail.FindControl("eBuDate_INS");
            DateTime vBuDate_Temp = (eBuDate_Temp.Text.Trim() != "") ? DateTime.Parse(eBuDate_Temp.Text.Trim()) : DateTime.Today;
            Label eBuMan_Temp = (Label)fvHDDCheckReport_Detail.FindControl("eBuMan_INS");
            string vBuMan_Temp = (eBuMan_Temp.Text.Trim() != "") ? eBuMan_Temp.Text.Trim() : vLoginID;
            string vCheckYMD = vCheckDate_Temp.Year.ToString() + vCheckDate_Temp.Month.ToString("D2") + vCheckDate_Temp.Day.ToString("D2");
            string vSQLStr_Temp = "select MAX(CaseNo) MaxCaseNo from HDDCheckReport where CaseNo like '" + vCheckYMD + "%' ";
            string vMaxCAseNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxCaseNo");
            int vIndex = (vMaxCAseNo != "") ? Int32.Parse(vMaxCAseNo.Replace(vCheckYMD, "")) + 1 : 1;
            string vNewCaseNo = vCheckYMD + vIndex.ToString("D4");

            e.Command.Parameters["@CaseNo"].Value = vNewCaseNo;
            e.Command.Parameters["@CaseDate"].Value = vCaseDate_Temp;
            e.Command.Parameters["@CheckDate"].Value = vCheckDate_Temp;
            e.Command.Parameters["@BuDate"].Value = vBuDate_Temp;
            e.Command.Parameters["@BuMan"].Value = vBuMan_Temp.Trim();
        }

        protected void sdsHDDCheckReport_Detail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridHDDCheckreport_List.DataBind();
            }
        }

        /// <summary>
        /// 更新資料庫資料前的處理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void sdsHDDCheckReport_Detail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            TextBox eCaseDate_Temp = (TextBox)fvHDDCheckReport_Detail.FindControl("eCaseDate_Edit");
            DateTime vCaseDate_Temp = (eCaseDate_Temp.Text.Trim() != "") ? DateTime.Parse(eCaseDate_Temp.Text.Trim()) : DateTime.Today;
            TextBox eCheckDate_Temp = (TextBox)fvHDDCheckReport_Detail.FindControl("eCheckDate_Edit");
            DateTime vCheckDate_Temp = (eCheckDate_Temp.Text.Trim() != "") ? DateTime.Parse(eCheckDate_Temp.Text.Trim()) : DateTime.Today;
            Label eModifyDate_Temp = (Label)fvHDDCheckReport_Detail.FindControl("eModifyDate_Edit");
            Label eModifyMan_Temp = (Label)fvHDDCheckReport_Detail.FindControl("eModifyMan_Edit");
            DateTime vModifyDate_Temp = (eModifyDate_Temp.Text.Trim() != "") ? DateTime.Parse(eModifyDate_Temp.Text.Trim()) : DateTime.Today;
            string vModifyMan_Temp = (eModifyMan_Temp.Text.Trim() != "") ? eModifyMan_Temp.Text.Trim() : vLoginID;

            e.Command.Parameters["@CaseDate"].Value = vCaseDate_Temp;
            e.Command.Parameters["@CheckDate"].Value = vCheckDate_Temp;
            e.Command.Parameters["@ModifyDate"].Value = vModifyDate_Temp;
            e.Command.Parameters["@ModifyMan"].Value = vModifyMan_Temp.Trim();
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
            string vFileName = "查核硬碟工作報告";
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
                                      (drExcel.GetName(i) == "DepName") ? "站別" :
                                      (drExcel.GetName(i) == "Driver") ? "" :
                                      (drExcel.GetName(i) == "DriverName") ? "駕駛員姓名" :
                                      (drExcel.GetName(i) == "CaseDate") ? "事件日" :
                                      (drExcel.GetName(i) == "Car_ID") ? "車號" :
                                      (drExcel.GetName(i) == "CheckNote") ? "查核項目" :
                                      (drExcel.GetName(i) == "Remark") ? "備註" :
                                      (drExcel.GetName(i) == "InspectorName") ? "查核人" :
                                      (drExcel.GetName(i) == "BuDate") ? "建檔日期" :
                                      (drExcel.GetName(i) == "BuManName") ? "建檔人" :
                                      (drExcel.GetName(i) == "ModifyDate") ? "異動日期" :
                                      (drExcel.GetName(i) == "ModifyManName") ? "異動人" :
                                      (drExcel.GetName(i) == "IsPassed") ? "已過勸導單" :
                                      (drExcel.GetName(i) == "AdviceReportNo") ? "勸導單號" :
                                      (drExcel.GetName(i) == "CheckDate") ? "查核日" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i) == "CaseDate") ||
                                 (drExcel.GetName(i) == "CheckDate") ||
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
                            string vCaseDateStr = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "自 " + eCaseDate_S_Search.Text.Trim() + " 起至 " + eCaseDate_E_Search.Text.Trim() + " 止" :
                                                  ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? eCaseDate_S_Search.Text.Trim() :
                                                  ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? eCaseDate_E_Search.Text.Trim() : "";
                            string vDepNoStr = ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? "自 " + eDepNo_S_Search.Text.Trim() + " 站起至 " + eDepNo_E_Search.Text.Trim() + " 站止" :
                                               ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? eDepNo_S_Search.Text.Trim() + " 站" :
                                               ((eDepNo_S_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? eDepNo_E_Search.Text.Trim() + " 站" : "全部";
                            string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vCheckDateStr = ((eCheckDate_S_Search.Text.Trim() != "") && (eCheckDate_E_Search.Text.Trim() != "")) ? "自 " + eCheckDate_S_Search.Text.Trim() + " 起至 " + eCheckDate_E_Search.Text.Trim() + " 止" :
                                                   ((eCheckDate_S_Search.Text.Trim() != "") && (eCheckDate_E_Search.Text.Trim() == "")) ? eCheckDate_S_Search.Text.Trim() :
                                                   ((eCheckDate_S_Search.Text.Trim() == "") && (eCheckDate_E_Search.Text.Trim() != "")) ? eCheckDate_E_Search.Text.Trim() : "";
                            string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "全部";
                            string vInspectorNoStr = (eInspector_Search.Text.Trim() != "") ? eInspector_Search.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_查核硬碟工作報告" + Environment.NewLine +
                                                 "HDDCheckReport.aspx" + Environment.NewLine +
                                                 "事件日期：" + vCaseDateStr + Environment.NewLine +
                                                 "站別：" + vDepNoStr + Environment.NewLine +
                                                 "駕駛員編號：" + vDriverStr + Environment.NewLine +
                                                 "查核日期：" + vCheckDateStr + Environment.NewLine +
                                                 "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                                 "查核人：" + vInspectorNoStr;
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