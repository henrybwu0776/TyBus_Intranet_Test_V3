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
    public partial class RunsheetDailyCheck : System.Web.UI.Page
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
                    string vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDateS_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDateS_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDateE_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDateE_Search.Attributes["onClick"] = vCaseDateScript;

                    //查核日期
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDateS_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDateS_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuDateE_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDateE_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        plSearch.Visible = true;
                        plShowData.Visible = true;
                        plPrint.Visible = false;
                        bbPrint.Visible = false;
                        eBuMan_Search.Text = vLoginID;
                        eBuManName_Search.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
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
        /// 取得查詢語法
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            //開單日期
            DateTime vBuDateS = (eBuDateS_Search.Text.Trim() != "") ?
                                DateTime.Parse(eBuDateS_Search.Text.Trim()) :
                                DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C"));
            string vBuDateS_Str = vBuDateS.Year.ToString() + "/" + vBuDateS.ToString("MM/dd");
            DateTime vBuDateE = (eBuDateE_Search.Text.Trim() != "") ?
                                DateTime.Parse(eBuDateE_Search.Text.Trim()) :
                                DateTime.Parse(PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C"));
            string vBuDateE_Str = vBuDateE.Year.ToString() + "/" + vBuDateE.ToString("MM/dd");
            string vWStr_BuDate = ((eBuDateS_Search.Text.Trim() != "") && (eBuDateE_Search.Text.Trim() != "")) ? "   and r.BuDate between '" + vBuDateS_Str + "' and '" + vBuDateE_Str + "' " + Environment.NewLine :
                                  ((eBuDateS_Search.Text.Trim() != "") && (eBuDateE_Search.Text.Trim() == "")) ? "   and r.BuDate = '" + vBuDateS_Str + "' " + Environment.NewLine :
                                  ((eBuDateS_Search.Text.Trim() == "") && (eBuDateE_Search.Text.Trim() != "")) ? "   and r.BuDate = '" + vBuDateE_Str + "' " + Environment.NewLine : "";
            //憑單及碼表日期
            DateTime vCaseDateS = (eCaseDateS_Search.Text.Trim() != "") ?
                                  DateTime.Parse(eCaseDateS_Search.Text.Trim()) :
                                  DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C"));
            string vCaseDateS_Str = vCaseDateS.Year.ToString() + "/" + vCaseDateS.ToString("MM/dd");
            DateTime vCaseDateE = (eCaseDateE_Search.Text.Trim() != "") ?
                                  DateTime.Parse(eCaseDateE_Search.Text.Trim()) :
                                  DateTime.Parse(PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C"));
            string vCaseDateE_Str = vCaseDateE.Year.ToString() + "/" + vCaseDateE.ToString("MM/dd");
            string vWStr_CaseDate = ((eCaseDateS_Search.Text.Trim() != "") && (eCaseDateE_Search.Text.Trim() != "")) ? "   and r.CaseDate between '" + vCaseDateS_Str + "' and '" + vCaseDateE_Str + "' " + Environment.NewLine :
                                  ((eCaseDateS_Search.Text.Trim() != "") && (eCaseDateE_Search.Text.Trim() == "")) ? "   and r.CaseDate = '" + vCaseDateS_Str + "' " + Environment.NewLine :
                                  ((eCaseDateS_Search.Text.Trim() == "") && (eCaseDateE_Search.Text.Trim() != "")) ? "   and r.CaseDate = '" + vCaseDateE_Str + "' " + Environment.NewLine : "";
            string vWStr_CarID = (eCarID_Search.Text.Trim() != "") ? "   and r.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and r.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? "   and r.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuMan = (eBuMan_Search.Text.Trim() != "") ? "   and r.BuMan = '" + eBuMan_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vReturnStr = "select CaseNo, CaseDate, DepNo, (select [Name] from Department where DepNo = r.DepNo) DepName, " + Environment.NewLine +
                                "       Car_ID, Driver, DriverName, Remark, " + Environment.NewLine +
                                "       Inspector, (select [Name] from Employee where EmpNo = r.Inspector) InspectorName, " + Environment.NewLine +
                                "       BuDate, BuMan, (select [Name] from Employee where EmpNo = r.BuMan) BuManName, " + Environment.NewLine +
                                "       ModifyDate, ModifyMan, (select [Name] from Employee where EmpNo = r.ModifyMan) ModifyManName " + Environment.NewLine +
                                "  from RunsheetDailyCheck r " + Environment.NewLine +
                                " where isnull(CaseNo, '') <> '' " + Environment.NewLine +
                                vWStr_BuDate + vWStr_CarID + vWStr_CaseDate + vWStr_DepNo + vWStr_Driver + vWStr_BuMan +
                                " order by CaseDate, DepNo, Car_ID";
            return vReturnStr;
        }

        /// <summary>
        /// 開啟資料庫
        /// </summary>
        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            sdsRunsheetDailyCheck_List.SelectCommand = "";
            sdsRunsheetDailyCheck_List.SelectCommand = vSelStr;
            gridRunsheetDailyCheck_List.DataSourceID = "";
            gridRunsheetDailyCheck_List.DataSourceID = "sdsRunsheetDailyCheck_List";
            gridRunsheetDailyCheck_List.DataBind();
            bbPrint.Visible = (gridRunsheetDailyCheck_List.Rows.Count > 0);
            plShowData.Visible = true;
        }

        protected void eBuMan_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vBuMan_Temp = eBuMan_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vBuMan_Temp + "' ";
            string vBuManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vBuManName_Temp.Trim() == "")
            {
                vBuManName_Temp = vBuMan_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vBuManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vBuMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eBuMan_Search.Text = vBuMan_Temp.Trim();
            eBuManName_Search.Text = vBuManName_Temp.Trim();
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Search.Text = vDepNo_Temp.Trim();
            eDepName_Search.Text = vDepName_Temp.Trim();
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
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "eMPnO");
            }
            eDriver_Search.Text = vDriver_Temp.Trim();
            eDriverName_Search.Text = vDriverName_Temp.Trim();
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
                DataTable dtPrintPoint = new DataTable("RunsheetDailyCheckP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("RunsheetDailyCheckP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\RunsheetDailyCheckP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "每日憑單及碼表抽查表"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;

                    string vBuDateStr = ((eBuDateS_Search.Text.Trim() != "") && (eBuDateE_Search.Text.Trim() != "")) ? "自 " + eBuDateS_Search.Text.Trim() + " 起至 " + eBuDateE_Search.Text.Trim() + " 止" :
                                        ((eBuDateS_Search.Text.Trim() != "") && (eBuDateE_Search.Text.Trim() == "")) ? eBuDateS_Search.Text.Trim() :
                                        ((eBuDateS_Search.Text.Trim() == "") && (eBuDateE_Search.Text.Trim() != "")) ? eBuDateE_Search.Text.Trim() : "不分日期";
                    string vCaseDateStr = ((eCaseDateS_Search.Text.Trim() != "") && (eCaseDateE_Search.Text.Trim() != "")) ? "自 " + eCaseDateS_Search.Text.Trim() + " 起至 " + eCaseDateE_Search.Text.Trim() + " 止" :
                                          ((eCaseDateS_Search.Text.Trim() != "") && (eCaseDateE_Search.Text.Trim() == "")) ? eCaseDateS_Search.Text.Trim() :
                                          ((eCaseDateS_Search.Text.Trim() == "") && (eCaseDateE_Search.Text.Trim() != "")) ? eCaseDateE_Search.Text.Trim() : "不分日期";
                    string vBumanStr = (eBuMan_Search.Text.Trim() != "") ? eBuMan_Search.Text.Trim() : "全部";
                    string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "全部";
                    string vDepNoStr = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "全部";
                    string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                    string vRecordNote = "預覽報表_每日憑單及碼表抽查表" + Environment.NewLine +
                                         "RunsheetDailyCheck.aspx" + Environment.NewLine +
                                         "開單日期：" + vBuDateStr + Environment.NewLine +
                                         "憑單或碼表日期：" + vCaseDateStr + Environment.NewLine +
                                         "開單人員：" + vBumanStr + Environment.NewLine +
                                         "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                         "站別：" + vDepNoStr + Environment.NewLine +
                                         "駕駛員：" + vDriverStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        protected void fvRunsheetDailyCheck_Detail_DataBound(object sender, EventArgs e)
        {
            string vDateURL = "";
            string vDateScript = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvRunsheetDailyCheck_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    Label eModifyDate_Edit = (Label)fvRunsheetDailyCheck_Detail.FindControl("eModifyDate_Edit");
                    eModifyDate_Edit.Text = PF.GetChinsesDate(DateTime.Today);
                    Label eModifyMan_Edit = (Label)fvRunsheetDailyCheck_Detail.FindControl("eModifyMan_Edit");
                    eModifyMan_Edit.Text = vLoginID;
                    Label eModifyManName_Edit = (Label)fvRunsheetDailyCheck_Detail.FindControl("eModifyManName_Edit");
                    eModifyManName_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    TextBox eCaseDate_Edit = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eCaseDate_Edit");
                    eCaseDate_Edit.Text = PF.GetChinsesDate(DateTime.Today.AddDays(-1));
                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Edit.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Edit.Attributes["onClick"] = vDateScript;
                    TextBox eInspector_Edit = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eInspector_Edit");
                    eInspector_Edit.Text = vLoginID;
                    Label eInspectorName_Edit = (Label)fvRunsheetDailyCheck_Detail.FindControl("eInspectorName_Edit");
                    eInspectorName_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    break;
                case FormViewMode.Insert:
                    Label eBuDate_INS = (Label)fvRunsheetDailyCheck_Detail.FindControl("eBuDate_INS");
                    eBuDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                    Label eBuMan_INS = (Label)fvRunsheetDailyCheck_Detail.FindControl("eBuMan_INS");
                    eBuMan_INS.Text = vLoginID;
                    Label eBuManName_INS = (Label)fvRunsheetDailyCheck_Detail.FindControl("eBuManNAme_INS");
                    eBuManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    TextBox eCaseDate_INS = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eCaseDate_INS");
                    eCaseDate_INS.Text = PF.GetChinsesDate(DateTime.Today.AddDays(-1));
                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_INS.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_INS.Attributes["onClick"] = vDateScript;
                    TextBox eInspector_INS = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eInspector_INS");
                    eInspector_INS.Text = vLoginID;
                    Label eInspectorName_INS = (Label)fvRunsheetDailyCheck_Detail.FindControl("eInspectorName_INS");
                    eInspectorName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    break;
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Temp = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eDepNo_Edit");
            Label eDepName_Temp = (Label)fvRunsheetDailyCheck_Detail.FindControl("eDepName_Edit");
            string vDepNo_Str = eDepNo_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Str + "' ";
            string vDepName_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Str == "")
            {
                vDepName_Str = vDepNo_Str.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Str.Trim() + "' ";
                vDepNo_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Temp.Text = vDepNo_Str.Trim();
            eDepName_Temp.Text = vDepName_Str.Trim();
        }

        protected void eDriver_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Temp = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eDriver_Edit");
            Label eDriverName_Temp = (Label)fvRunsheetDailyCheck_Detail.FindControl("eDrivername_Edit");
            string vDriver_Str = eDriver_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Str.Trim() + "' ";
            string vDriverName_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Str == "")
            {
                vDriverName_Str = vDriver_Str.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Str.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Temp.Text = vDriver_Str.Trim();
            eDriverName_Temp.Text = vDriverName_Str.Trim();
        }

        protected void eInspector_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eInspector_Temp = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eInspector_Edit");
            Label eInspectorName_Temp = (Label)fvRunsheetDailyCheck_Detail.FindControl("eInspectorname_Edit");
            string vInspector_Str = eInspector_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vInspector_Str.Trim() + "' ";
            string vInspectorName_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vInspectorName_Str == "")
            {
                vInspectorName_Str = vInspector_Str.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vInspectorName_Str.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vInspector_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eInspector_Temp.Text = vInspector_Str.Trim();
            eInspectorName_Temp.Text = vInspectorName_Str.Trim();
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Temp = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eDepNo_INS");
            Label eDepName_Temp = (Label)fvRunsheetDailyCheck_Detail.FindControl("eDepName_INS");
            string vDepNo_Str = eDepNo_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Str + "' ";
            string vDepName_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Str == "")
            {
                vDepName_Str = vDepNo_Str.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Str.Trim() + "' ";
                vDepNo_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Temp.Text = vDepNo_Str.Trim();
            eDepName_Temp.Text = vDepName_Str.Trim();
        }

        protected void eDriver_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Temp = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eDriver_INS");
            Label eDriverName_Temp = (Label)fvRunsheetDailyCheck_Detail.FindControl("eDrivername_INS");
            string vDriver_Str = eDriver_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Str.Trim() + "' ";
            string vDriverName_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Str == "")
            {
                vDriverName_Str = vDriver_Str.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Str.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Temp.Text = vDriver_Str.Trim();
            eDriverName_Temp.Text = vDriverName_Str.Trim();
        }

        protected void eInspector_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eInspector_Temp = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eInspector_INS");
            Label eInspectorName_Temp = (Label)fvRunsheetDailyCheck_Detail.FindControl("eInspectorname_INS");
            string vInspector_Str = eInspector_Temp.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vInspector_Str.Trim() + "' ";
            string vInspectorName_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vInspectorName_Str == "")
            {
                vInspectorName_Str = vInspector_Str.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vInspectorName_Str.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vInspector_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eInspector_Temp.Text = vInspector_Str.Trim();
            eInspectorName_Temp.Text = vInspectorName_Str.Trim();
        }

        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            Label eCaseNo_Temp = (Label)fvRunsheetDailyCheck_Detail.FindControl("eCaseNo_List");
            string vCaseNo_Temp = eCaseNo_Temp.Text.Trim();
            if (vCaseNo_Temp != "")
            {
                string vSQLStr_Temp = "select max(Items) MaxIndex from RunsheetDailyCheck_History where CaseNo = '" + vCaseNo_Temp.Trim() + "' ";
                string vMaxIndexStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex").Trim();
                int vNewIndex = (vMaxIndexStr != "") ? Int32.Parse(vMaxIndexStr) + 1 : 1;
                vSQLStr_Temp = "insert into RunsheetDailyCheck_History " + Environment.NewLine +
                               "       (CaseNoItem, Items, ModifyMode, CaseNo, CaseDate, DepNo, Car_ID, Driver, DriverName, " + Environment.NewLine +
                               "        Remark, Inspector, Budate, BuMan, ModifyDate, ModifyMan)" + Environment.NewLine +
                               "select '" + vCaseNo_Temp + vNewIndex.ToString("D4") + "', '" + vNewIndex.ToString("D4") + "', 'DEL', " + Environment.NewLine +
                               "       CaseNo, CaseDate, DepNo, Car_ID, Driver, DriverName, Remark, Inspector, Budate, BuMan, ModifyDate, ModifyMan " + Environment.NewLine +
                               "  from RunsheetDailyCheck " + Environment.NewLine +
                               " where CaseNo = '" + vCaseNo_Temp.Trim() + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    sdsRunsheetDailyCheck_Detail.DeleteParameters.Clear();
                    sdsRunsheetDailyCheck_Detail.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp));
                    sdsRunsheetDailyCheck_Detail.Delete();
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
        /// 結束預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
        }

        protected void sdsRunsheetDailyCheck_Detail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridRunsheetDailyCheck_List.DataBind();
            }
        }

        protected void sdsRunsheetDailyCheck_Detail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridRunsheetDailyCheck_List.DataBind();
            }
        }

        protected void sdsRunsheetDailyCheck_Detail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eBudate_INS = (Label)fvRunsheetDailyCheck_Detail.FindControl("eBudate_INS");
            if (eBudate_INS != null)
            {
                DateTime vBuDate_Temp = (eBudate_INS.Text.Trim() != "") ? DateTime.Parse(eBudate_INS.Text.Trim()) : DateTime.Today;
                e.Command.Parameters["@BuDate"].Value = vBuDate_Temp;
                string vCaseYMD = vBuDate_Temp.Year.ToString() + vBuDate_Temp.Month.ToString("D2") + vBuDate_Temp.Day.ToString("D2");
                string vSQLStr_Temp = "select MAX(CaseNo) MaxCaseNo from RunsheetDailyCheck where CaseNo like '" + vCaseYMD.Trim() + "%' ";
                string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxCaseNo").Replace(vCaseYMD, "").Trim();
                int vNewIndex = (vMaxIndex != "") ? Int32.Parse(vMaxIndex) + 1 : 1;
                string vNewCaseNo = vCaseYMD + vNewIndex.ToString("D4");
                e.Command.Parameters["@CaseNo"].Value = vNewCaseNo;
                TextBox eCaseDate_INS = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eCaseDate_INS");
                if ((eCaseDate_INS != null) && (eCaseDate_INS.Text.Trim() != ""))
                {
                    e.Command.Parameters["@CaseDate"].Value = DateTime.Parse(eCaseDate_INS.Text.Trim());
                }
            }
        }

        protected void sdsRunsheetDailyCheck_Detail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridRunsheetDailyCheck_List.DataBind();
            }
        }

        protected void sdsRunsheetDailyCheck_Detail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eCaseNo_Edit = (Label)fvRunsheetDailyCheck_Detail.FindControl("eCaseNo_Edit");
            if ((eCaseNo_Edit != null) && (eCaseNo_Edit.Text.Trim() != ""))
            {
                string vOldCaseNo = eCaseNo_Edit.Text.Trim();
                string vSQLStr_Temp = "select max(Items) MaxIndex from RunsheetDailyCheck_History where CaseNo = '" + vOldCaseNo.Trim() + "' ";
                string vMaxIndexStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex").Trim();
                int vNewIndex = (vMaxIndexStr != "") ? Int32.Parse(vMaxIndexStr) + 1 : 1;
                vSQLStr_Temp = "insert into RunsheetDailyCheck_History " + Environment.NewLine +
                               "       (CaseNoItem, Items, ModifyMode, CaseNo, CaseDate, DepNo, Car_ID, Driver, DriverName, " + Environment.NewLine +
                               "        Remark, Inspector, Budate, BuMan, ModifyDate, ModifyMan)" + Environment.NewLine +
                               "select '" + vOldCaseNo + vNewIndex.ToString("D4") + "', '" + vNewIndex.ToString("D4") + "', 'EDIT', " + Environment.NewLine +
                               "       CaseNo, CaseDate, DepNo, Car_ID, Driver, DriverName, Remark, Inspector, Budate, BuMan, ModifyDate, ModifyMan " + Environment.NewLine +
                               "  from RunsheetDailyCheck " + Environment.NewLine +
                               " where CaseNo = '" + vOldCaseNo.Trim() + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
            Label eModifyDate_Edit = (Label)fvRunsheetDailyCheck_Detail.FindControl("eModifyDate_Edit");
            if (eModifyDate_Edit != null)
            {
                DateTime vModifyDate_Temp = (eModifyDate_Edit.Text.Trim() != "") ? DateTime.Parse(eModifyDate_Edit.Text.Trim()) : DateTime.Today;
                e.Command.Parameters["@ModifyDate"].Value = vModifyDate_Temp;
            }
            TextBox eCaseDate_Edit = (TextBox)fvRunsheetDailyCheck_Detail.FindControl("eCaseDate_Edit");
            if ((eCaseDate_Edit != null) && (eCaseDate_Edit.Text.Trim() != ""))
            {
                DateTime vCaseDate_Temp = DateTime.Parse(eCaseDate_Edit.Text.Trim());
                e.Command.Parameters["@CaseDate"].Value = vCaseDate_Temp;
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

            string vHeaderText = "";
            int vLinesNo = 0;
            string vSelStr = GetSelectStr();
            string vFileName = "每日憑單及碼表抽查清冊";
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
                                      (drExcel.GetName(i) == "CaseDate") ? "憑單及碼表日期" :
                                      (drExcel.GetName(i) == "DepNo") ? "" :
                                      (drExcel.GetName(i) == "DepName") ? "站別" :
                                      (drExcel.GetName(i) == "Car_ID") ? "車號" :
                                      (drExcel.GetName(i) == "Driver") ? "" :
                                      (drExcel.GetName(i) == "DriverName") ? "駕駛員姓名" :
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
                            if (((drExcel.GetName(i) == "CaseDate") ||
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
                            string vBuDateStr = ((eBuDateS_Search.Text.Trim() != "") && (eBuDateE_Search.Text.Trim() != "")) ? "自 " + eBuDateS_Search.Text.Trim() + " 起至 " + eBuDateE_Search.Text.Trim() + " 止" :
                                                ((eBuDateS_Search.Text.Trim() != "") && (eBuDateE_Search.Text.Trim() == "")) ? eBuDateS_Search.Text.Trim() :
                                                ((eBuDateS_Search.Text.Trim() == "") && (eBuDateE_Search.Text.Trim() != "")) ? eBuDateE_Search.Text.Trim() : "不分日期";
                            string vCaseDateStr = ((eCaseDateS_Search.Text.Trim() != "") && (eCaseDateE_Search.Text.Trim() != "")) ? "自 " + eCaseDateS_Search.Text.Trim() + " 起至 " + eCaseDateE_Search.Text.Trim() + " 止" :
                                                  ((eCaseDateS_Search.Text.Trim() != "") && (eCaseDateE_Search.Text.Trim() == "")) ? eCaseDateS_Search.Text.Trim() :
                                                  ((eCaseDateS_Search.Text.Trim() == "") && (eCaseDateE_Search.Text.Trim() != "")) ? eCaseDateE_Search.Text.Trim() : "不分日期";
                            string vBumanStr = (eBuMan_Search.Text.Trim() != "") ? eBuMan_Search.Text.Trim() : "全部";
                            string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "全部";
                            string vDepNoStr = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "全部";
                            string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "全部";
                            string vRecordNote = "匯出檔案_每日憑單及碼表抽查表" + Environment.NewLine +
                                                 "RunsheetDailyCheck.aspx" + Environment.NewLine +
                                                 "開單日期：" + vBuDateStr + Environment.NewLine +
                                                 "憑單或碼表日期：" + vCaseDateStr + Environment.NewLine +
                                                 "開單人員：" + vBumanStr + Environment.NewLine +
                                                 "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                                 "站別：" + vDepNoStr + Environment.NewLine +
                                                 "駕駛員：" + vDriverStr;
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