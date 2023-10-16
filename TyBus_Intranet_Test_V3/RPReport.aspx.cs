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
    public partial class RPReport : System.Web.UI.Page
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

                    //呈報日期
                    string vCountDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eAssignDate_Search_S.ClientID;
                    string vCountDateScript = "window.open('" + vCountDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eAssignDate_Search_S.Attributes["onClick"] = vCountDateScript;

                    vCountDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eAssignDate_Search_E.ClientID;
                    vCountDateScript = "window.open('" + vCountDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eAssignDate_Search_E.Attributes["onClick"] = vCountDateScript;

                    //事件日期
                    vCountDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Search_S.ClientID;
                    vCountDateScript = "window.open('" + vCountDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Search_S.Attributes["onClick"] = vCountDateScript;

                    vCountDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Search_E.ClientID;
                    vCountDateScript = "window.open('" + vCountDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Search_E.Attributes["onClick"] = vCountDateScript;

                    if (!IsPostBack)
                    {
                        string vSQLStr_CaseComeFrom = "";
                        string vSQLStr_CaseType = "";
                        using (SqlConnection connCaseComeFrom = new SqlConnection(vConnStr))
                        {
                            vSQLStr_CaseComeFrom = "SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt " + Environment.NewLine +
                                           " UNION ALL " + Environment.NewLine +
                                           "SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '員工獎懲呈報單  RPReport        CaseComeFrom') ORDER BY ClassNo ";
                            SqlCommand cmdCaseComeFrom = new SqlCommand(vSQLStr_CaseComeFrom, connCaseComeFrom);
                            connCaseComeFrom.Open();

                            SqlDataReader drCaseComeFrom = cmdCaseComeFrom.ExecuteReader();
                            ddlCaseComeFrom_Search.Items.Clear();
                            while (drCaseComeFrom.Read())
                            {
                                ddlCaseComeFrom_Search.Items.Add(new ListItem(drCaseComeFrom["ClassTxt"].ToString().Trim(), drCaseComeFrom["ClassNo"].ToString().Trim()));
                            }
                            drCaseComeFrom.Close();
                            cmdCaseComeFrom.Cancel();
                        }
                        using (SqlConnection connCaseType = new SqlConnection(vConnStr))
                        {
                            vSQLStr_CaseType = "SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt " + Environment.NewLine +
                                               " UNION ALL " + Environment.NewLine +
                                               "SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '員工獎懲呈報單  RPReport        CaseType') ORDER BY ClassNo ";
                            SqlCommand cmdCaseType = new SqlCommand(vSQLStr_CaseType, connCaseType);
                            connCaseType.Open();

                            SqlDataReader drCaseType = cmdCaseType.ExecuteReader();
                            ddlCaseType_Search.Items.Clear();
                            while (drCaseType.Read())
                            {
                                ddlCaseType_Search.Items.Add(new ListItem(drCaseType["ClassTxt"].ToString().Trim(), drCaseType["ClassNo"].ToString().Trim()));
                            }
                            drCaseType.Close();
                            cmdCaseType.Cancel();
                        }
                        //呈報日期
                        eAssignDate_Search_S.Text = PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C");
                        eAssignDate_Search_E.Text = PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C");
                        //事件日期
                        //eCaseDate_Search_S.Text = PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C");
                        //eCaseDate_Search_E.Text = PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C");
                        ddlCaseComeFrom_Search.SelectedIndex = 0;
                        eCaseComeFrom_Search.Text = "";
                        ddlCaseType_Search.SelectedIndex = 0;
                        eCaseType_Search.Text = "";
                        plReport.Visible = false;
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
        /// 取回列印用的查詢字串
        /// </summary>
        /// <param name="fCaseNo">呈報單號碼</param>
        /// <returns>回傳字串</returns>
        private string GetSelectStr_Print(string fCaseNo)
        {
            string vReturnStr = "";
            vReturnStr = "SELECT CaseNo, CaseComeFrom, (select ClassTxt from DBDICB where ClassNo = r.CaseComeFrom and FKey = '員工獎懲呈報單  RPReport        CaseComeFrom') CaseComeFrom_C, " + Environment.NewLine +
                         "       CaseType, (SELECT ClassTxt FROM DBDICB WHERE(FKey = '員工獎懲呈報單  RPReport        CaseType') AND(ClassNo = r.CaseType)) AS CaseType_C, " + Environment.NewLine +
                         "       DepNo, (select[Name] from Department where DepNo = r.DepNo) DepName, " + Environment.NewLine +
                         "       EmpNo, EmpName, Title, (SELECT CLASSTXT FROM DBDICB WHERE(CLASSNO = r.Title) AND(FKEY = '人事資料檔      EMPLOYEE        TITLE')) AS Title_C, " + Environment.NewLine +
                         "       CaseDate, Car_ID, Position, CaseNote, AccordingTerms, Remark, AssignDate, " + Environment.NewLine +
                         "       AssignDepNo, (select[Name] from Department where DepNo = r.AssignDepNo) AssignDepName, " + Environment.NewLine +
                         "       AssignMan,(select[Name] from Employee where EmpNo = r.AssignMan) AssignManName, " + Environment.NewLine +
                         "       BuDate, BuMan, (select[Name] from Employee where EmpNo = r.BuMan) BuManName, " + Environment.NewLine +
                         "       ModifyDate, ModifyMan, (select[Name] from Employee where EmpNo = r.ModifyMan) ModifyManName, CustomServiceNo, " + Environment.NewLine +
                         "       case when isnull(GiveBounds, 0) = 1 then '頒發獎金 ' + cast(BoundsAmount as nvarchar) + '元。' else '' end + " + Environment.NewLine +
                         "       case when isnull(AskExact, 0) = 1 then '賠償扣款' + cast(ExactAmount as nvarchar) + '元。' else '' end + " + Environment.NewLine +
                         "       case when isnull(Demotion, 0) = 1 then '職務降調。' else '' end + " + Environment.NewLine +
                         "       case when isnull(Dismissal, 0) = 1 then '予以解職。' else '' end + " + Environment.NewLine +
                         "       case when isnull(Promotion, 0) = 1 then '職務晉升。'  else '' end + " + Environment.NewLine +
                         "       case when isnull(Advice, 0) = 1 then '予以勸導。' else '' end + " + Environment.NewLine +
                         "       case when isnull(Admonition, 0) = 1 then '口頭警告。' else '' end + " + Environment.NewLine +
                         "       case when isnull(Reprimand, 0) = 1 then '申誡' + cast(ReprimandCount as nvarchar) + '次。' else '' end + " + Environment.NewLine +
                         "       case when isnull(Demerit, 0) = 1 then '小過' + cast(DemeritCount as nvarchar) + '次。' else '' end + " + Environment.NewLine +
                         "       case when isnull(MajorDemerit, 0) = 1 then '大過' + cast(MajorDemeritCount as nvarchar) + '次。' else '' end + " + Environment.NewLine +
                         "       case when isnull(Commendation, 0) = 1 then '嘉獎' + cast(CommendationCount as nvarchar) + '次。' else '' end + " + Environment.NewLine +
                         "       case when isnull(MeritCitation, 0) = 1 then '小功' + cast(MeritCitationCount as nvarchar) + '次。' else '' end + " + Environment.NewLine +
                         "       case when isnull(MajorMeritCitation, 0) = 1 then '大功' + cast(MajorMeritCitationCount as nvarchar) + '次。' else '' end + " + Environment.NewLine +
                         "       isnull(Review, '') as Review_C " + Environment.NewLine +
                         "  FROM RP_Report r " + Environment.NewLine +
                         " where CaseNo = '" + fCaseNo + "' ";
            return vReturnStr;
        }

        /// <summary>
        /// 取回資料查詢字串
        /// </summary>
        /// <returns>回傳字串</returns>
        private string GetSelectStr()
        {
            string vReturnStr = "";
            DateTime vAssignDate_S_Temp = (eAssignDate_Search_S.Text.Trim() != "") ?
                                          DateTime.Parse(eAssignDate_Search_S.Text.Trim()) :
                                          DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C"));
            DateTime vAssignDate_E_Temp = (eAssignDate_Search_E.Text.Trim() != "") ?
                                          DateTime.Parse(eAssignDate_Search_E.Text.Trim()) :
                                          DateTime.Parse(PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C"));
            DateTime vCaseDate_S_Temp = (eCaseDate_Search_S.Text.Trim() != "") ?
                                        DateTime.Parse(eCaseDate_Search_S.Text.Trim()) :
                                        DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C"));
            DateTime vCaseDate_E_Temp = (eCaseDate_Search_E.Text.Trim() != "") ?
                                        DateTime.Parse(eCaseDate_Search_E.Text.Trim()) :
                                        DateTime.Parse(PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C"));
            string vWStr_AssignDate = ((eAssignDate_Search_S.Text.Trim() != "") && (eAssignDate_Search_E.Text.Trim() != "")) ?
                                      "   and AssignDate between '" + vAssignDate_S_Temp.Year.ToString() + "/" + vAssignDate_S_Temp.ToString("MM/dd") + "' and '" + vAssignDate_E_Temp.Year.ToString() + "/" + vAssignDate_E_Temp.ToString("MM/dd") + "' " + Environment.NewLine :
                                      ((eAssignDate_Search_S.Text.Trim() != "") && (eAssignDate_Search_E.Text.Trim() == "")) ?
                                      "   and AssignDate = '" + vAssignDate_S_Temp.Year.ToString() + "/" + vAssignDate_S_Temp.ToString("MM/dd") + "' " + Environment.NewLine :
                                      ((eAssignDate_Search_S.Text.Trim() == "") && (eAssignDate_Search_E.Text.Trim() != "")) ?
                                      "   and AssignDate = '" + vAssignDate_E_Temp.Year.ToString() + "/" + vAssignDate_E_Temp.ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_CaseDate = ((eCaseDate_Search_S.Text.Trim() != "") && (eCaseDate_Search_E.Text.Trim() != "")) ?
                                    "   and CaseDate between '" + vCaseDate_S_Temp.Year.ToString() + "/" + vCaseDate_S_Temp.ToString("MM/dd") + "' and '" + vCaseDate_E_Temp.Year.ToString() + "/" + vCaseDate_E_Temp.ToString("MM/dd") + "' " + Environment.NewLine :
                                    ((eCaseDate_Search_S.Text.Trim() != "") && (eCaseDate_Search_E.Text.Trim() == "")) ?
                                    "   and CaseDate = '" + vCaseDate_S_Temp.Year.ToString() + "/" + vCaseDate_S_Temp.ToString("MM/dd") + "' " + Environment.NewLine :
                                    ((eCaseDate_Search_S.Text.Trim() == "") && (eCaseDate_Search_E.Text.Trim() != "")) ?
                                    "   and CaseDate = '" + vCaseDate_E_Temp.Year.ToString() + "/" + vCaseDate_E_Temp.ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and EmpNo = '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CaseComeFrom = (eCaseComeFrom_Search.Text.Trim() != "") ? "   and CaseComeFrom = '" + eCaseComeFrom_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CaseType = (eCaseType_Search.Text.Trim() != "") ? "   and CaseType = '" + eCaseType_Search.Text.Trim() + "' " + Environment.NewLine : "";
            vReturnStr = "SELECT CaseNo, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = r.CaseComeFrom) AND (FKEY = '員工獎懲呈報單  RPReport        CaseComeFrom')) AS CaseComeFrom_C, " + Environment.NewLine +
                         "       (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工獎懲呈報單  RPReport        CaseType') AND (CLASSNO = r.CaseType)) AS CaseType_C, " + Environment.NewLine +
                         "       (SELECT [NAME] FROM DEPARTMENT WHERE (DEPNO = r.DepNo)) AS DepName, EmpNo, EmpName, " + Environment.NewLine +
                         "       (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = r.Title) AND (FKEY = '人事資料檔      EMPLOYEE        TITLE')) AS Title_C, " + Environment.NewLine +
                         "       CaseDate, Car_ID, Position, CaseNote " + Environment.NewLine +
                         "  FROM RP_Report AS r " + Environment.NewLine +
                         " WHERE isnull(CaseNo, '') <> '' " + Environment.NewLine +
                         vWStr_AssignDate + vWStr_CaseComeFrom + vWStr_CaseDate + vWStr_CaseType + vWStr_DepNo + vWStr_EmpNo +
                         " order by CaseNo DESC ";
            return vReturnStr;
        }

        /// <summary>
        /// 繫結資料
        /// </summary>
        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                DataTable dtTemp = new DataTable();
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelStr, connTemp);
                connTemp.Open();
                daTemp.Fill(dtTemp);
                gridRPReportList.DataSourceID = "";
                gridRPReportList.DataSource = dtTemp;
                gridRPReportList.DataBind();
            }
        }

        protected void ddlCaseComeFrom_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eCaseComeFrom_Search.Text = ddlCaseComeFrom_Search.SelectedValue.Trim();
        }

        protected void ddlCaseType_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eCaseType_Search.Text = ddlCaseType_Search.SelectedValue.Trim();
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
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Search.Text = vDepNo_Temp.Trim();
            eDepName_Search.Text = vDepName_Temp.Trim();
        }

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo_Temp = eEmpNo_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
            string vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp;
                vSQLStr_Temp = "select top 1 EmpNo from Emplyee where [Name] = '" + vEmpName_Temp + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eEmpNo_Search.Text = vEmpNo_Temp.Trim();
            eEmpName_Search.Text = vEmpName_Temp.Trim();
        }

        /// <summary>
        /// 搜尋
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
            plShowData.Visible = (gridRPReportList.Rows.Count > 0);
            //plSearch.Visible = (gridRPReportList.Rows.Count <= 0);
        }

        protected void gridRPReportList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridRPReportList.PageIndex = e.NewPageIndex;
            gridRPReportList.DataBind();
        }

        protected void fvRPReportDetail_DataBound(object sender, EventArgs e)
        {
            string vDateURL = "";
            string vDateScript = "";
            string vSQLStr_CaseComeFrom = "";
            string vSQLStr_CaseType = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            switch (fvRPReportDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;

                case FormViewMode.Edit:
                    //案件來源
                    using (SqlConnection connCaseComeFrom_Edit = new SqlConnection(vConnStr))
                    {
                        vSQLStr_CaseComeFrom = "SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt " + Environment.NewLine +
                                               " UNION ALL " + Environment.NewLine +
                                               "SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '員工獎懲呈報單  RPReport        CaseComeFrom') ORDER BY ClassNo ";
                        SqlCommand cmdCaseComeFrom_Edit = new SqlCommand(vSQLStr_CaseComeFrom, connCaseComeFrom_Edit);
                        connCaseComeFrom_Edit.Open();

                        SqlDataReader drCaseComeFrom_Edit = cmdCaseComeFrom_Edit.ExecuteReader();
                        DropDownList ddlCaseComeFrom_Edit = (DropDownList)fvRPReportDetail.FindControl("ddlCaseComeFrom_Edit");
                        ddlCaseComeFrom_Edit.Items.Clear();
                        while (drCaseComeFrom_Edit.Read())
                        {
                            ddlCaseComeFrom_Edit.Items.Add(new ListItem(drCaseComeFrom_Edit["ClassTxt"].ToString().Trim(), drCaseComeFrom_Edit["ClassNo"].ToString().Trim()));
                        }
                        Label eCaseComeFrom_Edit = (Label)fvRPReportDetail.FindControl("eCaseComeFrom_Edit");
                        ddlCaseComeFrom_Edit.SelectedIndex = ddlCaseComeFrom_Edit.Items.IndexOf(ddlCaseComeFrom_Edit.Items.FindByValue(eCaseComeFrom_Edit.Text.Trim()));
                        drCaseComeFrom_Edit.Close();
                        cmdCaseComeFrom_Edit.Cancel();
                    }

                    //案件類別
                    using (SqlConnection connCaseType_Edit = new SqlConnection(vConnStr))
                    {
                        vSQLStr_CaseType = "SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt " + Environment.NewLine +
                                           " UNION ALL " + Environment.NewLine +
                                           "SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '員工獎懲呈報單  RPReport        CaseType') ORDER BY ClassNo ";
                        SqlCommand cmdCaseType_Edit = new SqlCommand(vSQLStr_CaseType, connCaseType_Edit);
                        connCaseType_Edit.Open();

                        SqlDataReader drCaseType_Edit = cmdCaseType_Edit.ExecuteReader();
                        DropDownList ddlCaseType_Edit = (DropDownList)fvRPReportDetail.FindControl("ddlCaseType_Edit");
                        ddlCaseType_Edit.Items.Clear();
                        while (drCaseType_Edit.Read())
                        {
                            ddlCaseType_Edit.Items.Add(new ListItem(drCaseType_Edit["ClassTxt"].ToString().Trim(), drCaseType_Edit["ClassNo"].ToString().Trim()));
                        }
                        Label eCaseType_Edit = (Label)fvRPReportDetail.FindControl("eCaseType_Edit");
                        ddlCaseType_Edit.SelectedIndex = ddlCaseType_Edit.Items.IndexOf(ddlCaseType_Edit.Items.FindByValue(eCaseType_Edit.Text.Trim()));
                    }

                    //事件日期
                    TextBox eCaseDate_Edit = (TextBox)fvRPReportDetail.FindControl("eCaseDate_Edit");
                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Edit.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Edit.Attributes["onClick"] = vDateScript;

                    //呈報日期
                    TextBox eAssignDate_Edit = (TextBox)fvRPReportDetail.FindControl("eAssignDate_Edit");
                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eAssignDate_Edit.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eAssignDate_Edit.Attributes["onClick"] = vDateScript;

                    //異動人員
                    Label eModifyMan_Edit = (Label)fvRPReportDetail.FindControl("eModifyMan_Edit");
                    eModifyMan_Edit.Text = vLoginID;

                    Label eModifyManName_Edit = (Label)fvRPReportDetail.FindControl("eModifyManName_Edit");
                    eModifyManName_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID.Trim() + "' ", "Name");

                    //異動日期
                    Label eModifyDate_Edit = (Label)fvRPReportDetail.FindControl("eModifyDate_Edit");
                    eModifyDate_Edit.Text = PF.GetChinsesDate(DateTime.Today);
                    break;

                case FormViewMode.Insert:
                    //案件來源
                    using (SqlConnection connCaseComeFrom_INS = new SqlConnection(vConnStr))
                    {
                        vSQLStr_CaseComeFrom = "SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt " + Environment.NewLine +
                                               " UNION ALL " + Environment.NewLine +
                                               "SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '員工獎懲呈報單  RPReport        CaseComeFrom') ORDER BY ClassNo ";
                        SqlCommand cmdCaseComeFrom_INS = new SqlCommand(vSQLStr_CaseComeFrom, connCaseComeFrom_INS);
                        connCaseComeFrom_INS.Open();

                        SqlDataReader drCaseComeFrom_INS = cmdCaseComeFrom_INS.ExecuteReader();
                        DropDownList ddlCaseComeFrom_INS = (DropDownList)fvRPReportDetail.FindControl("ddlCaseComeFrom_INS");
                        ddlCaseComeFrom_INS.Items.Clear();
                        while (drCaseComeFrom_INS.Read())
                        {
                            ddlCaseComeFrom_INS.Items.Add(new ListItem(drCaseComeFrom_INS["ClassTxt"].ToString().Trim(), drCaseComeFrom_INS["ClassNo"].ToString().Trim()));
                        }
                        ddlCaseComeFrom_INS.SelectedIndex = 0;
                        Label eCaseComeFrom_INS = (Label)fvRPReportDetail.FindControl("eCaseComeFrom_INS");
                        eCaseComeFrom_INS.Text = "";
                        drCaseComeFrom_INS.Close();
                        cmdCaseComeFrom_INS.Cancel();
                    }

                    //案件類別
                    using (SqlConnection connCaseType_INS = new SqlConnection(vConnStr))
                    {
                        vSQLStr_CaseType = "SELECT CAST('' AS varchar) AS ClassNo, CAST('' AS varchar) AS ClassTxt " + Environment.NewLine +
                                              " UNION ALL " + Environment.NewLine +
                                              "SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '員工獎懲呈報單  RPReport        CaseType') ORDER BY ClassNo ";
                        SqlCommand cmdCaseType_INS = new SqlCommand(vSQLStr_CaseType, connCaseType_INS);
                        connCaseType_INS.Open();

                        SqlDataReader drCaseType_INS = cmdCaseType_INS.ExecuteReader();
                        DropDownList ddlCaseType_INS = (DropDownList)fvRPReportDetail.FindControl("ddlCaseType_INS");
                        ddlCaseType_INS.Items.Clear();
                        while (drCaseType_INS.Read())
                        {
                            ddlCaseType_INS.Items.Add(new ListItem(drCaseType_INS["ClassTxt"].ToString().Trim(), drCaseType_INS["ClassNo"].ToString().Trim()));
                        }
                        ddlCaseType_INS.SelectedIndex = 0;
                        Label eCaseType_INS = (Label)fvRPReportDetail.FindControl("eCaseType_INS");
                        eCaseType_INS.Text = "";
                        drCaseType_INS.Close();
                        cmdCaseType_INS.Cancel();
                    }

                    //事件日期
                    TextBox eCaseDate_INS = (TextBox)fvRPReportDetail.FindControl("eCaseDate_INS");
                    eCaseDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_INS.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_INS.Attributes["onClick"] = vDateScript;

                    //呈報日期
                    TextBox eAssignDate_INS = (TextBox)fvRPReportDetail.FindControl("eAssignDate_INS");
                    eAssignDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                    vDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eAssignDate_INS.ClientID;
                    vDateScript = "window.open('" + vDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eAssignDate_INS.Attributes["onClick"] = vDateScript;

                    //呈報人
                    TextBox eAssignMan_INS = (TextBox)fvRPReportDetail.FindControl("eAssignMan_INS");
                    eAssignMan_INS.Text = vLoginID.Trim();

                    Label eAssignManName_INS = (Label)fvRPReportDetail.FindControl("eAssignManName_INS");
                    eAssignManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID.Trim() + "' ", "Name");

                    //呈報單位
                    Label eAssignDepNo_INS = (Label)fvRPReportDetail.FindControl("eAssignDepNo_INS");
                    eAssignDepNo_INS.Text = vLoginDepNo.Trim();

                    Label eAssignDepName_INS = (Label)fvRPReportDetail.FindControl("eAssignDepName_INS");
                    eAssignDepName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Department where DepNo = '" + vLoginDepNo.Trim() + "' ", "Name");

                    //建檔人
                    Label eBuMan_INS = (Label)fvRPReportDetail.FindControl("eBuMan_INS");
                    eBuMan_INS.Text = vLoginID.Trim();

                    Label eBuManName_INS = (Label)fvRPReportDetail.FindControl("eBuManName_INS");
                    eBuManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID.Trim() + "' ", "Name");

                    //建檔日期
                    Label eBuDate_INS = (Label)fvRPReportDetail.FindControl("eBuDate_INS");
                    eBuDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                    break;
            }
        }

        protected void ddlCaseComeFrom_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlCaseComeFrom_Edit = (DropDownList)fvRPReportDetail.FindControl("ddlCaseComeFrom_Edit");
            Label eCaseComeFrom_Edit = (Label)fvRPReportDetail.FindControl("eCaseComeFrom_Edit");
            eCaseComeFrom_Edit.Text = ddlCaseComeFrom_Edit.SelectedValue.Trim();
        }

        protected void ddlCaseType_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlCaseType_Edit = (DropDownList)fvRPReportDetail.FindControl("ddlCaseType_Edit");
            Label eCaseType_Edit = (Label)fvRPReportDetail.FindControl("eCaseType_Edit");
            eCaseType_Edit.Text = ddlCaseType_Edit.SelectedValue.Trim();
        }

        protected void eEmpNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eEmpNo_Edit = (TextBox)fvRPReportDetail.FindControl("eEmpNo_Edit");
            Label eEmpName_Edit = (Label)fvRPReportDetail.FindControl("eEmpName_Edit");
            TextBox eDepNo_Edit = (TextBox)fvRPReportDetail.FindControl("eDepNo_Edit");
            Label eDepName_Edit = (Label)fvRPReportDetail.FindControl("eDepName_Edit");
            TextBox eTitle_Edit = (TextBox)fvRPReportDetail.FindControl("eTitle_Edit");
            Label eTitle_C_Edit = (Label)fvRPReportDetail.FindControl("eTitle_C_Edit");

            string vEmpNo_Edit = eEmpNo_Edit.Text.Trim();
            string vGetDataStr = "select EmpNo + ',' + [Name] + ',' + DepNo + ',' + (select [Name] from Department where DepNo = e.DepNo) + ',' + " + Environment.NewLine +
                                 "       Title + ',' + (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = e.Title) AND (FKEY = '人事資料檔      EMPLOYEE        TITLE')) as ResultStr " + Environment.NewLine +
                                 "  from Employee e " + Environment.NewLine +
                                 " where EmpNo = '" + vEmpNo_Edit + "' ";
            string vResultStr = PF.GetValue(vConnStr, vGetDataStr, "ResultStr");
            string[] arrayResult = vResultStr.Split(',');
            eEmpName_Edit.Text = arrayResult[1].ToString().Trim();
            eDepNo_Edit.Text = arrayResult[2].ToString().Trim();
            eDepName_Edit.Text = arrayResult[3].ToString().Trim();
            eTitle_Edit.Text = arrayResult[4].ToString().Trim();
            eTitle_C_Edit.Text = arrayResult[5].ToString().Trim();
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Edit = (TextBox)fvRPReportDetail.FindControl("eDepNo_Edit");
            Label eDepName_Edit = (Label)fvRPReportDetail.FindControl("eDepName_Edit");
            string vDepNo_Edit = eDepNo_Edit.Text.Trim();
            string vGetDataStr = "select [Name] from Department where DepNo = '" + vDepNo_Edit + "' ";
            eDepName_Edit.Text = PF.GetValue(vConnStr, vGetDataStr, "Name");
        }

        protected void eTitle_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eTitle_Edit = (TextBox)fvRPReportDetail.FindControl("eTitle_Edit");
            Label eTitle_C_Edit = (Label)fvRPReportDetail.FindControl("eTitle_C_Edit");
            string vTitle_Edit = eTitle_C_Edit.Text.Trim();
            string vGetDataStr = "select ClassTxt from DBDICB where ClassNo = '" + vTitle_Edit + "' and FKey = '人事資料檔      EMPLOYEE        TITLE'";
            eTitle_C_Edit.Text = PF.GetValue(vConnStr, vGetDataStr, "ClassTxt");
        }

        protected void ddlCaseComeFrom_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlCaseComeFrom_INS = (DropDownList)fvRPReportDetail.FindControl("ddlCaseComeFrom_INS");
            Label eCaseComeFrom_INS = (Label)fvRPReportDetail.FindControl("eCaseComeFrom_INS");
            eCaseComeFrom_INS.Text = ddlCaseComeFrom_INS.SelectedValue.Trim();
        }

        protected void ddlCaseType_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlCaseType_INS = (DropDownList)fvRPReportDetail.FindControl("ddlCaseType_INS");
            Label eCaseType_INS = (Label)fvRPReportDetail.FindControl("eCaseType_INS");
            eCaseType_INS.Text = ddlCaseType_INS.SelectedValue.Trim();
        }

        protected void eEmpNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eEmpNo_INS = (TextBox)fvRPReportDetail.FindControl("eEmpNo_INS");
            Label eEmpName_INS = (Label)fvRPReportDetail.FindControl("eEmpName_INS");
            TextBox eDepNo_INS = (TextBox)fvRPReportDetail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvRPReportDetail.FindControl("eDepName_INS");
            TextBox eTitle_INS = (TextBox)fvRPReportDetail.FindControl("eTitle_INS");
            Label eTitle_C_INS = (Label)fvRPReportDetail.FindControl("eTitle_C_INS");

            string vEmpNo_INS = eEmpNo_INS.Text.Trim();
            string vGetDataStr = "select EmpNo + ',' + [Name] + ',' + DepNo + ',' + (select [Name] from Department where DepNo = e.DepNo) + ',' + " + Environment.NewLine +
                                 "       Title + ',' + (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = e.Title) AND (FKEY = '人事資料檔      EMPLOYEE        TITLE')) as ResultStr " + Environment.NewLine +
                                 "  from Employee e " + Environment.NewLine +
                                 " where EmpNo = '" + vEmpNo_INS + "' ";
            string vResultStr = PF.GetValue(vConnStr, vGetDataStr, "ResultStr");
            string[] arrayResult = vResultStr.Split(',');
            eEmpName_INS.Text = arrayResult[1].ToString().Trim();
            eDepNo_INS.Text = arrayResult[2].ToString().Trim();
            eDepName_INS.Text = arrayResult[3].ToString().Trim();
            eTitle_INS.Text = arrayResult[4].ToString().Trim();
            eTitle_C_INS.Text = arrayResult[5].ToString().Trim();
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_INS = (TextBox)fvRPReportDetail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvRPReportDetail.FindControl("eDepName_INS");
            string vDepNo_INS = eDepNo_INS.Text.Trim();
            string vGetDataStr = "select [Name] from Department where DepNo = '" + vDepNo_INS + "' ";
            eDepName_INS.Text = PF.GetValue(vConnStr, vGetDataStr, "Name");
        }

        protected void eTitle_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eTitle_INS = (TextBox)fvRPReportDetail.FindControl("eTitle_INS");
            Label eTitle_C_INS = (Label)fvRPReportDetail.FindControl("eTitle_C_INS");
            string vTitle_INS = eTitle_C_INS.Text.Trim();
            string vGetDataStr = "select ClassTxt from DBDICB where ClassNo = '" + vTitle_INS + "' and FKey = '人事資料檔      EMPLOYEE        TITLE'";
            eTitle_C_INS.Text = PF.GetValue(vConnStr, vGetDataStr, "ClassTxt");
        }

        protected void eAssignMan_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eAssignMan_INS = (TextBox)fvRPReportDetail.FindControl("eAssignMan_INS");

            Label eAssignManName_INS = (Label)fvRPReportDetail.FindControl("eAssignManName_INS");
            eAssignManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + eAssignMan_INS.Text.Trim() + "' ", "Name");

            Label eAssignDepNo_INS = (Label)fvRPReportDetail.FindControl("eAssignDepNo_INS");
            eAssignDepNo_INS.Text = PF.GetValue(vConnStr, "select DepNo from Employee where EmpNo = '" + eAssignMan_INS.Text.Trim() + "' ", "DepNo");

            Label eAssignDepName_INS = (Label)fvRPReportDetail.FindControl("eAssignDepName_INS");
            eAssignDepName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Department where DepNo = '" + eAssignDepNo_INS.Text.Trim() + "' ", "Name");
        }

        /// <summary>
        /// 列印獎懲呈報單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_List_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eCaseNo_List = (Label)fvRPReportDetail.FindControl("eCaseNo_List");
            string vCaseNo_List = eCaseNo_List.Text.Trim();
            string vSelStr = GetSelectStr_Print(vCaseNo_List);
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintPoint = new SqlDataAdapter(vSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintPoint = new DataTable("RP_ReportP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("RP_ReportP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\RP_ReportP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "桃園汽車客運股份有限公司員工獎懲呈報單"));
                    rvPrint.LocalReport.Refresh();
                    plReport.Visible = true;
                    plShowData.Visible = false;

                    string vCaseComefromStr = (eCaseComeFrom_Search.Text.Trim() != "") ? eCaseComeFrom_Search.Text.Trim() : "全部";
                    string vCaseTypeStr = (eCaseType_Search.Text.Trim() != "") ? eCaseType_Search.Text.Trim() : "全部";
                    string vAssignDateStr = ((eAssignDate_Search_S.Text.Trim() != "") && (eAssignDate_Search_E.Text.Trim() != "")) ? "自 " + eAssignDate_Search_S.Text.Trim() + " 起至 " + eAssignDate_Search_E.Text.Trim() + " 止" :
                                            ((eAssignDate_Search_S.Text.Trim() != "") && (eAssignDate_Search_E.Text.Trim() == "")) ? eAssignDate_Search_S.Text.Trim() :
                                            ((eAssignDate_Search_S.Text.Trim() == "") && (eAssignDate_Search_E.Text.Trim() != "")) ? eAssignDate_Search_E.Text.Trim() : "不分日期";
                    string vCaseDateStr = ((eCaseDate_Search_S.Text.Trim() != "") && (eCaseDate_Search_E.Text.Trim() != "")) ? "自 " + eCaseDate_Search_S.Text.Trim() + " 起至 " + eCaseDate_Search_E.Text.Trim() + " 止" :
                                          ((eCaseDate_Search_S.Text.Trim() != "") && (eCaseDate_Search_E.Text.Trim() == "")) ? eCaseDate_Search_S.Text.Trim() :
                                          ((eCaseDate_Search_S.Text.Trim() == "") && (eCaseDate_Search_E.Text.Trim() != "")) ? eCaseDate_Search_E.Text.Trim() : "不分日期";
                    string vDepNoStr = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "全部";
                    string vEmpNoStr = (eEmpNo_Search.Text.Trim() != "") ? eEmpNo_Search.Text.Trim() : "全部";
                    string vCarIDStr = (eCar_ID_Search.Text.Trim() != "") ? eCar_ID_Search.Text.Trim() : "";
                    string vRecordNote = "預覽報表_員工獎懲呈報單" + Environment.NewLine +
                                         "RPReport.aspx" + Environment.NewLine +
                                         "案件來源：" + vCaseComefromStr + Environment.NewLine +
                                         "案件類別：" + vCaseTypeStr + Environment.NewLine +
                                         "呈報日期：" + vAssignDateStr + Environment.NewLine +
                                         "事件日期：" + vCaseDateStr + Environment.NewLine +
                                         "站別：" + vDepNoStr + Environment.NewLine +
                                         "員工工號：" + vEmpNoStr + Environment.NewLine +
                                         "牌照號碼：" + vCarIDStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        /// <summary>
        /// 刪除獎懲呈報單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDEL_List_Click(object sender, EventArgs e)
        {
            Label eCaseNo_List = (Label)fvRPReportDetail.FindControl("eCaseNo_List");
            string vCaseNo_List = eCaseNo_List.Text.Trim();
            if (vCaseNo_List != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                //刪除前先複製一份內容到異動記錄
                string vGetMaxCountStr = "select isnull(MAX(Items), 0) Items from RP_Report_History where CaseNo = '" + vCaseNo_List + "' ";
                int vItems = Int32.Parse(PF.GetValue(vConnStr, vGetMaxCountStr, "Items")) + 1;
                string vSaveToHistoryStr = "INSERT INTO [dbo].[RP_Report_History] " + Environment.NewLine +
                                           "       (CaseNoItem, Items,ModifyMode, CaseNo, CaseComeFrom, CaseType, CaseTypeB, CaseTypeC, " + Environment.NewLine +
                                           "        DepNo, EmpNo, EmpName, Title, CaseDate, Car_ID, Position, CaseNote, " + Environment.NewLine +
                                           "        AccordingTerms, Review, Remark, AssignDate, AssignDepNo, AssignMan, " + Environment.NewLine +
                                           "        BuDate, BuMan, ModifyDate, ModifyMan, CustomServiceNo)" + Environment.NewLine +
                                           "select '" + vCaseNo_List.Trim() + vItems.ToString("D4") + "', '" + vItems.ToString("D4") + "', 'DEL', " + Environment.NewLine +
                                           "       CaseNo, CaseComeFrom, CaseType, CaseTypeB, CaseTypeC, DepNo, EmpNo, EmpName, Title, " + Environment.NewLine +
                                           "       CaseDate, Car_ID, Position, CaseNote, AccordingTerms, Review, Remark, " + Environment.NewLine +
                                           "       AssignDate, AssignDepNo, AssignMan, BuDate, BuMan, ModifyDate, ModifyMan, CustomServiceNo " + Environment.NewLine +
                                           "  from RP_Report " + Environment.NewLine +
                                           " where CaseNo = '" + vCaseNo_List.Trim() + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vSaveToHistoryStr);
                    sdsRPReportDetail.DeleteParameters.Clear();
                    sdsRPReportDetail.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_List));
                    sdsRPReportDetail.Delete();
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
                Response.Write("alert('請先選擇要刪除的呈報單！')");
                Response.Write("</" + "Script>");
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
            plReport.Visible = false;
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        private string GetSelectStr_Excel()
        {
            string vReturnStr = "";
            DateTime vAssignDate_S_Temp = (eAssignDate_Search_S.Text.Trim() != "") ?
                                          DateTime.Parse(eAssignDate_Search_S.Text.Trim()) :
                                          DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C"));
            DateTime vAssignDate_E_Temp = (eAssignDate_Search_E.Text.Trim() != "") ?
                                          DateTime.Parse(eAssignDate_Search_E.Text.Trim()) :
                                          DateTime.Parse(PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C"));
            DateTime vCaseDate_S_Temp = (eCaseDate_Search_S.Text.Trim() != "") ?
                                        DateTime.Parse(eCaseDate_Search_S.Text.Trim()) :
                                        DateTime.Parse(PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "C"));
            DateTime vCaseDate_E_Temp = (eCaseDate_Search_E.Text.Trim() != "") ?
                                        DateTime.Parse(eCaseDate_Search_E.Text.Trim()) :
                                        DateTime.Parse(PF.GetMonthLastDay(DateTime.Today.AddMonths(-1), "C"));
            string vWStr_AssignDate = ((eAssignDate_Search_S.Text.Trim() != "") && (eAssignDate_Search_E.Text.Trim() != "")) ?
                                      "   and AssignDate between '" + vAssignDate_S_Temp.Year.ToString() + "/" + vAssignDate_S_Temp.ToString("MM/dd") + "' and '" + vAssignDate_E_Temp.Year.ToString() + "/" + vAssignDate_E_Temp.ToString("MM/dd") + "' " + Environment.NewLine :
                                      ((eAssignDate_Search_S.Text.Trim() != "") && (eAssignDate_Search_E.Text.Trim() == "")) ?
                                      "   and AssignDate = '" + vAssignDate_S_Temp.Year.ToString() + "/" + vAssignDate_S_Temp.ToString("MM/dd") + "' " + Environment.NewLine :
                                      ((eAssignDate_Search_S.Text.Trim() == "") && (eAssignDate_Search_E.Text.Trim() != "")) ?
                                      "   and AssignDate = '" + vAssignDate_E_Temp.Year.ToString() + "/" + vAssignDate_E_Temp.ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_CaseDate = ((eCaseDate_Search_S.Text.Trim() != "") && (eCaseDate_Search_E.Text.Trim() != "")) ?
                                    "   and CaseDate between '" + vCaseDate_S_Temp.Year.ToString() + "/" + vCaseDate_S_Temp.ToString("MM/dd") + "' and '" + vCaseDate_E_Temp.Year.ToString() + "/" + vCaseDate_E_Temp.ToString("MM/dd") + "' " + Environment.NewLine :
                                    ((eCaseDate_Search_S.Text.Trim() != "") && (eCaseDate_Search_E.Text.Trim() == "")) ?
                                    "   and CaseDate = '" + vCaseDate_S_Temp.Year.ToString() + "/" + vCaseDate_S_Temp.ToString("MM/dd") + "' " + Environment.NewLine :
                                    ((eCaseDate_Search_S.Text.Trim() == "") && (eCaseDate_Search_E.Text.Trim() != "")) ?
                                    "   and CaseDate = '" + vCaseDate_E_Temp.Year.ToString() + "/" + vCaseDate_E_Temp.ToString("MM/dd") + "' " + Environment.NewLine : "";
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and EmpNo = '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CaseComeFrom = (eCaseComeFrom_Search.Text.Trim() != "") ? "   and CaseComeFrom = '" + eCaseComeFrom_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CaseType = (eCaseType_Search.Text.Trim() != "") ? "   and CaseType = '" + eCaseType_Search.Text.Trim() + "' " + Environment.NewLine : "";
            vReturnStr = "SELECT CaseNo, (select ClassTxt from DBDICB where ClassNo = r.CaseComeFrom and FKey = '員工獎懲呈報單  RPReport        CaseComeFrom') CaseComeFrom_C, " + Environment.NewLine +
                         "       (SELECT ClassTxt FROM DBDICB WHERE(FKey = '員工獎懲呈報單  RPReport        CaseType') AND(ClassNo = r.CaseType)) AS CaseType_C, " + Environment.NewLine +
                         "       case when isnull(GiveBounds, 0) = 1 then BoundsAmount else null end BoundsAmount, " + Environment.NewLine +
                         "       case when isnull(Commendation, 0) = 1 then CommendationCount else null end CommendationCount, " + Environment.NewLine +
                         "       case when isnull(MeritCitation, 0) = 1 then MeritCitationCount else null end MeritCitationCount, " + Environment.NewLine +
                         "       case when isnull(MajorMeritCitation, 0) = 1 then MajorMeritCitationCount else null end MajorMeritCitationCount, " + Environment.NewLine +
                         "       case when isnull(AskExact, 0) = 1 then ExactAmount else null end ExactAmount, " + Environment.NewLine +
                         "       case when isnull(Advice, 0) = 1 then 'V' else '' end Advice, " + Environment.NewLine +
                         "       case when isnull(Admonition, 0) = 1 then 'V' else '' end Admonition, " + Environment.NewLine +
                         "       case when isnull(Reprimand, 0) = 1 then ReprimandCount else null end ReprimandCount, " + Environment.NewLine +
                         "       case when isnull(Demerit, 0) = 1 then DemeritCount else null end DemeritCount, " + Environment.NewLine +
                         "       case when isnull(MajorDemerit, 0) = 1 then MajorDemeritCount else null end MajorDemeritCount, " + Environment.NewLine +
                         "       case when isnull(Promotion, 0) = 1 then 'V'  else '' end Promotion, " + Environment.NewLine +
                         "       case when isnull(Demotion, 0) = 1 then 'V' else '' end Demotion, " + Environment.NewLine +
                         "       case when isnull(Dismissal, 0) = 1 then 'V' else '' end Dismissal, " + Environment.NewLine +
                         "       isnull(Review, '') as Review, " + Environment.NewLine +
                         "       (select[Name] from Department where DepNo = r.DepNo) DepName, " + Environment.NewLine +
                         "       EmpNo, EmpName, (SELECT CLASSTXT FROM DBDICB WHERE(CLASSNO = r.Title) AND(FKEY = '人事資料檔      EMPLOYEE        TITLE')) AS Title_C, " + Environment.NewLine +
                         "       CaseDate, AssignDate, Car_ID, Position, CaseNote " + Environment.NewLine +
                         "  FROM RP_Report r " + Environment.NewLine +
                         " WHERE isnull(CaseNo, '') <> '' " + Environment.NewLine +
                         vWStr_AssignDate + vWStr_CaseComeFrom + vWStr_CaseDate + vWStr_CaseType + vWStr_DepNo + vWStr_EmpNo +
                         " order by CaseNo DESC ";
            return vReturnStr;
        }

        /// <summary>
        /// 資料匯出 EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_List_Click(object sender, EventArgs e)
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
            string vFileName = "員工獎懲呈報清冊";
            DateTime vBuDate;
            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                string vSelStr = GetSelectStr_Excel();
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
                        vHeaderText = (drExcel.GetName(i).ToUpper() == "CASENO") ? "單號" :
                                      (drExcel.GetName(i).ToUpper() == "CASECOMEFROM_C") ? "案件來源" :
                                      (drExcel.GetName(i).ToUpper() == "CASETYPE_C") ? "獎懲類別" :
                                      (drExcel.GetName(i).ToUpper() == "BOUNDSAMOUNT") ? "獎金" :
                                      (drExcel.GetName(i).ToUpper() == "COMMENDATIONCOUNT") ? "嘉獎" :
                                      (drExcel.GetName(i).ToUpper() == "MERITCITATIONCOUNT") ? "記功" :
                                      (drExcel.GetName(i).ToUpper() == "MAJORMERITCITATIONCOUNT") ? "大功" :
                                      (drExcel.GetName(i).ToUpper() == "EXACTAMOUNT") ? "賠償" :
                                      (drExcel.GetName(i).ToUpper() == "ADVICE") ? "勸導" :
                                      (drExcel.GetName(i).ToUpper() == "ADMONITION") ? "口頭警告" :
                                      (drExcel.GetName(i).ToUpper() == "REPRIMANDCOUNT") ? "申誡" :
                                      (drExcel.GetName(i).ToUpper() == "DEMERITCOUNT") ? "記過" :
                                      (drExcel.GetName(i).ToUpper() == "MAJORDEMERITCOUNT") ? "大過" :
                                      (drExcel.GetName(i).ToUpper() == "PROMOTION") ? "晉升" :
                                      (drExcel.GetName(i).ToUpper() == "DEMOTION") ? "降調" :
                                      (drExcel.GetName(i).ToUpper() == "DISMISSAL") ? "解職" :
                                      (drExcel.GetName(i).ToUpper() == "REVIEW") ? "其他" :
                                      (drExcel.GetName(i).ToUpper() == "DEPNAME") ? "站別" :
                                      (drExcel.GetName(i).ToUpper() == "EMPNO") ? "工號" :
                                      (drExcel.GetName(i).ToUpper() == "EMPNAME") ? "員工姓名" :
                                      (drExcel.GetName(i).ToUpper() == "TITLE_C") ? "職稱" :
                                      (drExcel.GetName(i).ToUpper() == "CASEDATE") ? "事件日期" :
                                      (drExcel.GetName(i).ToUpper() == "ASSIGNDATE") ? "呈報日期" :
                                      (drExcel.GetName(i).ToUpper() == "CAR_ID") ? "車號" :
                                      (drExcel.GetName(i).ToUpper() == "POSITION") ? "地點" :
                                      (drExcel.GetName(i).ToUpper() == "CASENOTE") ? "事件概述" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i).ToUpper() == "CASEDATE") ||
                                 (drExcel.GetName(i).ToUpper() == "ASSIGNDATE")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if (((drExcel.GetName(i).ToUpper() == "BOUNDSAMOUNT") ||
                                      (drExcel.GetName(i).ToUpper() == "COMMENDATIONCOUNT") ||
                                      (drExcel.GetName(i).ToUpper() == "MERITCITATIONCOUNT") ||
                                      (drExcel.GetName(i).ToUpper() == "MAJORMERITCITATIONCOUNT") ||
                                      (drExcel.GetName(i).ToUpper() == "EXACTAMOUNT") ||
                                      (drExcel.GetName(i).ToUpper() == "REPRIMANDCOUNT") ||
                                      (drExcel.GetName(i).ToUpper() == "DEMERITCOUNT") ||
                                      (drExcel.GetName(i).ToUpper() == "MAJORDEMERITCOUNT")) && (drExcel[i].ToString() != ""))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(Double.Parse(drExcel[i].ToString().Trim()));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
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
                            string vCaseComefromStr = (eCaseComeFrom_Search.Text.Trim() != "") ? eCaseComeFrom_Search.Text.Trim() : "全部";
                            string vCaseTypeStr = (eCaseType_Search.Text.Trim() != "") ? eCaseType_Search.Text.Trim() : "全部";
                            string vAssignDateStr = ((eAssignDate_Search_S.Text.Trim() != "") && (eAssignDate_Search_E.Text.Trim() != "")) ? "自 " + eAssignDate_Search_S.Text.Trim() + " 起至 " + eAssignDate_Search_E.Text.Trim() + " 止" :
                                                    ((eAssignDate_Search_S.Text.Trim() != "") && (eAssignDate_Search_E.Text.Trim() == "")) ? eAssignDate_Search_S.Text.Trim() :
                                                    ((eAssignDate_Search_S.Text.Trim() == "") && (eAssignDate_Search_E.Text.Trim() != "")) ? eAssignDate_Search_E.Text.Trim() : "不分日期";
                            string vCaseDateStr = ((eCaseDate_Search_S.Text.Trim() != "") && (eCaseDate_Search_E.Text.Trim() != "")) ? "自 " + eCaseDate_Search_S.Text.Trim() + " 起至 " + eCaseDate_Search_E.Text.Trim() + " 止" :
                                                  ((eCaseDate_Search_S.Text.Trim() != "") && (eCaseDate_Search_E.Text.Trim() == "")) ? eCaseDate_Search_S.Text.Trim() :
                                                  ((eCaseDate_Search_S.Text.Trim() == "") && (eCaseDate_Search_E.Text.Trim() != "")) ? eCaseDate_Search_E.Text.Trim() : "不分日期";
                            string vDepNoStr = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : "全部";
                            string vEmpNoStr = (eEmpNo_Search.Text.Trim() != "") ? eEmpNo_Search.Text.Trim() : "全部";
                            string vCarIDStr = (eCar_ID_Search.Text.Trim() != "") ? eCar_ID_Search.Text.Trim() : "";
                            string vRecordNote = "匯出檔案_員工獎懲呈報單" + Environment.NewLine +
                                                 "RPReport.aspx" + Environment.NewLine +
                                                 "案件來源：" + vCaseComefromStr + Environment.NewLine +
                                                 "案件類別：" + vCaseTypeStr + Environment.NewLine +
                                                 "呈報日期：" + vAssignDateStr + Environment.NewLine +
                                                 "事件日期：" + vCaseDateStr + Environment.NewLine +
                                                 "站別：" + vDepNoStr + Environment.NewLine +
                                                 "員工工號：" + vEmpNoStr + Environment.NewLine +
                                                 "牌照號碼：" + vCarIDStr;
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

        /// <summary>
        /// 修改確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            try
            {
                Label eCaseNo_Edit = (Label)fvRPReportDetail.FindControl("eCaseNo_Edit");
                //先備份到異動檔
                string vCaseNo_Edit = eCaseNo_Edit.Text.Trim();
                string vGetMaxCountStr = "select isnull(MAX(Items), 0) Items from RP_Report_History where CaseNo = '" + vCaseNo_Edit + "' ";
                int vItems = Int32.Parse(PF.GetValue(vConnStr, vGetMaxCountStr, "Items")) + 1;
                string vSaveToHistoryStr = "INSERT INTO RP_Report_History " + Environment.NewLine +
                                           "       (CaseNoItem, Items, ModifyMode, CaseNo, CaseComeFrom, CaseType, CaseTypeB, CaseTypeC, DepNo, EmpNo, " + Environment.NewLine +
                                           "        EmpName, Title, CaseDate, Car_ID, Position, CaseNote, AccordingTerms, Review, Remark, AssignDate, " + Environment.NewLine +
                                           "        AssignDepNo, AssignMan, BuDate, BuMan, ModifyDate, ModifyMan, CustomServiceNo, GiveBounds, BoundsAmount, " + Environment.NewLine +
                                           "        AskExact, ExactAmount, Demotion, Dismissal, Promotion, Advice, Admonition, Reprimand, ReprimandCount, Demerit, " + Environment.NewLine +
                                           "        DemeritCount, MajorDemerit, MajorDemeritCount, Commendation, CommendationCount, MeritCitation, " + Environment.NewLine +
                                           "        MeritCitationCount, MajorMeritCitation, MajorMeritCitationCount, Others) " + Environment.NewLine +
                                           "SELECT '" + vCaseNo_Edit.Trim() + vItems.ToString("D4") + "', '" + vItems.ToString("D4") + "', 'EDIT', " + Environment.NewLine +
                                           "       CaseNo, CaseComeFrom, CaseType, CaseTypeB, CaseTypeC, DepNo, EmpNo, EmpName, Title, " + Environment.NewLine +
                                           "       CaseDate, Car_ID, Position, CaseNote, AccordingTerms, Review, Remark, AssignDate, AssignDepNo, AssignMan, " + Environment.NewLine +
                                           "       BuDate, BuMan, ModifyDate, ModifyMan, CustomServiceNo, GiveBounds, BoundsAmount, AskExact, ExactAmount, " + Environment.NewLine +
                                           "       Demotion, Dismissal, Promotion, Advice, Admonition, Reprimand, ReprimandCount, Demerit, DemeritCount, " + Environment.NewLine +
                                           "       MajorDemerit, MajorDemeritCount, Commendation, CommendationCount, MeritCitation, MeritCitationCount, " + Environment.NewLine +
                                           "       MajorMeritCitation, MajorMeritCitationCount, Others " + Environment.NewLine +
                                           "  FROM RP_Report " + Environment.NewLine +
                                           " where CaseNo = '" + vCaseNo_Edit.Trim() + "' ";
                PF.ExecSQL(vConnStr, vSaveToHistoryStr);

                //開始寫入修改資料
                Label eCaseComeFrom_Edit = (Label)fvRPReportDetail.FindControl("eCaseComeFrom_Edit");
                Label eCaseType_Edit = (Label)fvRPReportDetail.FindControl("eCaseType_Edit");
                TextBox eDepNo_Edit = (TextBox)fvRPReportDetail.FindControl("eDepNo_Edit");
                TextBox eEmpNo_Edit = (TextBox)fvRPReportDetail.FindControl("eEmpNo_Edit");
                Label eEmpName_Edit = (Label)fvRPReportDetail.FindControl("eEmpName_Edit");
                TextBox eTitle_Edit = (TextBox)fvRPReportDetail.FindControl("eTitle_Edit");
                TextBox eCaseDate_Edit = (TextBox)fvRPReportDetail.FindControl("eCaseDate_Edit");
                TextBox eCar_ID_Edit = (TextBox)fvRPReportDetail.FindControl("eCar_ID_Edit");
                TextBox ePosition_Edit = (TextBox)fvRPReportDetail.FindControl("ePosition_Edit");
                TextBox eCaseNote_Edit = (TextBox)fvRPReportDetail.FindControl("eCaseNote_Edit");
                TextBox eAccordingTerms_Edit = (TextBox)fvRPReportDetail.FindControl("eAccordingTerms_Edit");
                TextBox eReview_Edit = (TextBox)fvRPReportDetail.FindControl("eReview_Edit");
                TextBox eRemark_Edit = (TextBox)fvRPReportDetail.FindControl("eRemark_Edit");
                //TextBox eAssignDate_Edit = (TextBox)fvRPReportDetail.FindControl("eAssignDate_Edit");
                //DateTime vAssignDate_Temp = DateTime.Parse(eAssignDate_Edit.Text.Trim());
                //Label eAssignDepNo_Edit = (Label)fvRPReportDetail.FindControl("eAssignDepNo_Edit");
                //Label eAssignMan_Edit = (Label)fvRPReportDetail.FindControl("eAssignMan_Edit");
                CheckBox eGiveBounds_Edit = (CheckBox)fvRPReportDetail.FindControl("eGiveBounds_Edit");
                TextBox eBoundsAmount_Edit = (TextBox)fvRPReportDetail.FindControl("eBoundsAmount_Edit");
                CheckBox eAskExact_Edit = (CheckBox)fvRPReportDetail.FindControl("eAskExact_Edit");
                TextBox eExactAmount_Edit = (TextBox)fvRPReportDetail.FindControl("eExactAmount_Edit");
                CheckBox eDemotion_Edit = (CheckBox)fvRPReportDetail.FindControl("eDemotion_Edit");
                CheckBox eDismissal_Edit = (CheckBox)fvRPReportDetail.FindControl("eDismissal_Edit");
                CheckBox ePromotion_Edit = (CheckBox)fvRPReportDetail.FindControl("ePromotion_Edit");
                CheckBox eAdvice_Edit = (CheckBox)fvRPReportDetail.FindControl("eAdvice_Edit");
                CheckBox eAdmonition_Edit = (CheckBox)fvRPReportDetail.FindControl("eAdmonition_Edit");
                CheckBox eReprimand_Edit = (CheckBox)fvRPReportDetail.FindControl("eReprimand_Edit");
                TextBox eReprimandCount_Edit = (TextBox)fvRPReportDetail.FindControl("eReprimandCount_Edit");
                CheckBox eDemerit_Edit = (CheckBox)fvRPReportDetail.FindControl("eDemerit_Edit");
                TextBox eDemeritCount_Edit = (TextBox)fvRPReportDetail.FindControl("eDemeritCount_Edit");
                CheckBox eMajorDemerit_Edit = (CheckBox)fvRPReportDetail.FindControl("eMajorDemerit_Edit");
                TextBox eMajorDemeritCount_Edit = (TextBox)fvRPReportDetail.FindControl("eMajorDemeritCount_Edit");
                CheckBox eCommendation_Edit = (CheckBox)fvRPReportDetail.FindControl("eCommendation_Edit");
                TextBox eCommendationCount_Edit = (TextBox)fvRPReportDetail.FindControl("eCommendationCount_Edit");
                CheckBox eMeritCitation_Edit = (CheckBox)fvRPReportDetail.FindControl("eMeritCitation_Edit");
                TextBox eMeritCitationCount_Edit = (TextBox)fvRPReportDetail.FindControl("eMeritCitationCount_Edit");
                CheckBox eMajorMeritCitation_Edit = (CheckBox)fvRPReportDetail.FindControl("eMajorMeritCitation_Edit");
                TextBox eMajorMeritCitationCount_Edit = (TextBox)fvRPReportDetail.FindControl("eMajorMeritCitationCount_Edit");
                CheckBox eOthers_Edit = (CheckBox)fvRPReportDetail.FindControl("eOthers_Edit");

                string vSQLStr_Update = "UPDATE RP_Report " + Environment.NewLine +
                                        "   SET CaseComeFrom = @CaseComeFrom, CaseType = @CaseType, DepNo = @DepNo, EmpNo = @EmpNo, " + Environment.NewLine +
                                        "       EmpName = @EmpName, Title = @Title, CaseDate = @CaseDate, Car_ID = @Car_ID, Position = @Position, " + Environment.NewLine +
                                        "       CaseNote = @CaseNote, AccordingTerms = @AccordingTerms, Review = @Review, Remark = @Remark, " + Environment.NewLine +
                                        //"       AssignDate = @AssignDate, AssignDepNo = @AssignDepNo, AssignMan = @AssignMan, " + Environment.NewLine +
                                        "       ModifyDate = @ModifyDate, ModifyMan = @ModifyMan, GiveBounds = @GiveBounds, " + Environment.NewLine +
                                        "       BoundsAmount = @BoundsAmount, AskExact = @AskExact, ExactAmount = @ExactAmount, Demotion = @Demotion, " + Environment.NewLine +
                                        "       Dismissal = @Dismissal, Promotion = @Promotion, Advice = @Advice, Admonition = @Admonition, " + Environment.NewLine +
                                        "       Reprimand = @Reprimand, ReprimandCount = @ReprimandCount, Demerit = @Demerit, " + Environment.NewLine +
                                        "       DemeritCount = @DemeritCount, MajorDemerit = @MajorDemerit, MajorDemeritCount = @MajorDemeritCount, " + Environment.NewLine +
                                        "       Commendation = @Commendation, CommendationCount = @CommendationCount, MeritCitation = @MeritCitation, " + Environment.NewLine +
                                        "       MeritCitationCount = @MeritCitationCount, MajorMeritCitation = @MajorMeritCitation, " + Environment.NewLine +
                                        "       MajorMeritCitationCount = @MajorMeritCitationCount, Others = @Others " + Environment.NewLine +
                                        " WHERE (CaseNo = @CaseNo)";
                sdsRPReportDetail.UpdateCommand = vSQLStr_Update;
                sdsRPReportDetail.UpdateParameters.Clear();
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Edit.Trim()));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("CaseComeFrom", DbType.String, (eCaseComeFrom_Edit.Text.Trim() != "") ? eCaseComeFrom_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("CaseType", DbType.String, (eCaseType_Edit.Text.Trim() != "") ? eCaseType_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_Edit.Text.Trim() != "") ? eDepNo_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("EmpNo", DbType.String, (eEmpNo_Edit.Text.Trim() != "") ? eEmpNo_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("EmpName", DbType.String, (eEmpName_Edit.Text.Trim() != "") ? eEmpName_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Title", DbType.String, (eTitle_Edit.Text.Trim() != "") ? eTitle_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("CaseDate", DbType.DateTime, (eCaseDate_Edit.Text.Trim() != "") ? eCaseDate_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Car_ID", DbType.String, (eCar_ID_Edit.Text.Trim() != "") ? eCar_ID_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Position", DbType.String, (ePosition_Edit.Text.Trim() != "") ? ePosition_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("CaseNote", DbType.String, (eCaseNote_Edit.Text.Trim() != "") ? eCaseNote_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("AccordingTerms", DbType.String, (eAccordingTerms_Edit.Text.Trim() != "") ? eAccordingTerms_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Review", DbType.String, (eReview_Edit.Text.Trim() != "") ? eReview_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));
                //sdsRPReportDetail.UpdateParameters.Add(new Parameter("AssignDate", DbType.DateTime, (eAssignDate_Edit.Text.Trim() != "") ? eAssignDate_Edit.Text.Trim() : String.Empty));
                //sdsRPReportDetail.UpdateParameters.Add(new Parameter("AssignDepNo", DbType.String, (eAssignDepNo_Edit.Text.Trim() != "") ? eAssignDepNo_Edit.Text.Trim() : String.Empty));
                //sdsRPReportDetail.UpdateParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan_Edit.Text.Trim() != "") ? eAssignMan_Edit.Text.Trim() : String.Empty));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.DateTime, DateTime.Today.ToShortDateString()));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("GiveBounds", DbType.Boolean, (eGiveBounds_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("BoundsAmount", DbType.Int32, (eBoundsAmount_Edit.Text.Trim() != "") ? eBoundsAmount_Edit.Text.Trim() : "0"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("AskExact", DbType.Boolean, (eAskExact_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("ExactAmount", DbType.Int32, (eExactAmount_Edit.Text.Trim() != "") ? eExactAmount_Edit.Text.Trim() : "0"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Demotion", DbType.Boolean, (eDemotion_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Dismissal", DbType.Boolean, (eDismissal_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Promotion", DbType.Boolean, (ePromotion_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Advice", DbType.Boolean, (eAdvice_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Admonition", DbType.Boolean, (eAdmonition_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Reprimand", DbType.Boolean, (eReprimand_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("ReprimandCount", DbType.Int32, (eReprimandCount_Edit.Text.Trim() != "") ? eReprimandCount_Edit.Text.Trim() : "0"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Demerit", DbType.Boolean, (eDemerit_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("DemeritCount", DbType.Int32, (eDemeritCount_Edit.Text.Trim() != "") ? eDemeritCount_Edit.Text.Trim() : "0"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("MajorDemerit", DbType.Boolean, (eMajorDemerit_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("MajorDemeritCount", DbType.Int32, (eMajorDemeritCount_Edit.Text.Trim() != "") ? eMajorDemeritCount_Edit.Text.Trim() : "0"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Commendation", DbType.Boolean, (eCommendation_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("CommendationCount", DbType.Int32, (eCommendationCount_Edit.Text.Trim() != "") ? eCommendationCount_Edit.Text.Trim() : "0"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("MeritCitation", DbType.Boolean, (eMeritCitation_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("MeritCitationCount", DbType.Int32, (eMeritCitationCount_Edit.Text.Trim() != "") ? eMeritCitationCount_Edit.Text.Trim() : "0"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("MajorMeritCitation", DbType.Boolean, (eMajorMeritCitation_Edit.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("MajorMeritCitationCount", DbType.Int32, (eMajorMeritCitationCount_Edit.Text.Trim() != "") ? eMajorMeritCitationCount_Edit.Text.Trim() : "0"));
                sdsRPReportDetail.UpdateParameters.Add(new Parameter("Others", DbType.Boolean, (eOthers_Edit.Checked == true) ? "true" : "false"));

                sdsRPReportDetail.Update();
                fvRPReportDetail.ChangeMode(FormViewMode.ReadOnly);
                fvRPReportDetail.DataBind();
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        /// <summary>
        /// 新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            try
            {
                string vCaseNo_Temp = "";
                Label eCaseComeFrom_INS = (Label)fvRPReportDetail.FindControl("eCaseComeFrom_INS");
                Label eCaseType_INS = (Label)fvRPReportDetail.FindControl("eCaseType_INS");
                TextBox eDepNo_INS = (TextBox)fvRPReportDetail.FindControl("eDepNo_INS");
                TextBox eEmpNo_INS = (TextBox)fvRPReportDetail.FindControl("eEmpNo_INS");
                Label eEmpName_INS = (Label)fvRPReportDetail.FindControl("eEmpName_INS");
                TextBox eTitle_INS = (TextBox)fvRPReportDetail.FindControl("eTitle_INS");
                TextBox eCaseDate_INS = (TextBox)fvRPReportDetail.FindControl("eCaseDate_INS");
                TextBox eCar_ID_INS = (TextBox)fvRPReportDetail.FindControl("eCar_ID_INS");
                TextBox ePosition_INS = (TextBox)fvRPReportDetail.FindControl("ePosition_INS");
                TextBox eCaseNote_INS = (TextBox)fvRPReportDetail.FindControl("eCaseNote_INS");
                TextBox eAccordingTerms_INS = (TextBox)fvRPReportDetail.FindControl("eAccordingTerms_INS");
                TextBox eReview_INS = (TextBox)fvRPReportDetail.FindControl("eReview_INS");
                TextBox eRemark_INS = (TextBox)fvRPReportDetail.FindControl("eRemark_INS");
                TextBox eAssignDate_INS = (TextBox)fvRPReportDetail.FindControl("eAssignDate_INS");
                DateTime vAssignDate_Temp = DateTime.Parse(eAssignDate_INS.Text.Trim());
                Label eAssignDepNo_INS = (Label)fvRPReportDetail.FindControl("eAssignDepNo_INS");
                TextBox eAssignMan_INS = (TextBox)fvRPReportDetail.FindControl("eAssignMan_INS");
                CheckBox eGiveBounds_INS = (CheckBox)fvRPReportDetail.FindControl("eGiveBounds_INS");
                TextBox eBoundsAmount_INS = (TextBox)fvRPReportDetail.FindControl("eBoundsAmount_INS");
                CheckBox eAskExact_INS = (CheckBox)fvRPReportDetail.FindControl("eAskExact_INS");
                TextBox eExactAmount_INS = (TextBox)fvRPReportDetail.FindControl("eExactAmount_INS");
                CheckBox eDemotion_INS = (CheckBox)fvRPReportDetail.FindControl("eDemotion_INS");
                CheckBox eDismissal_INS = (CheckBox)fvRPReportDetail.FindControl("eDismissal_INS");
                CheckBox ePromotion_INS = (CheckBox)fvRPReportDetail.FindControl("ePromotion_INS");
                CheckBox eAdvice_INS = (CheckBox)fvRPReportDetail.FindControl("eAdvice_INS");
                CheckBox eAdmonition_INS = (CheckBox)fvRPReportDetail.FindControl("eAdmonition_INS");
                CheckBox eReprimand_INS = (CheckBox)fvRPReportDetail.FindControl("eReprimand_INS");
                TextBox eReprimandCount_INS = (TextBox)fvRPReportDetail.FindControl("eReprimandCount_INS");
                CheckBox eDemerit_INS = (CheckBox)fvRPReportDetail.FindControl("eDemerit_INS");
                TextBox eDemeritCount_INS = (TextBox)fvRPReportDetail.FindControl("eDemeritCount_INS");
                CheckBox eMajorDemerit_INS = (CheckBox)fvRPReportDetail.FindControl("eMajorDemerit_INS");
                TextBox eMajorDemeritCount_INS = (TextBox)fvRPReportDetail.FindControl("eMajorDemeritCount_INS");
                CheckBox eCommendation_INS = (CheckBox)fvRPReportDetail.FindControl("eCommendation_INS");
                TextBox eCommendationCount_INS = (TextBox)fvRPReportDetail.FindControl("eCommendationCount_INS");
                CheckBox eMeritCitation_INS = (CheckBox)fvRPReportDetail.FindControl("eMeritCitation_INS");
                TextBox eMeritCitationCount_INS = (TextBox)fvRPReportDetail.FindControl("eMeritCitationCount_INS");
                CheckBox eMajorMeritCitation_INS = (CheckBox)fvRPReportDetail.FindControl("eMajorMeritCitation_INS");
                TextBox eMajorMeritCitationCount_INS = (TextBox)fvRPReportDetail.FindControl("eMajorMeritCitationCount_INS");
                CheckBox eOthers_INS = (CheckBox)fvRPReportDetail.FindControl("eOthers_INS");

                string vAssignYM = (vAssignDate_Temp.Year - 1911).ToString("D3") + vAssignDate_Temp.Month.ToString("D2");
                string vGetMaxItemsStr = "select Max(CaseNo) MaxCaseNo from RP_Report where CaseNo like '" + eAssignDepNo_INS.Text.Trim() + "-" + vAssignYM + "%' ";
                string vLastMaxCaseNo = PF.GetValue(vConnStr, vGetMaxItemsStr, "MaxCaseNo");
                int Items = (vLastMaxCaseNo == "") ? 1 : Int32.Parse(vLastMaxCaseNo.Replace(eAssignDepNo_INS.Text.Trim() + "-" + vAssignYM + "-", "")) + 1;
                vCaseNo_Temp = eAssignDepNo_INS.Text.Trim() + "-" + vAssignYM + "-" + Items.ToString("D4");

                string vSQLStr_INS = "INSERT INTO RP_Report " + Environment.NewLine +
                                     "       (CaseNo, CaseComeFrom, CaseType, DepNo, EmpNo, EmpName, Title, " + Environment.NewLine +
                                     "        CaseDate, Car_ID, Position, CaseNote, AccordingTerms, Review, Remark, " + Environment.NewLine +
                                     "        AssignDate, AssignDepNo, AssignMan, BuDate, BuMan, " + Environment.NewLine +
                                     "        GiveBounds, BoundsAmount, AskExact, ExactAmount, Demotion, Dismissal, Promotion, Advice, Admonition, " + Environment.NewLine +
                                     "        Reprimand, ReprimandCount, Demerit, DemeritCount, MajorDemerit, MajorDemeritCount, " + Environment.NewLine +
                                     "        Commendation, CommendationCount, MeritCitation, MeritCitationCount, MajorMeritCitation, MajorMeritCitationCount, Others) " + Environment.NewLine +
                                     "VALUES (@CaseNo, @CaseComeFrom, @CaseType, @DepNo, @EmpNo, @EmpName, @Title, " + Environment.NewLine +
                                     "        @CaseDate, @Car_ID, @Position, @CaseNote, @AccordingTerms, @Review, @Remark, " + Environment.NewLine +
                                     "        @AssignDate, @AssignDepNo, @AssignMan, @BuDate, @BuMan, " + Environment.NewLine +
                                     "        @GiveBounds, @BoundsAmount, @AskExact, @ExactAmount, @Demotion, @Dismissal, @Promotion, @Advice, @Admonition, " + Environment.NewLine +
                                     "        @Reprimand, @ReprimandCount, @Demerit, @DemeritCount, @MajorDemerit, @MajorDemeritCount, " + Environment.NewLine +
                                     "        @Commendation, @CommendationCount, @MeritCitation, @MeritCitationCount, @MajorMeritCitation, @MajorMeritCitationCount, @Others)";
                sdsRPReportDetail.InsertCommand = vSQLStr_INS;
                sdsRPReportDetail.InsertParameters.Clear();
                sdsRPReportDetail.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("CaseComeFrom", DbType.String, (eCaseComeFrom_INS.Text.Trim() != "") ? eCaseComeFrom_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("CaseType", DbType.String, (eCaseType_INS.Text.Trim() != "") ? eCaseType_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_INS.Text.Trim() != "") ? eDepNo_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("EmpNo", DbType.String, (eEmpNo_INS.Text.Trim() != "") ? eEmpNo_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("EmpName", DbType.String, (eEmpName_INS.Text.Trim() != "") ? eEmpName_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Title", DbType.String, (eTitle_INS.Text.Trim() != "") ? eTitle_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("CaseDate", DbType.DateTime, (eCaseDate_INS.Text.Trim() != "") ? eCaseDate_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Car_ID", DbType.String, (eCar_ID_INS.Text.Trim() != "") ? eCar_ID_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Position", DbType.String, (ePosition_INS.Text.Trim() != "") ? ePosition_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("CaseNote", DbType.String, (eCaseNote_INS.Text.Trim() != "") ? eCaseNote_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("AccordingTerms", DbType.String, (eAccordingTerms_INS.Text.Trim() != "") ? eAccordingTerms_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Review", DbType.String, (eReview_INS.Text.Trim() != "") ? eReview_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark_INS.Text.Trim() != "") ? eRemark_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("AssignDate", DbType.DateTime, (eAssignDate_INS.Text.Trim() != "") ? eAssignDate_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("AssignDepNo", DbType.String, (eAssignDepNo_INS.Text.Trim() != "") ? eAssignDepNo_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan_INS.Text.Trim() != "") ? eAssignMan_INS.Text.Trim() : String.Empty));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("BuDate", DbType.DateTime, DateTime.Today.ToShortDateString()));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("GiveBounds", DbType.Boolean, (eGiveBounds_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("BoundsAmount", DbType.Int32, (eBoundsAmount_INS.Text.Trim() != "") ? eBoundsAmount_INS.Text.Trim() : "0"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("AskExact", DbType.Boolean, (eAskExact_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("ExactAmount", DbType.Int32, (eExactAmount_INS.Text.Trim() != "") ? eExactAmount_INS.Text.Trim() : "0"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Demotion", DbType.Boolean, (eDemotion_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Dismissal", DbType.Boolean, (eDismissal_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Promotion", DbType.Boolean, (ePromotion_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Advice", DbType.Boolean, (eAdvice_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Admonition", DbType.Boolean, (eAdmonition_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Reprimand", DbType.Boolean, (eReprimand_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("ReprimandCount", DbType.Int32, (eReprimandCount_INS.Text.Trim() != "") ? eReprimandCount_INS.Text.Trim() : "0"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Demerit", DbType.Boolean, (eDemerit_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("DemeritCount", DbType.Int32, (eDemeritCount_INS.Text.Trim() != "") ? eDemeritCount_INS.Text.Trim() : "0"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("MajorDemerit", DbType.Boolean, (eMajorDemerit_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("MajorDemeritCount", DbType.Int32, (eMajorDemeritCount_INS.Text.Trim() != "") ? eMajorDemeritCount_INS.Text.Trim() : "0"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Commendation", DbType.Boolean, (eCommendation_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("CommendationCount", DbType.Int32, (eCommendationCount_INS.Text.Trim() != "") ? eCommendationCount_INS.Text.Trim() : "0"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("MeritCitation", DbType.Boolean, (eMeritCitation_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("MeritCitationCount", DbType.Int32, (eMeritCitationCount_INS.Text.Trim() != "") ? eMeritCitationCount_INS.Text.Trim() : "0"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("MajorMeritCitation", DbType.Boolean, (eMajorMeritCitation_INS.Checked == true) ? "true" : "false"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("MajorMeritCitationCount", DbType.Int32, (eMajorMeritCitationCount_INS.Text.Trim() != "") ? eMajorMeritCitationCount_INS.Text.Trim() : "0"));
                sdsRPReportDetail.InsertParameters.Add(new Parameter("Others", DbType.Boolean, (eOthers_INS.Checked == true) ? "true" : "false"));

                sdsRPReportDetail.Insert();
                fvRPReportDetail.ChangeMode(FormViewMode.ReadOnly);
                fvRPReportDetail.DataBind();
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