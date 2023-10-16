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
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class AdviceReport : Page
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

                //從 SESSION 取回使用者帳號資料
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
                    //有登入成功取得使用者帳號
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
                    }

                    //定義畫面上事件日期欄位各自的 onclick 事件
                    //因為畫面日期的定義是民國年，所以對應的日期選擇視窗是 InputDate_ChineseYears.aspx
                    string vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_S_Search.ClientID; 
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        //不是從 PostBack 觸發頁面更新，表示是第一次開啟頁面
                        plSearch.Visible = true;
                        plReport.Visible = false;
                    }
                    else
                    {
                        //從 PostBack 觸發頁面更新
                        OpenData();
                    }
                }
                else
                {
                    //跳回首頁
                    Response.Redirect("~/default.aspx");
                }
            }
            else
            {
                //跳回首頁
                Response.Redirect("~/default.aspx");
            }
        }

        /// <summary>
        /// 取回資料查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            //因為頁面的時間格式已經定義成民國年，
            //所以日期元件如果直接用 ToString() 函式取回的日期字串也會是民國年
            //可是要送進 SQL 查詢字串的日期字串只能是西元年格式
            //所以要先用 Year 命令取回西元年，再跟 ToString() 取回的月/日字串組合出正確的日期字串
            DateTime vCaseDate_S = (eCaseDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_S_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_S_Str = vCaseDate_S.Year.ToString() + "/" + vCaseDate_S.ToString("MM/dd");
            DateTime vCaseDate_E = (eCaseDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_E_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_E_Str = vCaseDate_E.Year.ToString() + "/" + vCaseDate_E.ToString("MM/dd");
            //===================================================================================================================================
            string vReturnStr = "";
            string vWStr_CaseDate = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and a.CaseDate between '" + vCaseDate_S_Str + "' and '" + vCaseDate_E_Str + "' " :
                                    ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? "   and a.CaseDate = '" + vCaseDate_S_Str + "' " :
                                    ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and a.CaseDate = '" + vCaseDate_E_Str + "' " : "";
            string vWStr_DepNo = ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   and a.DepNo between '" + eDepNo_S_Search.Text.Trim() + "' and '" + eDepNo_E_Search.Text.Trim() + "' " :
                                 ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? "   and a.DepNo = '" + eDepNo_S_Search.Text.Trim() + "' " :
                                 ((eDepNo_S_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? "   and a.DepNo = '" + eDepNo_E_Search.Text.Trim() + "' " : "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and a.EmpNo = '" + eEmpNo_Search.Text.Trim() + "' " : "";
            vReturnStr = "SELECT CaseNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, " + Environment.NewLine +
                         "       EmpNo, EmpName, Title, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = a.Title)) AS Title_C, " + Environment.NewLine +
                         "       CaseDate, CaseTime, Position, CaseNote, Remark, DenialReason, BuDate, " + Environment.NewLine +
                         "       BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, " + Environment.NewLine +
                         "       ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, " + Environment.NewLine +
                         "       a.SourceType, " + Environment.NewLine +
                         "       a.IsClose, (SELECT CLASSTXT FROM DBDICB WHERE CLASSNO = a.IsClose AND FKEY = '員工疏失勸導單  AdviceReport    IsClose') IsClose_C " + Environment.NewLine +
                         "  FROM AdviceReport AS a " + Environment.NewLine +
                         " WHERE 1 = 1 " + Environment.NewLine +
                         vWStr_CaseDate + vWStr_DepNo + vWStr_EmpNo +
                         " ORDER BY CaseNo DESC ";
            return vReturnStr;
        }

        /// <summary>
        /// 取回列印勸導單用的資料查詢字串
        /// </summary>
        /// <param name="fCaseNo">勸導單單號</param>
        /// <returns></returns>
        private string GetSelectStr_Print(string fCaseNo)
        {
            string vReturnStr = "";
            vReturnStr = "SELECT CaseNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = a.DepNo)) AS DepName, " + Environment.NewLine +
                         "       EmpNo, EmpName, Title, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = a.Title)) AS Title_C, " + Environment.NewLine +
                         "       CaseDate, CaseTime, Position, CaseNote, Remark, DenialReason, BuDate, " + Environment.NewLine +
                         "       BuMan, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = a.BuMan)) AS BuManName, " + Environment.NewLine +
                         "       ModifyDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = a.ModifyMan)) AS ModifyManName, " + Environment.NewLine +
                         "       a.SourceType, " + Environment.NewLine +
                         "       a.IsClose, (SELECT CLASSTXT FROM DBDICB WHERE CLASSNO = a.IsClose AND FKEY = '員工疏失勸導單  AdviceReport    IsClose') IsClose_C " + Environment.NewLine +
                         "  FROM AdviceReport AS a " + Environment.NewLine +
                         " WHERE CaseNo = '" + fCaseNo + "' ";
            return vReturnStr;
        }

        /// <summary>
        /// 查詢資料
        /// </summary>
        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            sdsAdviceReport_List.SelectCommand = vSelStr;
            gridAdviceReportList.DataSourceID = ""; //清除 DataSourceID 的設定，確保不會跟後面的 DataSource 設定衝突
            gridAdviceReportList.DataSource = sdsAdviceReport_List;
            gridAdviceReportList.DataBind(); //進行資料繫結
        }

        protected void eDepNo_S_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            string vDepNo_Temp = eDepNo_S_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_S_Search.Text = vDepNo_Temp.Trim();
            eDepName_S_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDepNo_E_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            string vDepNo_Temp = eDepNo_E_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_E_Search.Text = vDepNo_Temp.Trim();
            eDepName_E_Search.Text = vDepName_Temp.Trim();
        }

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            string vEmpNo_Temp = eEmpNo_Search.Text.Trim();
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
            string vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp;
                vSQLStr = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Temp + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eEmpNo_Search.Text = vEmpNo_Temp;
            eEmpName_Search.Text = vEmpName_Temp;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void gridAdviceReportList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridAdviceReportList.PageIndex = e.NewPageIndex;
            gridAdviceReportList.DataBind();
        }

        protected void fvAdviceReportDetail_DataBound(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            string vSQLStr_Temp = "";
            string vCaseDateURL = "";
            string vCaseDateScript = "";
            switch (fvAdviceReportDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    //資料顯示模式
                    break;

                case FormViewMode.Edit:
                    //資料編輯模式
                    Label eModifyDate_Edit = (Label)fvAdviceReportDetail.FindControl("eModifyDate_Edit");
                    Label eModifyMan_Edit = (Label)fvAdviceReportDetail.FindControl("eModifyMan_Edit");
                    Label eModifyManNAme_Edit = (Label)fvAdviceReportDetail.FindControl("eModifyManName_Edit");
                    TextBox eCaseDate_Edit = (TextBox)fvAdviceReportDetail.FindControl("eCaseDate_Edit");
                    DropDownList ddlIsClose_Edit = (DropDownList)fvAdviceReportDetail.FindControl("ddlIsClose_Edit");
                    Label eIsClose_Edit = (Label)fvAdviceReportDetail.FindControl("eIsClose_Edit");

                    using (SqlConnection connTemp_Edit = new SqlConnection(vConnStr))
                    {
                        //填入 "處理進度" 的下拉選單選項
                        vSQLStr_Temp = "select ClassNo, CLassTxt from DBDICB where FKey = '員工疏失勸導單  AdviceReport    IsClose' order by ClassNo";
                        SqlCommand cmdTemp_Edit = new SqlCommand(vSQLStr_Temp, connTemp_Edit);
                        connTemp_Edit.Open();
                        SqlDataReader drTemp_Edit = cmdTemp_Edit.ExecuteReader();
                        ddlIsClose_Edit.Items.Clear();
                        while (drTemp_Edit.Read())
                        {
                            ddlIsClose_Edit.Items.Add(new ListItem(drTemp_Edit["ClassTxt"].ToString().Trim(), drTemp_Edit["ClassNo"].ToString().Trim()));
                        }
                    }

                    eModifyDate_Edit.Text = PF.GetChinsesDate(DateTime.Today);
                    eModifyMan_Edit.Text = vLoginID;
                    vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vLoginID + "' ";
                    eModifyManNAme_Edit.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                    //定義 "事件日期" 的觸發事件
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Edit.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Edit.Attributes["onClick"] = vCaseDateScript;
                    //=================================================================
                    ddlIsClose_Edit.SelectedIndex = (eIsClose_Edit.Text.Trim() == "") ? 0 : ddlIsClose_Edit.Items.IndexOf(ddlIsClose_Edit.Items.FindByValue(eIsClose_Edit.Text.Trim()));
                    ddlIsClose_Edit.Enabled = ((vLoginDepNo == "06") || (vLoginDepNo == "09"));
                    break;

                case FormViewMode.Insert:
                    //資料新增模式
                    Label eBuDate_INS = (Label)fvAdviceReportDetail.FindControl("eBuDate_INS");
                    Label eBuMan_INS = (Label)fvAdviceReportDetail.FindControl("eBuMan_INS");
                    Label eBuManName_INS = (Label)fvAdviceReportDetail.FindControl("eBuManNAme_INS");
                    TextBox eCaseDate_INS = (TextBox)fvAdviceReportDetail.FindControl("eCaseDate_INS");
                    DropDownList ddlIsClose_INS = (DropDownList)fvAdviceReportDetail.FindControl("ddlIsClose_INS");
                    Label eIsClose_INS = (Label)fvAdviceReportDetail.FindControl("eIsClose_INS");

                    using (SqlConnection connTemp_INS = new SqlConnection(vConnStr))
                    {
                        //填入 "處理進度" 的下拉選單選項
                        vSQLStr_Temp = "select ClassNo, CLassTxt from DBDICB where FKey = '員工疏失勸導單  AdviceReport    IsClose' order by ClassNo";
                        SqlCommand cmdTemp_INS = new SqlCommand(vSQLStr_Temp, connTemp_INS);
                        connTemp_INS.Open();
                        SqlDataReader drTemp_INS = cmdTemp_INS.ExecuteReader();
                        ddlIsClose_INS.Items.Clear();
                        while (drTemp_INS.Read())
                        {
                            ddlIsClose_INS.Items.Add(new ListItem(drTemp_INS["ClassTxt"].ToString().Trim(), drTemp_INS["ClassNo"].ToString().Trim()));
                        }
                    }

                    eBuDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                    eBuMan_INS.Text = vLoginID;
                    vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vLoginID + "' ";
                    eBuManName_INS.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                    //定義 "事件日期" 的觸發事件
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_INS.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_INS.Attributes["onClick"] = vCaseDateScript;
                    //==================================================================================================================================
                    ddlIsClose_INS.SelectedIndex = 0;
                    eIsClose_INS.Text = "00";
                    ddlIsClose_INS.Enabled = ((vLoginDepNo == "06") || (vLoginDepNo == "09"));
                    break;
            }
        }

        protected void eEmpNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            TextBox eEmpNo_Edit = (TextBox)fvAdviceReportDetail.FindControl("eEmpNo_Edit");
            Label eEmpName_Edit = (Label)fvAdviceReportDetail.FindControl("eEmpName_Edit");
            TextBox eDepNo_Edit = (TextBox)fvAdviceReportDetail.FindControl("eDepNo_Edit");
            Label eDepName_Edit = (Label)fvAdviceReportDetail.FindControl("eDepName_Edit");
            TextBox eTitle_Edit = (TextBox)fvAdviceReportDetail.FindControl("eTitle_Edit");
            Label eTitle_C_Edit = (Label)fvAdviceReportDetail.FindControl("eTitle_C_Edit");

            string vEmpNo_Temp = eEmpNo_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
            string vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            string vEmpDataStr = "";
            string[] vaEmpData;

            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp;
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Temp + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            if (vEmpNo_Temp != "")
            {
                vSQLStr_Temp = "select e.DepNo + ',' + (select [Name] from Department where DepNo = e.DepNo) + ',' + " + Environment.NewLine +
                               "       e.Title + ',' + (select ClassTxt from DBDICB where ClassNo = e.Title and FKey = '員工資料        EMPLOYEE        TITLE') EmpData " + Environment.NewLine +
                               "  from Employee e " + Environment.NewLine +
                               " where EmpNo = '" + vEmpNo_Temp.Trim() + "' ";
                vEmpDataStr = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpData");
                vaEmpData = vEmpDataStr.Split(',');
                if (vaEmpData.Length > 0)
                {
                    eDepNo_Edit.Text = vaEmpData[0].Trim();
                    eDepName_Edit.Text = vaEmpData[1].Trim();
                    eTitle_Edit.Text = vaEmpData[2].Trim();
                    eTitle_C_Edit.Text = vaEmpData[3].Trim();
                }
            }
            eEmpNo_Edit.Text = vEmpNo_Temp;
            eEmpName_Edit.Text = vEmpName_Temp;
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            TextBox eDepNo_Edit = (TextBox)fvAdviceReportDetail.FindControl("eDepNo_Edit");
            Label eDepName_Edit = (Label)fvAdviceReportDetail.FindControl("eDepName_Edit");
            string vDepNo_Temp = eDepNo_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vLoginDepNo;
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Edit.Text = vDepNo_Temp;
            eDepName_Edit.Text = vDepName_Temp;
        }

        protected void eTitle_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            TextBox eTitle_Edit = (TextBox)fvAdviceReportDetail.FindControl("eTitle_Edit");
            Label eTitle_C_Edit = (Label)fvAdviceReportDetail.FindControl("eTitle_C_Edit");
            string vTitle_Temp = eTitle_Edit.Text.Trim();
            string vSQLStr_Temp = "select CLassTxt from DBDICB where ClassNo = '" + vTitle_Temp + "' and FKey = '員工資料        EMPLOYEE        TITLE' ";
            string vTitle_C_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "ClassTxt");
            if (vTitle_C_Temp == "")
            {
                vTitle_C_Temp = vTitle_Temp;
                vSQLStr_Temp = "select CLassNo from DBDICB where ClassTXT = '" + vTitle_C_Temp + "' and FKey = '員工資料        EMPLOYEE        TITLE' ";
                vTitle_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "ClassNo");
            }
            eTitle_Edit.Text = vTitle_Temp;
            eTitle_C_Edit.Text = vTitle_C_Temp;
        }

        protected void eEmpNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            TextBox eEmpNo_INS = (TextBox)fvAdviceReportDetail.FindControl("eEmpNo_INS");
            Label eEmpName_INS = (Label)fvAdviceReportDetail.FindControl("eEmpName_INS");
            TextBox eDepNo_INS = (TextBox)fvAdviceReportDetail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvAdviceReportDetail.FindControl("eDepName_INS");
            TextBox eTitle_INS = (TextBox)fvAdviceReportDetail.FindControl("eTitle_INS");
            Label eTitle_C_INS = (Label)fvAdviceReportDetail.FindControl("eTitle_C_INS");

            string vEmpNo_Temp = eEmpNo_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
            string vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            string vEmpDataStr = "";
            string[] vaEmpData;

            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp;
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Temp + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            if (vEmpNo_Temp != "")
            {
                vSQLStr_Temp = "select e.DepNo + ',' + (select [Name] from Department where DepNo = e.DepNo) + ',' + " + Environment.NewLine +
                               "       e.Title + ',' + (select ClassTxt from DBDICB where ClassNo = e.Title and FKey = '員工資料        EMPLOYEE        TITLE') EmpData " + Environment.NewLine +
                               "  from Employee e " + Environment.NewLine +
                               " where EmpNo = '" + vEmpNo_Temp.Trim() + "' ";
                vEmpDataStr = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpData");
                vaEmpData = vEmpDataStr.Split(',');
                if (vaEmpData.Length > 0)
                {
                    eDepNo_INS.Text = vaEmpData[0].Trim();
                    eDepName_INS.Text = vaEmpData[1].Trim();
                    eTitle_INS.Text = vaEmpData[2].Trim();
                    eTitle_C_INS.Text = vaEmpData[3].Trim();
                }
            }
            eEmpNo_INS.Text = vEmpNo_Temp;
            eEmpName_INS.Text = vEmpName_Temp;
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            TextBox eDepNo_INS = (TextBox)fvAdviceReportDetail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvAdviceReportDetail.FindControl("eDepName_INS");
            string vDepNo_Temp = eDepNo_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vLoginDepNo;
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_INS.Text = vDepNo_Temp;
            eDepName_INS.Text = vDepName_Temp;
        }

        protected void eTitle_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            TextBox eTitle_INS = (TextBox)fvAdviceReportDetail.FindControl("eTitle_INS");
            Label eTitle_C_INS = (Label)fvAdviceReportDetail.FindControl("eTitle_C_INS");
            string vTitle_Temp = eTitle_INS.Text.Trim();
            string vSQLStr_Temp = "select CLassTxt from DBDICB where ClassNo = '" + vTitle_Temp + "' and FKey = '員工資料        EMPLOYEE        TITLE' ";
            string vTitle_C_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "ClassTxt");
            if (vTitle_C_Temp == "")
            {
                vTitle_C_Temp = vTitle_Temp;
                vSQLStr_Temp = "select CLassNo from DBDICB where ClassTXT = '" + vTitle_C_Temp + "' and FKey = '員工資料        EMPLOYEE        TITLE' ";
                vTitle_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "ClassNo");
            }
            eTitle_INS.Text = vTitle_Temp;
            eTitle_C_INS.Text = vTitle_C_Temp;
        }

        protected void bbPrint_List_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            Label eCaseNo_List = (Label)fvAdviceReportDetail.FindControl("eCaseNo_List");
            string vCaseNo_List = eCaseNo_List.Text.Trim();
            string vSelStr = GetSelectStr_Print(vCaseNo_List);
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintPoint = new SqlDataAdapter(vSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintPoint = new DataTable("AdviceReportP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("AdviceReportP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\AdviceReportP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "桃園客運員工疏失（違規）改善勸導單"));
                    rvPrint.LocalReport.Refresh();
                    plReport.Visible = true;
                    plShowData.Visible = false;
                }
            }
        }

        /// <summary>
        /// 刪除勸導單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDel_List_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            Label eCaseNo_List = (Label)fvAdviceReportDetail.FindControl("eCaseNo_List");
            Label eSourceType_List = (Label)fvAdviceReportDetail.FindControl("eSourceType_List");
            string vSourceType = (eSourceType_List != null) ? eSourceType_List.Text.Trim() : "";
            string vSQLStr_Temp = "select isnull(MAX(Items), 0) MaxItems from AdviceReport_History where CaseNo = '" + eCaseNo_List.Text.Trim() + "' ";
            int vMaxItems = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems")) + 1;
            string vCaseNoItem = eCaseNo_List.Text.Trim() + vMaxItems.ToString("D4");
            vSQLStr_Temp = "insert into AdviceReport_History (CaseNoItem, Items, ModifyMode, CaseNo, DepNo, EmpNo, EmpName, Title, " + Environment.NewLine +
                           "                                  CaseDate, CaseTime, Position, CaseNote, Remark, DenialReason, " + Environment.NewLine +
                           "                                  BuDate, BuMan, ModifyDate, ModifyMan, SourceType, IsClose) " + Environment.NewLine +
                           "select '" + vCaseNoItem + "', '" + vMaxItems.ToString("D4") + "', 'DEL', " + Environment.NewLine +
                           "       CaseNo, DepNo, EmpNo, EmpName, Title, CaseDate, CaseTime, Position, " + Environment.NewLine +
                           "       CaseNote, Remark, DenialReason, BuDate, BuMan, ModifyDate, ModifyMan, SourceType, IsClose " + Environment.NewLine +
                           "  from AdviceReport " + Environment.NewLine +
                           " where CaseNo = '" + eCaseNo_List.Text.Trim() + "' ";
            try
            {
                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                vSQLStr_Temp = (vSourceType.Trim() == "H") ? "update HDDCheckReport set IsPassed = 0, AdviceReportNo = NULL where AdviceReportNo = '" + eCaseNo_List.Text.Trim() + "' " :
                               (vSourceType.Trim() == "T") ? "update TicketCheckreport set IsPassed = 0, RP_ReportNo = NULL where RP_ReportNo = '" + eCaseNo_List.Text.Trim() + "' " :
                               "";
                if (vSQLStr_Temp.Trim() != "")
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                }
                sdsAdviceReport_Detail.DeleteParameters.Clear();
                sdsAdviceReport_Detail.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, eCaseNo_List.Text.Trim()));
                sdsAdviceReport_Detail.Delete();
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        protected void sdsAdviceReport_Detail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridAdviceReportList.DataBind();
            }
        }

        protected void sdsAdviceReport_Detail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridAdviceReportList.DataBind();
            }
        }

        protected void sdsAdviceReport_Detail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridAdviceReportList.DataBind();
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plReport.Visible = false;
        }

        /// <summary>
        /// 匯出EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExcel_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
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
            string vFileName = "員工疏失勸導清冊";
            DateTime vBuDate;

            using (SqlConnection connExcel = new SqlConnection(vConnStr))
            {
                string vSelStr = GetSelectStr();
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
                        vHeaderText = (drExcel.GetName(i) == "CaseNo") ? "單號" :
                                      (drExcel.GetName(i) == "DepNo") ? "" :
                                      (drExcel.GetName(i) == "DepName") ? "站別" :
                                      (drExcel.GetName(i) == "EmpNo") ? "" :
                                      (drExcel.GetName(i) == "EmpName") ? "員工姓名" :
                                      (drExcel.GetName(i) == "Title") ? "" :
                                      (drExcel.GetName(i) == "Title_C") ? "職稱" :
                                      (drExcel.GetName(i) == "CaseDate") ? "事件日期" :
                                      (drExcel.GetName(i) == "CaseTime") ? "事件時間" :
                                      (drExcel.GetName(i) == "Position") ? "地點" :
                                      (drExcel.GetName(i) == "CaseNote") ? "勸導事由" :
                                      (drExcel.GetName(i) == "Remark") ? "補充說明" :
                                      (drExcel.GetName(i) == "DenialReason") ? "拒簽原因" :
                                      (drExcel.GetName(i) == "BuDate") ? "開單日期" :
                                      (drExcel.GetName(i) == "BuMan") ? "" :
                                      (drExcel.GetName(i) == "BuManName") ? "開單人員" :
                                      (drExcel.GetName(i) == "ModifyDate") ? "異動日期" :
                                      (drExcel.GetName(i) == "ModifyMan") ? "" :
                                      (drExcel.GetName(i) == "ModifyManName") ? "異動人員" :
                                      (drExcel.GetName(i) == "SourceType") ? "來源 (H:硬碟查核 T:查票工作)" : drExcel.GetName(i);
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
                            if (((drExcel.GetName(i) == "DeductionDate") ||
                                 (drExcel.GetName(i) == "BuildDate")) && (drExcel[i].ToString() != ""))
                            {
                                string vTempStr = drExcel[i].ToString();
                                vBuDate = DateTime.Parse(drExcel[i].ToString());
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd"));
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                            else if ((drExcel.GetName(i) == "AnecdotalResRatio") && (drExcel[i].ToString() != ""))
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(double.Parse(drExcel[i].ToString()));
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
                            string vRecordDepNoStr = ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() != "")) ? eDepNo_S_Search.Text.Trim() + "~" + eDepNo_E_Search.Text.Trim() :
                                                     ((eDepNo_S_Search.Text.Trim() != "") && (eDepNo_E_Search.Text.Trim() == "")) ? eDepNo_S_Search.Text.Trim() :
                                                     ((eDepNo_S_Search.Text.Trim() == "") && (eDepNo_E_Search.Text.Trim() != "")) ? eDepNo_E_Search.Text.Trim() : "全部";
                            string vRecordEmployeeStr = (eEmpNo_Search.Text.Trim() != "") ? eEmpNo_Search.Text.Trim() : "全部";
                            string vRecordDateStr = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? eCaseDate_S_Search.Text.Trim() + "~" + eCaseDate_E_Search.Text.Trim() :
                                                    ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? eCaseDate_S_Search.Text.Trim() :
                                                    ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? eCaseDate_E_Search.Text.Trim() : "不指定日期";
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "AdviceReport.aspx" + Environment.NewLine +
                                                 "事件日期：" + vRecordDateStr + Environment.NewLine +
                                                 "單位：" + vRecordDepNoStr + Environment.NewLine +
                                                 "員工：" + vRecordEmployeeStr;
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
            }
        }

        protected void ddlIsClose_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlIsClose_Edit = (DropDownList)fvAdviceReportDetail.FindControl("ddlIsClose_Edit");
            if (ddlIsClose_Edit != null)
            {
                Label eIsClose_Edit = (Label)fvAdviceReportDetail.FindControl("eIsClose_Edit");
                eIsClose_Edit.Text = ddlIsClose_Edit.SelectedValue.Trim();
            }
        }

        protected void ddlIsClose_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlIsClose_INS = (DropDownList)fvAdviceReportDetail.FindControl("ddlIsClose_INS");
            if (ddlIsClose_INS != null)
            {
                Label eIsClose_INS = (Label)fvAdviceReportDetail.FindControl("eIsClose_INS");
                eIsClose_INS.Text = ddlIsClose_INS.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 修改存檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            try
            {
                Label eCaseNo_Edit = (Label)fvAdviceReportDetail.FindControl("eCaseNo_Edit");
                if (eCaseNo_Edit != null)
                {
                    string vCaseNo_Temp = eCaseNo_Edit.Text.Trim();
                    //複製一份到異動檔
                    string vSQLStr_Temp = "select isnull(MAX(Items), 0) MaxItems from AdviceReport_History where CaseNo = '" + vCaseNo_Temp + "' ";
                    int vMaxItems = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems")) + 1;
                    string vCaseNoItem = vCaseNo_Temp + vMaxItems.ToString("D4");
                    vSQLStr_Temp = "insert into AdviceReport_History (CaseNoItem, Items, ModifyMode, CaseNo, DepNo, EmpNo, EmpName, Title, " + Environment.NewLine +
                                   "                                  CaseDate, CaseTime, Position, CaseNote, Remark, DenialReason, " + Environment.NewLine +
                                   "                                  BuDate, BuMan, ModifyDate, ModifyMan, SourceType, IsClose) " + Environment.NewLine +
                                   "select '" + vCaseNoItem + "', '" + vMaxItems.ToString("D4") + "', 'EDIT', " + Environment.NewLine +
                                   "       CaseNo, DepNo, EmpNo, EmpName, Title, CaseDate, CaseTime, Position, " + Environment.NewLine +
                                   "       CaseNote, Remark, DenialReason, BuDate, BuMan, ModifyDate, ModifyMan, SourceType, IsClose " + Environment.NewLine +
                                   "  from AdviceReport " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";

                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //開始寫入修改內容
                    TextBox eDepNo_Edit = (TextBox)fvAdviceReportDetail.FindControl("eDepNo_Edit");
                    TextBox eEmpNo_Edit = (TextBox)fvAdviceReportDetail.FindControl("eEmpNo_Edit");
                    Label eEmpName_Edit = (Label)fvAdviceReportDetail.FindControl("eEmpName_Edit");
                    TextBox eTitle_Edit = (TextBox)fvAdviceReportDetail.FindControl("eTitle_Edit");
                    TextBox eCaseDate_Edit = (TextBox)fvAdviceReportDetail.FindControl("eCaseDAte_Edit");
                    TextBox eCaseTime_Edit = (TextBox)fvAdviceReportDetail.FindControl("eCaseTime_Edit");
                    TextBox ePosition_Edit = (TextBox)fvAdviceReportDetail.FindControl("ePosition_Edit");
                    TextBox eCaseNote_Edit = (TextBox)fvAdviceReportDetail.FindControl("eCaseNote_Edit");
                    TextBox eRemark_Edit = (TextBox)fvAdviceReportDetail.FindControl("eRemark_Edit");
                    TextBox eDenialReason_Edit = (TextBox)fvAdviceReportDetail.FindControl("eDenialReason_Edit");
                    Label eIsClose_Edit = (Label)fvAdviceReportDetail.FindControl("eIsClose_Edit");

                    sdsAdviceReport_Detail.UpdateCommand = "UPDATE AdviceReport " + Environment.NewLine +
                                                           "   SET DepNo = @DepNo, EmpNo = @EmpNo, EmpName = @EmpName, Title = @Title, " + Environment.NewLine +
                                                           "       CaseDate = @CaseDate, CaseTime = @CaseTime, Position = @Position, " + Environment.NewLine +
                                                           "       CaseNote = @CaseNote, Remark = @Remark, DenialReason = @DenialReason, " + Environment.NewLine +
                                                           "       ModifyDate = @ModifyDate, ModifyMan = @ModifyMan, IsClose = @IsClose " + Environment.NewLine +
                                                           " WHERE (CaseNo = @CaseNo)";
                    sdsAdviceReport_Detail.UpdateParameters.Clear();
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_Edit.Text.Trim() != "") ? eDepNo_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("EmpNo", DbType.String, (eEmpNo_Edit.Text.Trim() != "") ? eEmpNo_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("EmpName", DbType.String, (eEmpName_Edit.Text.Trim() != "") ? eEmpName_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("Title", DbType.String, (eTitle_Edit.Text.Trim() != "") ? eTitle_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("CaseDate", DbType.DateTime, (eCaseDate_Edit.Text.Trim() != "") ? eCaseDate_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("CaseTime", DbType.String, (eCaseTime_Edit.Text.Trim() != "") ? eCaseTime_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("Position", DbType.String, (ePosition_Edit.Text.Trim() != "") ? ePosition_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("CaseNote", DbType.String, (eCaseNote_Edit.Text.Trim() != "") ? eCaseNote_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("DenialReason", DbType.String, (eDenialReason_Edit.Text.Trim() != "") ? eDenialReason_Edit.Text.Trim() : String.Empty));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.DateTime, DateTime.Today.ToShortDateString()));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("IsClose", DbType.String, (eIsClose_Edit.Text.Trim() != "") ? eIsClose_Edit.Text.Trim() : "00"));
                    sdsAdviceReport_Detail.UpdateParameters.Add(new Parameter("CaseNo", DbType.String, eCaseNo_Edit.Text.Trim()));

                    sdsAdviceReport_Detail.Update();
                    fvAdviceReportDetail.ChangeMode(FormViewMode.ReadOnly);
                    fvAdviceReportDetail.DataBind();
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

        /// <summary>
        /// 新增存檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
            }
            try
            {
                TextBox eDepNo_INS = (TextBox)fvAdviceReportDetail.FindControl("eDepNo_INS");
                TextBox eEmpNo_INS = (TextBox)fvAdviceReportDetail.FindControl("eEmpNo_INS");
                Label eEmpName_INS = (Label)fvAdviceReportDetail.FindControl("eEmpName_INS");
                TextBox eTitle_INS = (TextBox)fvAdviceReportDetail.FindControl("eTitle_INS");
                TextBox eCaseDate_INS = (TextBox)fvAdviceReportDetail.FindControl("eCaseDate_INS");
                DateTime vCaseDate_Temp = DateTime.Parse(eCaseDate_INS.Text.Trim());
                TextBox eCaseTime_INS = (TextBox)fvAdviceReportDetail.FindControl("eCaseTime_INS");
                TextBox ePosition_INS = (TextBox)fvAdviceReportDetail.FindControl("ePosition_INS");
                TextBox eCaseNote_INS = (TextBox)fvAdviceReportDetail.FindControl("eCaseNote_INS");
                TextBox eRemark_INS = (TextBox)fvAdviceReportDetail.FindControl("eRemark_INS");
                TextBox eDenialReason_INS = (TextBox)fvAdviceReportDetail.FindControl("eDenialReason_INS");
                Label eBuDate_INS = (Label)fvAdviceReportDetail.FindControl("eBuDate_INS");
                DateTime vBuDate_Temp = DateTime.Parse(eBuDate_INS.Text.Trim());
                Label eIsClose_INS = (Label)fvAdviceReportDetail.FindControl("eIsCLose_INS");

                string vCaseYear = (eCaseDate_INS.Text.Trim() != "") ?
                                   (vCaseDate_Temp.Year - 1911).ToString() + "A" :
                                   (DateTime.Today.Year - 1911).ToString() + "A";
                string vSQLStr_Temp = "select MAX(CaseNo) MaxCaseNo from AdviceReport where CaseNo like '" + vCaseYear + "%' " +
                                      " group by CaseNo ";
                string vMaxCaseNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxCaseNo");
                int vIndex = (vMaxCaseNo != "") ? Int32.Parse(vMaxCaseNo.Replace(vCaseYear, "")) + 1 : 1;
                string vCaseNo_Temp = vCaseYear + vIndex.ToString("D4");

                sdsAdviceReport_Detail.InsertCommand = "INSERT INTO AdviceReport " + Environment.NewLine +
                                                       "       (CaseNo, DepNo, EmpNo, EmpName, Title, CaseDate, CaseTime, Position, " + Environment.NewLine +
                                                       "        CaseNote, Remark, DenialReason, BuDate, BuMan, IsClose) " + Environment.NewLine +
                                                       "VALUES (@CaseNo, @DepNo, @EmpNo, @EmpName, @Title, @CaseDate, @CaseTime, @Position, " + Environment.NewLine +
                                                       "        @CaseNote, @Remark, @DenialReason, @BuDate, @BuMan, @IsClose)";
                sdsAdviceReport_Detail.InsertParameters.Clear();
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_INS.Text.Trim() != "") ? eDepNo_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("EmpNo", DbType.String, (eEmpNo_INS.Text.Trim() != "") ? eEmpNo_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("EmpName", DbType.String, (eEmpName_INS.Text.Trim() != "") ? eEmpName_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("Title", DbType.String, (eTitle_INS.Text.Trim() != "") ? eTitle_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("CaseDate", DbType.DateTime, (eCaseDate_INS.Text.Trim() != "") ? eCaseDate_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("CaseTime", DbType.String, (eCaseTime_INS.Text.Trim() != "") ? eCaseTime_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("Position", DbType.String, (ePosition_INS.Text.Trim() != "") ? ePosition_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("CaseNote", DbType.String, (eCaseNote_INS.Text.Trim() != "") ? eCaseNote_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark_INS.Text.Trim() != "") ? eRemark_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("DenialReason", DbType.String, (eDenialReason_INS.Text.Trim() != "") ? eDenialReason_INS.Text.Trim() : String.Empty));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.DateTime, DateTime.Today.ToShortDateString()));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                sdsAdviceReport_Detail.InsertParameters.Add(new Parameter("IsClose", DbType.String, (eIsClose_INS.Text.Trim() != "") ? eIsClose_INS.Text.Trim() : "00"));

                sdsAdviceReport_Detail.Insert();
                fvAdviceReportDetail.ChangeMode(FormViewMode.ReadOnly);
                fvAdviceReportDetail.DataBind();
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