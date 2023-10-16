using Amaterasu_Function;
using System;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ViolationCaseDetail : System.Web.UI.Page
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
        private string vStationNo = "";
        private DateTime vToday;

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

                UnobtrusiveValidationMode = System.Web.UI.UnobtrusiveValidationMode.None;

                if (vLoginID != "")
                {
                    vToday = DateTime.Today;
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vViolationDateURL_Search = "InputDate_ChineseYears.aspx?TextboxID=" + eViolationDate_Start_Search.ClientID;
                    string vViolationDateScript_Search = "window.open('" + vViolationDateURL_Search + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eViolationDate_Start_Search.Attributes["onClick"] = vViolationDateScript_Search;

                    vViolationDateURL_Search = "InputDate_ChineseYears.aspx?TextboxID=" + eViolationDate_End_Search.ClientID;
                    vViolationDateScript_Search = "window.open('" + vViolationDateURL_Search + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eViolationDate_End_Search.Attributes["onClick"] = vViolationDateScript_Search;

                    if (!IsPostBack)
                    {
                        vStationNo = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vLoginDepNo + "' and InSHReport = 'V'", "DepNo");

                        eDepNo_Start_Search.Text = vStationNo;
                        eDepNo_End_Search.Text = "";
                        eDepNo_Start_Search.Enabled = (vStationNo == "");
                        eDepNo_End_Search.Enabled = (vStationNo == "");
                    }
                    else
                    {
                        DataSourceBinded();
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

        private void DataSourceBinded()
        {
            string vSQLStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "SELECT CaseNo, CaseType, " + Environment.NewLine +
                      "       (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = ViolationCase.CaseType) AND (FKEY = '違規記錄檔      ViolationCase   CaseType')) AS CaseType_C, " + Environment.NewLine +
                      "       BuildDate, BuildManName, LinesNo, DepName, Car_ID, DriverName, TicketTitle, PenaltyDep, FineAmount, ViolationDate " + Environment.NewLine +
                      "  FROM ViolationCase " + Environment.NewLine +
                      " WHERE (1 = 1)";
            string vWStr_CaseType = (eCaseType_Search.Text.Trim() != "") ? " AND CaseType = '" + eCaseType_Search.Text.Trim() + "' " : "";
            string vWStr_DepNo = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? " AND DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " :
                                 ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? " AND DepNo like '" + eDepNo_Start_Search.Text.Trim() + "%' " :
                                 ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? " AND DepNo like '" + eDepNo_End_Search.Text.Trim() + "%' " : "";
            string vWStr_CarID = ((eCarID_Start_Search.Text.Trim() != "") && (eCarID_End_Search.Text.Trim() != "")) ? " AND Car_ID between '" + eCarID_Start_Search.Text.Trim() + "' and '" + eCarID_End_Search.Text.Trim() + "' " :
                                 ((eCarID_Start_Search.Text.Trim() != "") && (eCarID_End_Search.Text.Trim() == "")) ? " AND Car_ID like '" + eCarID_Start_Search.Text.Trim() + "%' " :
                                 ((eCarID_Start_Search.Text.Trim() == "") && (eCarID_End_Search.Text.Trim() != "")) ? " AND Car_ID like '" + eCarID_End_Search.Text.Trim() + "%' " : "";
            string vWStr_ViolationDate = ((eViolationDate_Start_Search.Text.Trim() != "") && (eViolationDate_End_Search.Text.Trim() != "")) ? " AND ViolationDate between '" + DateTime.Parse(eViolationDate_Start_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' and '" + DateTime.Parse(eViolationDate_End_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " :
                                         ((eViolationDate_Start_Search.Text.Trim() != "") && (eViolationDate_End_Search.Text.Trim() == "")) ? " AND ViolationDate = '" + DateTime.Parse(eViolationDate_Start_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' " :
                                         ((eViolationDate_Start_Search.Text.Trim() == "") && (eViolationDate_End_Search.Text.Trim() != "")) ? " AND ViolationDate = '" + DateTime.Parse(eViolationDate_End_Search.Text.Trim()).Year.ToString("D4") + "/" + DateTime.Parse(eViolationDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " : "";
            vSQLStr = vSQLStr + Environment.NewLine +
                      vWStr_CaseType + Environment.NewLine +
                      vWStr_DepNo + Environment.NewLine +
                      vWStr_CarID + Environment.NewLine +
                      vWStr_ViolationDate + Environment.NewLine +
                      " ORDER BY BuildDate DESC, CaseNo";
            sdsViolationCase_List.SelectCommand = "";
            sdsViolationCase_List.SelectCommand = vSQLStr;
            gridViolationCaseList.DataSourceID = "sdsViolationCase_List";
            gridViolationCaseList.DataBind();
        }

        protected void ddlCaseType_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eCaseType_Search.Text = ddlCaseType_Search.SelectedValue.Trim();
        }

        protected void fvViolationCaseDetail_DataBound(object sender, EventArgs e)
        {
            string vDateURL_Temp = "";
            string vDateScript_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            switch (fvViolationCaseDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    plSearch.Visible = true;
                    /* 以下是用來限制各車站人員登入之後的權限用的...現在都先開放不限制
                    string vStationNo = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vDepNo + "' and InSHReport = 'V'", "DepNo");
                    Button bbNew_List = (Button)fvViolationCaseDetail.FindControl("bbCreate_List");
                    Button bbNew_Empty = (Button)fvViolationCaseDetail.FindControl("bbCreate_Empty");
                    Button bbEdit_List = (Button)fvViolationCaseDetail.FindControl("bbEdit_List");
                    Button bbDelete_List = (Button)fvViolationCaseDetail.FindControl("bbDelete_List");

                    if (bbNew_List != null)
                    {
                        bbNew_List.Visible = (vStationNo == "");
                    }
                    if (bbNew_Empty != null)
                    {
                        bbNew_Empty.Visible = (vStationNo == "");
                    }
                    if (bbEdit_List != null)
                    {
                        bbEdit_List.Visible = (vStationNo == "");
                    }
                    if (bbDelete_List != null)
                    {
                        bbDelete_List.Visible = (vStationNo == "");
                    }
                    //*/
                    break;
                case FormViewMode.Edit:
                    plSearch.Visible = false;
                    Label vModifyDate_Edit = (Label)fvViolationCaseDetail.FindControl("eModifyDate_Edit");
                    Label vModifyMan_Edit = (Label)fvViolationCaseDetail.FindControl("eModifyMan_Edit");
                    Label vModifyMan_C_Edit = (Label)fvViolationCaseDetail.FindControl("eModifyMan_C_Edit");
                    TextBox vContactDate_Edit = (TextBox)fvViolationCaseDetail.FindControl("eContactDate_Edit");
                    TextBox vAssignDate_Edit = (TextBox)fvViolationCaseDetail.FindControl("eAssignDate_Edit");

                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + vContactDate_Edit.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    vContactDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + vAssignDate_Edit.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    vAssignDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                    vModifyDate_Edit.Text = PF.GetChinsesDate(DateTime.Today);
                    vModifyMan_Edit.Text = vLoginID;
                    vModifyMan_C_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    break;
                case FormViewMode.Insert:
                    if (gridViolationCaseList.SelectedIndex > -1)
                    {
                        plSearch.Visible = false;
                        Label vCaseNo_INS = (Label)fvViolationCaseDetail.FindControl("eCaseNo_INS");
                        Label vBuildDate_INS = (Label)fvViolationCaseDetail.FindControl("eBuildDate_INS");
                        Label vBuildMan_INS = (Label)fvViolationCaseDetail.FindControl("eBuildMan_INS");
                        Label vBuildMan_C_INS = (Label)fvViolationCaseDetail.FindControl("eBuildMan_C_INS");
                        TextBox vContactDate_INS = (TextBox)fvViolationCaseDetail.FindControl("eContactDate_INS");
                        TextBox vAssignDate_INS = (TextBox)fvViolationCaseDetail.FindControl("eAssignDate_INS");

                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + vContactDate_INS.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        vContactDate_INS.Attributes["onClick"] = vDateScript_Temp;

                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + vAssignDate_INS.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        vAssignDate_INS.Attributes["onClick"] = vDateScript_Temp;

                        vBuildDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                        vBuildMan_INS.Text = vLoginID;
                        vBuildMan_C_INS.Text = Session["LoginName"].ToString().Trim();
                        vCaseNo_INS.Text = gridViolationCaseList.SelectedDataKey.Value.ToString();
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('請先選擇違規單！')");
                        Response.Write("</" + "Script>");
                        fvViolationCaseDetail.ChangeMode(FormViewMode.ReadOnly);
                    }
                    break;
            }
        }

        protected void eExcutort_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox vExcutort_Edit = (TextBox)fvViolationCaseDetail.FindControl("eExcutort_Edit");
            Label vExcutort_C_Edit = (Label)fvViolationCaseDetail.FindControl("eExcutort_C_Edit");
            string vEmpNo_Edit = vExcutort_Edit.Text.Trim();
            string vSQLStr_Edit = "select [Name] from Employee where EmpNo = '" + vEmpNo_Edit + "' ";
            string vEmpName_Edit = PF.GetValue(vConnStr, vSQLStr_Edit, "Name");
            if (vEmpName_Edit == "")
            {
                vEmpName_Edit = vEmpNo_Edit;
                vSQLStr_Edit = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Edit + "' order by EmpNo DESC";
                vEmpNo_Edit = PF.GetValue(vConnStr, vSQLStr_Edit, "EmpNo");
            }
            vExcutort_Edit.Text = vEmpNo_Edit;
            vExcutort_C_Edit.Text = vEmpName_Edit;
        }

        protected void eAssignedMan_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox vAssignedMan_Edit = (TextBox)fvViolationCaseDetail.FindControl("eAssignedMan_Edit");
            Label vAssignedMan_C_Edit = (Label)fvViolationCaseDetail.FindControl("eAssignedMan_C_Edit");
            string vEmpNo_Edit = vAssignedMan_Edit.Text.Trim();
            string vSQLStr_Edit = "select [Name] from Employee where EmpNo = '" + vEmpNo_Edit + "' ";
            string vEmpName_Edit = PF.GetValue(vConnStr, vSQLStr_Edit, "Name");
            if (vEmpName_Edit == "")
            {
                vEmpName_Edit = vEmpNo_Edit;
                vSQLStr_Edit = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Edit + "' order by EmpNo DESC";
                vEmpNo_Edit = PF.GetValue(vConnStr, vSQLStr_Edit, "EmpNo");
            }
            vAssignedMan_Edit.Text = vEmpNo_Edit;
            vAssignedMan_C_Edit.Text = vEmpName_Edit;
        }

        protected void eExcutort_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox vExcutort_INS = (TextBox)fvViolationCaseDetail.FindControl("eExcutort_INS");
            Label vExcutort_C_INS = (Label)fvViolationCaseDetail.FindControl("eExcutort_C_INS");
            string vEmpNo_INS = vExcutort_INS.Text.Trim();
            string vSQLStr_INS = "select [Name] from Employee where EmpNo = '" + vEmpNo_INS + "' ";
            string vEmpName_INS = PF.GetValue(vConnStr, vSQLStr_INS, "Name");
            if (vEmpName_INS == "")
            {
                vEmpName_INS = vEmpNo_INS;
                vSQLStr_INS = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_INS + "' order by EmpNo DESC";
                vEmpNo_INS = PF.GetValue(vConnStr, vSQLStr_INS, "EmpNo");
            }
            vExcutort_INS.Text = vEmpNo_INS;
            vExcutort_C_INS.Text = vEmpName_INS;
        }

        protected void eAssignedMan_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox vAssignedMan_INS = (TextBox)fvViolationCaseDetail.FindControl("eAssignedMan_INS");
            Label vAssignedMan_C_INS = (Label)fvViolationCaseDetail.FindControl("eAssignedMan_C_INS");
            string vEmpNo_INS = vAssignedMan_INS.Text.Trim();
            string vSQLStr_INS = "select [Name] from Employee where EmpNo = '" + vEmpNo_INS + "' ";
            string vEmpName_INS = PF.GetValue(vConnStr, vSQLStr_INS, "Name");
            if (vEmpName_INS == "")
            {
                vEmpName_INS = vEmpNo_INS;
                vSQLStr_INS = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_INS + "' order by EmpNo DESC";
                vEmpNo_INS = PF.GetValue(vConnStr, vSQLStr_INS, "EmpNo");
            }
            vAssignedMan_INS.Text = vEmpNo_INS;
            vAssignedMan_C_INS.Text = vEmpName_INS;
        }

        protected void sdsViolationCaseDetail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridViolationCaseDetailList.DataBind();
                plSearch.Visible = true;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('刪除完成！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + e.Exception.Message + "')");
                Response.Write("</" + "Script>");
            }
        }

        protected void sdsViolationCaseDetail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridViolationCaseDetailList.DataBind();
                plSearch.Visible = true;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('新增完成！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + e.Exception.Message + "')");
                Response.Write("</" + "Script>");
            }
        }

        protected void sdsViolationCaseDetail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            Label vCaseNo_INS = (Label)fvViolationCaseDetail.FindControl("eCaseNo_INS");
            if (vCaseNo_INS.Text.Trim() != "")
            {
                TextBox vContactDate_INS = (TextBox)fvViolationCaseDetail.FindControl("eContactDate_INS");
                TextBox vAssignDate_INS = (TextBox)fvViolationCaseDetail.FindControl("eAssignDate_INS");
                string vSQLStr_INS = "";
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCaseNo = vCaseNo_INS.Text.Trim();
                int vItemsMax = 0;
                vSQLStr_INS = "select Isnull(MAX(Items), 0) ItemsMax from ViolationCaseB where CaseNo = '" + vCaseNo + "' ";
                vItemsMax = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_INS, "ItemsMax"));
                string vItemsStr = (vItemsMax + 1).ToString("D4");
                string vCaseNoItem = vCaseNo + vItemsStr;
                e.Command.Parameters["@CaseNo"].Value = vCaseNo;
                e.Command.Parameters["@Items"].Value = vItemsStr;
                e.Command.Parameters["@CaseNoItem"].Value = vCaseNoItem;
                e.Command.Parameters["@BuildDate"].Value = DateTime.Today;
                e.Command.Parameters["@BuildMan"].Value = vLoginID;
                if ((vContactDate_INS != null) && (vContactDate_INS.Text.Trim() != ""))
                {
                    e.Command.Parameters["@ContactDate"].Value = DateTime.Parse(vContactDate_INS.Text.Trim());
                }
                if ((vAssignDate_INS != null) && (vAssignDate_INS.Text.Trim() != ""))
                {
                    e.Command.Parameters["@AssignDate"].Value = DateTime.Parse(vAssignDate_INS.Text.Trim());
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('違規單號空白，請先選擇違規單！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void sdsViolationCaseDetail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridViolationCaseDetailList.DataBind();
                plSearch.Visible = true;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('修改完成！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + e.Exception.Message + "')");
                Response.Write("</" + "Script>");
            }
        }

        protected void sdsViolationCaseDetail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            Label vCaseNoItem_Edit = (Label)fvViolationCaseDetail.FindControl("eCaseNoItem_Edit");
            string vCaseNoItem = vCaseNoItem_Edit.Text.Trim();
            if (vCaseNoItem != "")
            {
                TextBox vContactDate_Edit = (TextBox)fvViolationCaseDetail.FindControl("eContactDate_Edit");
                TextBox vAssignDate_Edit = (TextBox)fvViolationCaseDetail.FindControl("eAssignDate_Edit");
                if ((vContactDate_Edit != null) && (vContactDate_Edit.Text.Trim() != ""))
                {
                    e.Command.Parameters["@ContactDate"].Value = DateTime.Parse(vContactDate_Edit.Text.Trim());
                }
                if ((vAssignDate_Edit != null) && (vAssignDate_Edit.Text.Trim() != ""))
                {
                    e.Command.Parameters["@AssignDate"].Value = DateTime.Parse(vAssignDate_Edit.Text.Trim());
                }
                e.Command.Parameters["@ModifyDate"].Value = DateTime.Today;
                e.Command.Parameters["@ModifyMan"].Value = vLoginID;
            }
            else
            {

            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            DataSourceBinded();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}