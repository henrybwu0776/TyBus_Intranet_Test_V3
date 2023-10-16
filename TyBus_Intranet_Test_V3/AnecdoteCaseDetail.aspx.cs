using Amaterasu_Function;
using System;
using System.Globalization;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class AnecdoteCaseDetail : Page
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

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                System.Threading.Thread.CurrentThread.CurrentCulture = Cal;
                //
                DateTime vToday = DateTime.Today;
                vLoginID = (Session["LoginID"] != null) ? Session["LoginID"].ToString().Trim() : "";
                vLoginName = (Session["LoginName"] != null) ? Session["LoginName"].ToString().Trim() : "";
                vLoginDepNo = (Session["LoginDepNo"] != null) ? Session["LoginDepNo"].ToString().Trim() : "";
                vLoginDepName = (Session["LoginDepName"] != null) ? Session["LoginDepName"].ToString().Trim() : "";
                vLoginTitle = (Session["LoginTitle"] != null) ? Session["LoginTitle"].ToString().Trim() : "";
                vLoginTitleName = (Session["LoginTitleName"] != null) ? Session["LoginTitleName"].ToString().Trim() : "";
                vLoginEmpType = (Session["LoginEmpType"] != null) ? Session["LoginEmpType"].ToString().Trim() : "";
                vComputerName = Environment.MachineName; //2021.09.27 新增取得電腦名稱

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Start_Check.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Start_Check.Attributes["onClick"] = vCaseDateScript;
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_End_Check.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_End_Check.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        vStationNo = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vLoginDepNo + "' and InSHReport = 'V'", "DepNo");

                        eDepNo_Start_Check.Text = vStationNo;
                        eDepNo_End_Check.Text = "";
                        eDepNo_Start_Check.Enabled = (vStationNo == "");
                        eDepNo_End_Check.Enabled = (vStationNo == "");
                    }
                    CaseDataBind();
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

        private void CaseDataBind()
        {
            if (vLoginID != "")
            {
                string vHasInsuStr = (chkHasInsurance.Checked) ? " and HasInsurance = 1" : " and HasInsurance = 0";
                string vDepNoWhereStr = ((eDepNo_Start_Check.Text.Trim() != "") && (eDepNo_End_Check.Text.Trim() != "")) ? " and DepNo between '" + eDepNo_Start_Check.Text.Trim() + "' and '" + eDepNo_End_Check.Text.Trim() + "' " :
                                        ((eDepNo_Start_Check.Text.Trim() != "") && (eDepNo_End_Check.Text.Trim() == "")) ? " and DepNo = '" + eDepNo_Start_Check.Text.Trim() + "' " :
                                        ((eDepNo_Start_Check.Text.Trim() == "") && (eDepNo_End_Check.Text.Trim() != "")) ? " and DepNo = '" + eDepNo_End_Check.Text.Trim() + "' " :
                                        "";
                string vCaseDateWhereStr = ((eCaseDate_Start_Check.Text.Trim() != "") && (eCaseDate_End_Check.Text.Trim() != "")) ?
                                           " and BuildDate between '" + PF.TransDateString(eCaseDate_Start_Check.Text.Trim(), "B") + "' and '" + PF.TransDateString(eCaseDate_End_Check.Text.Trim(), "B") + "' " :
                                           ((eCaseDate_Start_Check.Text.Trim() != "") && (eCaseDate_End_Check.Text.Trim() == "")) ?
                                           " and BuildDate = '" + PF.TransDateString(eCaseDate_Start_Check.Text.Trim(), "B") + "' " :
                                           ((eCaseDate_Start_Check.Text.Trim() == "") && (eCaseDate_End_Check.Text.Trim() != "")) ?
                                           " and BuildDate = '" + PF.TransDateString(eCaseDate_End_Check.Text.Trim(), "B") + "' " :
                                           "";
                string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? " and Car_ID like '%" + eCarID_Search.Text.Trim() + "%' " + Environment.NewLine : "";
                string vDriverStr = (eDriver_Search.Text.Trim() != "") ? " and Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
                sdsAnecdoteCaseA_List.SelectCommand = "SELECT CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, " + Environment.NewLine +
                                                      "       (SELECT NAME FROM Employee WHERE (EMPNO = AnecdoteCase.BuildMan)) AS BuildManName, " + Environment.NewLine +
                                                      "       Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, " + Environment.NewLine +
                                                      "       IsNoDeduction, DeductionDate, Remark, CaseOccurrence " + Environment.NewLine +
                                                      "  FROM AnecdoteCase " + Environment.NewLine +
                                                      " where isnull(CaseNo, '') <> '' " + Environment.NewLine +
                                                      vHasInsuStr +
                                                      vDepNoWhereStr +
                                                      vCaseDateWhereStr +
                                                      vCarIDStr +
                                                      vDriverStr +
                                                      " order by CaseNo DESC";
                gridAnecdoteCaseA_List.DataBind();
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            CaseDataBind();
            gridAnecdoteCaseA_List.PageIndex = 0;
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvAnecdoteCaseC_DataBound(object sender, EventArgs e)
        {
            string vDateURL_Temp = "";
            string vDateScript_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvAnecdoteCaseC.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    plSearch.Visible = true;
                    /* 以下是用來限制各車站人員登入之後的權限用的...現在都先開放不限制
                    string vStationNo = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vDepNo + "' and InSHReport = 'V'", "DepNo");
                    Button bbNew_List = (Button)fvAnecdoteCaseC.FindControl("bbNew_List");
                    Button bbNew_Empty = (Button)fvAnecdoteCaseC.FindControl("bbNew_Empty");
                    Button bbEdit_List = (Button)fvAnecdoteCaseC.FindControl("bbEdit_List");
                    Button bbDelete_List = (Button)fvAnecdoteCaseC.FindControl("bbDelete_List");

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
                    Label vModifyDate_Edit = (Label)fvAnecdoteCaseC.FindControl("eModifyDate_Edit");
                    Label vModifyMan_Edit = (Label)fvAnecdoteCaseC.FindControl("eModifyMan_Edit");
                    Label vModifyMan_C_Edit = (Label)fvAnecdoteCaseC.FindControl("eModifyMan_C_Edit");
                    TextBox vContactDate_Edit = (TextBox)fvAnecdoteCaseC.FindControl("eContactDate_Edit");
                    TextBox vAssignDate_Edit = (TextBox)fvAnecdoteCaseC.FindControl("eAssignDate_Edit");

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
                    if (gridAnecdoteCaseA_List.SelectedIndex > -1)
                    {
                        plSearch.Visible = false;
                        Label vCaseNo_INS = (Label)fvAnecdoteCaseC.FindControl("eCaseNo_INS");
                        Label vBuildDate_INS = (Label)fvAnecdoteCaseC.FindControl("eBuildDate_INS");
                        Label vBuildMan_INS = (Label)fvAnecdoteCaseC.FindControl("eBuildMan_INS");
                        Label vBuildMan_C_INS = (Label)fvAnecdoteCaseC.FindControl("eBuildMan_C_INS");
                        TextBox vContactDate_INS = (TextBox)fvAnecdoteCaseC.FindControl("eContactDate_INS");
                        TextBox vAssignDate_INS = (TextBox)fvAnecdoteCaseC.FindControl("eAssignDate_INS");

                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + vContactDate_INS.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        vContactDate_INS.Attributes["onClick"] = vDateScript_Temp;

                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + vAssignDate_INS.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        vAssignDate_INS.Attributes["onClick"] = vDateScript_Temp;

                        vBuildDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                        vBuildMan_INS.Text = vLoginID;
                        vBuildMan_C_INS.Text = Session["LoginName"].ToString().Trim();
                        vCaseNo_INS.Text = gridAnecdoteCaseA_List.SelectedDataKey.Value.ToString();
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('請先選擇違規單！')");
                        Response.Write("</" + "Script>");
                        fvAnecdoteCaseC.ChangeMode(FormViewMode.ReadOnly);
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
            TextBox vExcutort_Edit = (TextBox)fvAnecdoteCaseC.FindControl("eExcutort_Edit");
            Label vExcutort_C_Edit = (Label)fvAnecdoteCaseC.FindControl("eExcutort_C_Edit");
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
            TextBox vAssignedMan_Edit = (TextBox)fvAnecdoteCaseC.FindControl("eAssignedMan_Edit");
            Label vAssignedMan_C_Edit = (Label)fvAnecdoteCaseC.FindControl("eAssignedMan_C_Edit");
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
            TextBox vExcutort_INS = (TextBox)fvAnecdoteCaseC.FindControl("eExcutort_INS");
            Label vExcutort_C_INS = (Label)fvAnecdoteCaseC.FindControl("eExcutort_C_INS");
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
            TextBox vAssignedMan_INS = (TextBox)fvAnecdoteCaseC.FindControl("eAssignedMan_INS");
            Label vAssignedMan_C_INS = (Label)fvAnecdoteCaseC.FindControl("eAssignedMan_C_INS");
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

        protected void sdsAnecdoteCaseC_Data_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridAnecdoteCaseDetailList.DataBind();
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

        protected void sdsAnecdoteCaseC_Data_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridAnecdoteCaseDetailList.DataBind();
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

        protected void sdsAnecdoteCaseC_Data_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            Label vCaseNo_INS = (Label)fvAnecdoteCaseC.FindControl("eCaseNo_INS");
            if (vCaseNo_INS.Text.Trim() != "")
            {
                TextBox vContactDate_INS = (TextBox)fvAnecdoteCaseC.FindControl("eContactDate_INS");
                TextBox vAssignDate_INS = (TextBox)fvAnecdoteCaseC.FindControl("eAssignDate_INS");
                string vSQLStr_INS = "";
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCaseNo = vCaseNo_INS.Text.Trim();
                int vItemsMax = 0;
                vSQLStr_INS = "select Isnull(MAX(Items), 0) ItemsMax from AnecdoteCaseC where CaseNo = '" + vCaseNo + "' ";
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
                Response.Write("alert('肇事單號空白，請先選擇肇事單！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void sdsAnecdoteCaseC_Data_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridAnecdoteCaseDetailList.DataBind();
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

        protected void sdsAnecdoteCaseC_Data_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            Label vCaseNoItem_Edit = (Label)fvAnecdoteCaseC.FindControl("eCaseNoItem_Edit");
            string vCaseNoItem = vCaseNoItem_Edit.Text.Trim();
            if (vCaseNoItem != "")
            {
                TextBox vContactDate_Edit = (TextBox)fvAnecdoteCaseC.FindControl("eContactDate_Edit");
                TextBox vAssignDate_Edit = (TextBox)fvAnecdoteCaseC.FindControl("eAssignDate_Edit");
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

        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            string vDriver = eDriver_Search.Text.Trim();
            string vDriverName = "";
            string vSQLStr_Temp = "select [Name] from  Employee where EmpNo = '" + vDriver + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vDriverName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName == "")
            {
                vDriverName = vDriver;
                vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vDriverName + "' order by EmpNo DESC";
                vDriver = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Search.Text = vDriver.Trim();
            eDriverName_Search.Text = vDriverName.Trim();
        }
    }
}