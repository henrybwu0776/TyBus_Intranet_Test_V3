using Amaterasu_Function;
using System;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class CustomServiceHistory : System.Web.UI.Page
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

                string vDateURL_Search = "InputDate_ChineseYears.aspx?TextboxID=" + eStartDate_Modify_Search.ClientID;
                string vScript_Search = "window.open('" + vDateURL_Search + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eStartDate_Modify_Search.Attributes["onClick"] = vScript_Search;

                vDateURL_Search = "InputDate_ChineseYears.aspx?TextboxID=" + eStartDate_Build_Search.ClientID;
                vScript_Search = "window.open('" + vDateURL_Search + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eStartDate_Build_Search.Attributes["onClick"] = vScript_Search;

                vDateURL_Search = "InputDate_ChineseYears.aspx?TextboxID=" + eEndDate_Build_Search.ClientID;
                vScript_Search = "window.open('" + vDateURL_Search + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eEndDate_Build_Search.Attributes["onClick"] = vScript_Search;

                vDateURL_Search = "InputDate_ChineseYears.aspx?TextboxID=" + eEndDate_Modify_Search.ClientID;
                vScript_Search = "window.open('" + vDateURL_Search + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eEndDate_Modify_Search.Attributes["onClick"] = vScript_Search;

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
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

        private string GetSelectStr()
        {
            string vResultStr = "SELECT HistoryIndex, ServiceNo, BuildDate, BuildTime, " + Environment.NewLine +
                                "       BuildMan, (SELECT NAME FROM EMployee WHERE (EMPNO = a.BuildMan)) AS BuildManName, " + Environment.NewLine +
                                "       ServiceType, (SELECT TypeText FROM CustomServiceType WHERE (TypeStep = 1) AND (TypeLevel1 = a.ServiceType)) AS Type1_Name, " + Environment.NewLine +
                                "       ServiceTypeB, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_2 WHERE (TypeStep = 2) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB)) AS Type2_Name, " + Environment.NewLine +
                                "       ServiceTypeC, (SELECT TypeText FROM CustomServiceType AS CustomServiceType_1 WHERE (TypeStep = 3) AND (TypeLevel1 = a.ServiceType) AND (TypeLevel2 = a.ServiceTypeB) AND (TypeLevel3 = a.ServiceTypeC)) AS Type3_Name, " + Environment.NewLine +
                                "       LinesNo, Car_ID, " + Environment.NewLine +
                                "       Driver, (SELECT NAME FROM Employee AS Employee_1 WHERE (EMPNO = a.Driver)) AS Driver_Name, " + Environment.NewLine +
                                "       LinesNo2, Car_ID2, Driver2, (SELECT NAME FROM Employee AS Employee_1 WHERE (EMPNO = a.Driver2)) AS Driver_Name2, " + Environment.NewLine +
                                "       CaseSource, (select ClassTxt from DBDICB where FKey = (cast('客服專線記錄表' as char(16))+cast('fmCustomService' as char(16))+'CaseSource') and ClassNo = a.CaseSource) CaseSource_C, " + Environment.NewLine +
                                "       ServiceDate, IsTrue " + Environment.NewLine +
                                "  FROM CustomServiceHistory AS a WHERE (1 = 1)";
            string vWStr_ModifyMan = (eModifyMan_Search.Text.Trim() != "") ? " and a.ModifyMan = '" + eModifyMan_Search.Text.Trim() + "' " : "";
            string vWStr_BuildMan = (eBuildMan_Search.Text.Trim() != "") ? " and a.BuildMan = '" + eBuildMan_Search.Text.Trim() + "' " : "";
            string vWStr_ModifyDate = ((eStartDate_Modify_Search.Text.Trim() != "") && (eEndDate_Modify_Search.Text.Trim() != "")) ? " and a.ModifyDate between '" + eStartDate_Modify_Search.Text.Trim() + "' and '" + eEndDate_Modify_Search.Text.Trim() + "' " :
                                      ((eStartDate_Modify_Search.Text.Trim() != "") && (eEndDate_Modify_Search.Text.Trim() == "")) ? " and a.ModifyDate = '" + eStartDate_Modify_Search.Text.Trim() + "' " :
                                      ((eStartDate_Modify_Search.Text.Trim() == "") && (eEndDate_Modify_Search.Text.Trim() != "")) ? " and a.ModifyDate = '" + eEndDate_Modify_Search.Text.Trim() + "' " : "";
            string vWStr_BuildDate = ((eStartDate_Build_Search.Text.Trim() != "") && (eEndDate_Build_Search.Text.Trim() != "")) ? " and a.BuildDate between '" + eStartDate_Build_Search.Text.Trim() + "' and '" + eEndDate_Build_Search.Text.Trim() + "' " :
                                      ((eStartDate_Build_Search.Text.Trim() != "") && (eEndDate_Build_Search.Text.Trim() == "")) ? " and a.BuildDate = '" + eStartDate_Build_Search.Text.Trim() + "' " :
                                      ((eStartDate_Build_Search.Text.Trim() == "") && (eEndDate_Build_Search.Text.Trim() != "")) ? " and a.BuildDate = '" + eEndDate_Build_Search.Text.Trim() + "' " : "";
            vResultStr = vResultStr + vWStr_ModifyMan + vWStr_BuildMan + vWStr_ModifyDate + vWStr_BuildDate;
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelectCommandStr = GetSelectStr();
            sdsServiceHistory.SelectCommand = vSelectCommandStr;
            gridServiceHistoryList.DataSourceID = sdsServiceHistory.ID;
            sdsServiceHistory.DataBind();
        }

        protected void ddlModifyMan_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eModifyMan_Search.Text = ddlModifyMan_Search.SelectedValue;
        }

        protected void ddlBuildMan_Searech_SelectedIndexChanged(object sender, EventArgs e)
        {
            eBuildMan_Search.Text = ddlModifyMan_Search.SelectedValue;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvServiceDataView_DataBound(object sender, EventArgs e)
        {
            if (fvServiceDataView.CurrentMode == FormViewMode.ReadOnly)
            {
                TextBox eIsTrue = (TextBox)fvServiceDataView.FindControl("eIsTrue");
                //RadioButtonList rbIsTrue = (RadioButtonList)fvServiceDataView.FindControl("rbIsTrue");
                DropDownList ddlIsTrue = (DropDownList)fvServiceDataView.FindControl("ddlIsTrue");
                string vIsTrue = ((eIsTrue != null) && (eIsTrue.Text.Trim() != "")) ? eIsTrue.Text.Trim() : "";
                //if (rbIsTrue != null)
                if ((ddlIsTrue != null) && (eIsTrue != null))
                {
                    ddlIsTrue.SelectedIndex = ddlIsTrue.Items.IndexOf(ddlIsTrue.Items.FindByValue(vIsTrue));
                }
            }
        }
    }
}