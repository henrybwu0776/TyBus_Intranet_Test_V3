using Amaterasu_Function;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class CustomServiceDetail : System.Web.UI.Page
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

                UnobtrusiveValidationMode = System.Web.UI.UnobtrusiveValidationMode.None;

                string vTempDate_URL = "InputDate_ChineseYears.aspx?TextboxID=" + eServiceDate_Start_Search.ClientID;
                string vTempDate_Script = "window.open('" + vTempDate_URL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eServiceDate_Start_Search.Attributes["onClick"] = vTempDate_Script;

                vTempDate_URL = "InputDate_ChineseYears.aspx?TextboxID=" + eServiceDate_End_Search.ClientID;
                vTempDate_Script = "window.open('" + vTempDate_URL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eServiceDate_End_Search.Attributes["onClick"] = vTempDate_Script;

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
                        Session["CustomServicePrint"] = "-1"; //先給定初值，不然會跳錯誤訊息
                        eBuildDate_Start_Search.Text = PF.GetMonthFirstDay(DateTime.Today, "C");
                        eBuildDate_End_Search.Text = PF.GetMonthLastDay(DateTime.Today, "C");

                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            string vSQLStr_Temp = "select Cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                                  " union all " + Environment.NewLine +
                                                  "select CLassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                                  " where FKey = '客服專線記錄表  fmCustomService IsTrue' " + Environment.NewLine +
                                                  "   and ClassNo <> 'IT02' " + Environment.NewLine +
                                                  " order by CLassNo";
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            ddlIsTrue_Search.Items.Clear();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ddlIsTrue_Search.Items.Add(new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim()));
                            }
                            ddlIsTrue_Search.SelectedIndex = 0;
                            eIsTrue_Search.Text = "";
                        }
                    }
                    ServiceDataBind();
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
            string vWStr_BuildMan = (eBuildMan_Searech.Text.Trim() != "") ? " and BuildMan = '" + eBuildMan_Searech.Text.Trim() + "' " + Environment.NewLine : "";

            string vWStr_BuildDate = ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() != "")) ? " and BuildDate between '" + (DateTime.Parse(eBuildDate_Start_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' and '" + (DateTime.Parse(eBuildDate_End_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() == "")) ? " and BuildDate = '" + (DateTime.Parse(eBuildDate_Start_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eBuildDate_Start_Search.Text.Trim() == "") && (eBuildDate_End_Search.Text.Trim() != "")) ? " and BuildDate = '" + (DateTime.Parse(eBuildDate_End_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";

            string vWStr_ServiceDate = ((eServiceDate_Start_Search.Text.Trim() != "") && (eServiceDate_End_Search.Text.Trim() != "")) ? " and ServiceDate between '" + (DateTime.Parse(eServiceDate_Start_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' and '" + (DateTime.Parse(eServiceDate_End_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eServiceDate_Start_Search.Text.Trim() != "") && (eServiceDate_End_Search.Text.Trim() == "")) ? " and ServiceDate = '" + (DateTime.Parse(eServiceDate_Start_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine :
                                     ((eServiceDate_Start_Search.Text.Trim() == "") && (eServiceDate_End_Search.Text.Trim() != "")) ? " and ServiceDate = '" + (DateTime.Parse(eServiceDate_End_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " + Environment.NewLine : "";

            string vWStr_ServiceType = (eServiceType_Search.Text.Trim() != "") ? " and ServiceType = '" + eServiceType_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vWStr_ServiceTypeB = (eServiceTypeB_Search.Text.Trim() != "") ? " and ServiceTypeB = '" + eServiceTypeB_Search.Text.Trim().Substring(3) + "' " + Environment.NewLine : "";

            string vWStr_ServiceTypeC = (eServiceTypeC_Search.Text.Trim() != "") ? " and ServiceTypeC = '" + eServiceTypeC_Search.Text.Trim().Substring(5) + "' " + Environment.NewLine : "";

            string vWStr_LinesNo = (eLinesNo_Search.Text.Trim() != "") ? " and (LinesNo like '" + eLinesNo_Search.Text.Trim() + "%' or LinesNo2 like '" + eLinesNo_Search.Text.Trim() + "') " + Environment.NewLine : "";

            string vWStr_CarID = (eCarID_Search.Text.Trim() != "") ? " and (Car_ID like '" + eCarID_Search.Text.Trim() + "' or Car_ID2 like '" + eCarID_Search.Text.Trim() + "') " + Environment.NewLine : "";

            string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? " and (Driver like '" + eDriver_Search.Text.Trim() + "%' or DriverName like '" + eDriver_Search.Text.Trim() + "%') or (Driver2 like '" + eDriver_Search.Text.Trim() + "%' or DriverName2 like '" + eDriver_Search.Text.Trim() + "%')" + Environment.NewLine : "";

            string vWStr_AthorityDep = (eAthorityDep_Search.Text.Trim() != "") ? " and (AthorityDepNo = '" + eAthorityDep_Search.Text.Trim() + "' or AthorityDepNo = '" + eAthorityDep_Search.Text.Trim() + "') " + Environment.NewLine : "";

            string vWStr_ClosedOnly = (rbIsClosed.SelectedValue == "1") ? " and isnull(IsClosed, 0) = 1 " + Environment.NewLine :
                                      (rbIsClosed.SelectedValue == "2") ? " and isnull(IsClosed, 0) = 0 " + Environment.NewLine :
                                      "";

            string vWStr_CaseSource = (eCaseSource_Search.Text.Trim() != "") ? " and CaseSource = '" + eCaseSource_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vWStr_CivicName = (eCivicName_Search.Text.Trim() != "") ? " and CivicName like '%" + eCivicName_Search.Text.Trim() + "%'" + Environment.NewLine : "";

            string vWStr_CivicTel = (eCivicTel_Search.Text.Trim() != "") ? " and (CivicTelNo like '%" + eCivicTel_Search.Text.Trim() + "%' or CivicCellPhone like '%" + eCivicTel_Search.Text.Trim() + "%' )" + Environment.NewLine : "";

            string vWStr_IsTrue = (eIsTrue_Search.Text.Trim() != "") ? " and isTrue = '" + eIsTrue_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vResultStr = "select ServiceNo, BuildDate, BuildTime, BuildMan, (select [Name] from Employee where EmpNo = a.BuildMan) BuildMan_C, " + Environment.NewLine +
                                "       ServiceType, (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) ServiceType_C, " + Environment.NewLine +
                                "       ServiceTypeB, (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) ServiceTypeB_C, " + Environment.NewLine +
                                "       ServiceTypeC, (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) ServiceTypeC_C, " + Environment.NewLine +
                                "       LinesNo, Car_ID, Driver, DriverName, BoardTime, BoardStation, " + Environment.NewLine +
                                "       GetoffTime, GetoffStation, ServiceNote, IsReplied, IsPending, AssignDate, AssignMan, " + Environment.NewLine +
                                "       (select [Name] from Employee where EmpNo = a.AssignMan) AssignMan_C, IsClosed, CloseDate, CloseMan, " + Environment.NewLine +
                                "       (select [Name] from Employee where EmpNo = a.CloseMan) CloseMan_C, " + Environment.NewLine +
                                "       CaseSource, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') AND (CLASSNO = a.CaseSource)) CaseSource_C, " + Environment.NewLine +
                                "       ServiceDate, (select ClassTxt from DBDICB where FKey = '客服專線記錄表  fmCustomService IsTrue' and CLassNo = a.IsTrue) IsTrue_C " + Environment.NewLine +
                                "  from CustomService a " + Environment.NewLine +
                                " where isnull(ServiceNo, '') <> '' " + Environment.NewLine +
                                vWStr_BuildMan + vWStr_ServiceDate + vWStr_BuildDate + vWStr_ServiceType + vWStr_ServiceTypeB + vWStr_ServiceTypeC + vWStr_LinesNo +
                                vWStr_CarID + vWStr_Driver + vWStr_AthorityDep + vWStr_ClosedOnly + vWStr_CaseSource + vWStr_CivicName + vWStr_CivicTel + vWStr_IsTrue +
                                " order by a.BuildDate DESC";
            return vResultStr;
        }

        private void ServiceDataBind()
        {
            if (vLoginID != "")
            {
                string vResearchSQLStr = GetSelectStr();
                dsServiceDataList.SelectCommand = "";
                dsServiceDataList.SelectCommand = vResearchSQLStr;
                gridDataList.DataSourceID = "dsServiceDataList";
                dsServiceDataList.DataBind();
            }
        }

        protected void ddlBuildMan_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eBuildMan_Searech.Text = ddlBuildMan_Search.SelectedValue;
        }

        protected void ddlServiceType_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eServiceType_Search.Text = ddlServiceType_Search.SelectedValue;
        }

        protected void ddlServiceTypeB_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eServiceTypeB_Search.Text = ddlServiceTypeB_Search.SelectedValue;
        }

        protected void ddlServiceTypeC_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eServiceTypeC_Search.Text = ddlServiceTypeC_Search.SelectedValue;
        }

        protected void ddlCaseSource_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eCaseSource_Search.Text = ddlCaseSource_Search.SelectedValue;
        }

        protected void ddlAthorityDep_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eAthorityDep_Search.Text = ddlAthorityDep_Search.SelectedValue;
        }

        protected void bbRequset_Click(object sender, EventArgs e)
        {
            ServiceDataBind();
            gridDataList.PageIndex = 0;
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvCustomServiceDetail_ItemCreated(object sender, EventArgs e)
        {
            string vDateURL_Temp = "";
            string vDateScript_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            switch (fvCustomServiceDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    plSearch.Visible = true;
                    break;

                case FormViewMode.Edit:
                    plSearch.Visible = false;
                    Label vModifyDate_Edit = (Label)fvCustomServiceDetail.FindControl("eModifyDate_Edit");
                    Label vModifyMan_Edit = (Label)fvCustomServiceDetail.FindControl("eModifyMan_Edit");
                    Label vModifyMan_C_Edit = (Label)fvCustomServiceDetail.FindControl("eModifyMan_C_Edit");
                    TextBox vContactDate_Edit = (TextBox)fvCustomServiceDetail.FindControl("eContactDate_Edit");
                    TextBox vAssignDate_Edit = (TextBox)fvCustomServiceDetail.FindControl("eAssignDate_Edit");

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
                    if (gridDataList.SelectedIndex > -1)
                    {
                        plSearch.Visible = false;
                        Label vServiceNo_INS = (Label)fvCustomServiceDetail.FindControl("eServiceNo_INS");
                        Label vBuildDate_INS = (Label)fvCustomServiceDetail.FindControl("eBuildDate_INS");
                        Label vBuildMan_INS = (Label)fvCustomServiceDetail.FindControl("eBuildMan_INS");
                        Label vBuildMan_C_INS = (Label)fvCustomServiceDetail.FindControl("eBuildMan_C_INS");
                        TextBox vContactDate_INS = (TextBox)fvCustomServiceDetail.FindControl("eContactDate_INS");
                        TextBox vAssignDate_INS = (TextBox)fvCustomServiceDetail.FindControl("eAssignDate_INS");

                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + vContactDate_INS.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        vContactDate_INS.Attributes["onClick"] = vDateScript_Temp;

                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + vAssignDate_INS.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        vAssignDate_INS.Attributes["onClick"] = vDateScript_Temp;

                        vBuildDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                        vBuildMan_INS.Text = vLoginID;
                        vBuildMan_C_INS.Text = Session["LoginName"].ToString().Trim();
                        vServiceNo_INS.Text = gridDataList.SelectedDataKey.Value.ToString();
                    }
                    else
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('請先選擇客訴單！')");
                        Response.Write("</" + "Script>");
                        fvCustomServiceDetail.ChangeMode(FormViewMode.ReadOnly);
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
            TextBox eExcutort_Edit = (TextBox)fvCustomServiceDetail.FindControl("eExcutort_Edit");
            Label eExcutort_C_Edit = (Label)fvCustomServiceDetail.FindControl("eExcutort_C_Edit");
            string vEmpNo_Edit = eExcutort_Edit.Text.Trim();
            string vSQLStr_Edit = "select [Name] from Employee where EmpNo = '" + vEmpNo_Edit + "' ";
            string vEmpName_Edit = PF.GetValue(vConnStr, vSQLStr_Edit, "Name");
            if (vEmpName_Edit == "")
            {
                vEmpName_Edit = vEmpNo_Edit;
                vSQLStr_Edit = "select Top 1 EmpNo from EmpLoyee where [Name] = '" + vEmpName_Edit + "' order by EmpNo DESC ";
                vEmpNo_Edit = PF.GetValue(vConnStr, vSQLStr_Edit, "EmpNo");
            }
            eExcutort_Edit.Text = vEmpNo_Edit;
            eExcutort_C_Edit.Text = vEmpName_Edit;
        }

        protected void eAssignedMan_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eAssignedMan_Edit = (TextBox)fvCustomServiceDetail.FindControl("eAssignedMan_Edit");
            Label eAssignedMan_C_Edit = (Label)fvCustomServiceDetail.FindControl("eAssignedMan_C_Edit");
            string vEmpNo_Edit = eAssignedMan_Edit.Text.Trim();
            string vSQLStr_Edit = "select [Name] from Employee where EmpNo = '" + vEmpNo_Edit + "' ";
            string vEmpName_Edit = PF.GetValue(vConnStr, vSQLStr_Edit, "Name");
            if (vEmpName_Edit == "")
            {
                vEmpName_Edit = vEmpNo_Edit;
                vSQLStr_Edit = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_Edit + "' order by EmpNo DESC";
                vEmpNo_Edit = PF.GetValue(vConnStr, vSQLStr_Edit, "EmpNo");
            }
            eAssignedMan_Edit.Text = vEmpNo_Edit;
            eAssignedMan_C_Edit.Text = vEmpName_Edit;
        }

        protected void eExcutort_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eExcutort_INS = (TextBox)fvCustomServiceDetail.FindControl("eExcutort_INS");
            Label eExcutort_C_INS = (Label)fvCustomServiceDetail.FindControl("eExcutort_C_INS");
            string vEmpNo_INS = eExcutort_INS.Text.Trim();
            string vSQLStr_INS = "select [Name] from Employee where EmpNo = '" + vEmpNo_INS + "' ";
            string vEmpName_INS = PF.GetValue(vConnStr, vSQLStr_INS, "Name");
            if (vEmpName_INS == "")
            {
                vEmpName_INS = vEmpNo_INS;
                vSQLStr_INS = "select Top 1 EmpNo from EmpLoyee where [Name] = '" + vEmpName_INS + "' order by EmpNo DESC ";
                vEmpNo_INS = PF.GetValue(vConnStr, vSQLStr_INS, "EmpNo");
            }
            eExcutort_INS.Text = vEmpNo_INS;
            eExcutort_C_INS.Text = vEmpName_INS;
        }

        protected void eAssignedMan_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eAssignedMan_INS = (TextBox)fvCustomServiceDetail.FindControl("eAssignedMan_INS");
            Label eAssignedMan_C_INS = (Label)fvCustomServiceDetail.FindControl("eAssignedMan_C_INS");
            string vEmpNo_INS = eAssignedMan_INS.Text.Trim();
            string vSQLStr_INS = "select [Name] from Employee where EmpNo = '" + vEmpNo_INS + "' ";
            string vEmpName_INS = PF.GetValue(vConnStr, vSQLStr_INS, "Name");
            if (vEmpName_INS == "")
            {
                vEmpName_INS = vEmpNo_INS;
                vSQLStr_INS = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_INS + "' order by EmpNo DESC";
                vEmpNo_INS = PF.GetValue(vConnStr, vSQLStr_INS, "EmpNo");
            }
            eAssignedMan_INS.Text = vEmpNo_INS;
            eAssignedMan_C_INS.Text = vEmpName_INS;
        }

        protected void dsCustomServiceDetail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridDetailDataList.DataBind();
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

        protected void dsCustomServiceDetail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridDetailDataList.DataBind();
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

        protected void dsCustomServiceDetail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            Label vServiceNo_INS = (Label)fvCustomServiceDetail.FindControl("eServiceNo_INS");
            if (vServiceNo_INS.Text.Trim() != "")
            {
                TextBox vContactDate_INS = (TextBox)fvCustomServiceDetail.FindControl("eContactDate_INS");
                TextBox vAssignDate_INS = (TextBox)fvCustomServiceDetail.FindControl("eAssignDate_INS");
                string vSQLStr_INS = "";
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vServiceNo = vServiceNo_INS.Text.Trim();
                int vItemsMax = 0;
                vSQLStr_INS = "select Isnull(MAX(Items), 0) ItemsMax from CustomServiceDetail where ServiceNo = '" + vServiceNo + "' ";
                vItemsMax = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_INS, "ItemsMax"));
                string vItemsStr = (vItemsMax + 1).ToString("D4");
                string vServiceNoItem = vServiceNo + vItemsStr;
                e.Command.Parameters["@ServiceNo"].Value = vServiceNo;
                e.Command.Parameters["@Items"].Value = vItemsStr;
                e.Command.Parameters["@ServiceNoItem"].Value = vServiceNoItem;
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
                Response.Write("alert('客訴單號空白，請先選擇客訴單！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void dsCustomServiceDetail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridDetailDataList.DataBind();
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

        protected void dsCustomServiceDetail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            Label vServiceNoItem_Edit = (Label)fvCustomServiceDetail.FindControl("eServiceNoItem_Edit");
            string vServiceNoItem = vServiceNoItem_Edit.Text.Trim();
            if (vServiceNoItem != "")
            {
                TextBox vContactDate_Edit = (TextBox)fvCustomServiceDetail.FindControl("eContactDate_Edit");
                TextBox vAssignDate_Edit = (TextBox)fvCustomServiceDetail.FindControl("eAssignDate_Edit");
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

        protected void ddlIsTrue_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eIsTrue_Search.Text = ddlIsTrue_Search.SelectedValue.Trim();
        }
    }
}