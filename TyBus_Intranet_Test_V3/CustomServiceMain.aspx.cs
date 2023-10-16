using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class CustomServiceMain : System.Web.UI.Page
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

                vTempDate_URL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_Start_Search.ClientID;
                vTempDate_Script = "window.open('" + vTempDate_URL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eBuildDate_Start_Search.Attributes["onClick"] = vTempDate_Script;

                vTempDate_URL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate_End_Search.ClientID;
                vTempDate_Script = "window.open('" + vTempDate_URL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eBuildDate_End_Search.Attributes["onClick"] = vTempDate_Script;

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
                    else
                    {
                        ServiceDataBind();
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

            string vWStr_AthorityDep = (eAthorityDep_Search.Text.Trim() != "") ? " and (AthorityDepNo = '" + eAthorityDep_Search.Text.Trim() + "' or AthorityDepNo2 = '" + eAthorityDep_Search.Text.Trim() + "') " + Environment.NewLine : "";

            string vWStr_ClosedOnly = (rbIsClosed.SelectedValue == "1") ? " and isnull(IsClosed, 0) = 1 " + Environment.NewLine :
                                      (rbIsClosed.SelectedValue == "2") ? " and isnull(IsClosed, 0) = 0 " + Environment.NewLine :
                                      "";

            string vWStr_CaseSource = (eCaseSource_Search.Text.Trim() != "") ? " and CaseSource = '" + eCaseSource_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vWStr_CivicName = (eCivicName_Search.Text.Trim() != "") ? " and CivicName like '%" + eCivicName_Search.Text.Trim() + "%'" + Environment.NewLine : "";

            string vWStr_CivicTel = (eCivicTel_Search.Text.Trim() != "") ? " and (CivicTelNo Like '%" + eCivicTel_Search.Text.Trim() + "%' or CivicCellPhone Like '%" + eCivicTel_Search.Text.Trim() + "%' )" + Environment.NewLine : "";

            string vWStr_IsTrue = (eIsTrue_Search.Text.Trim() != "") ? " and IsTrue = '" + eIsTrue_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vResultStr = "select ServiceNo, BuildDate, BuildTime, BuildMan, (select [Name] from Employee where EmpNo = a.BuildMan) BuildMan_C, " + Environment.NewLine +
                                "       ServiceType, (select TypeText from CustomServiceType where TypeStep = '1' and TypeLevel1 = a.ServiceType) ServiceType_C, " + Environment.NewLine +
                                "       ServiceTypeB, (select TypeText from CustomServiceType where TypeStep = '2' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB) ServiceTypeB_C, " + Environment.NewLine +
                                "       ServiceTypeC, (select TypeText from CustomServiceType where TypeStep = '3' and TypeLevel1 = a.ServiceType and TypeLevel2 = a.ServiceTypeB and TypeLevel3 = a.ServiceTypeC) ServiceTypeC_C, " + Environment.NewLine +
                                "       LinesNo, Car_ID, Driver, DriverName, BoardTime, BoardStation, " + Environment.NewLine +
                                "       GetoffTime, GetoffStation, ServiceNote, IsReplied, IsPending, AssignDate, AssignMan, " + Environment.NewLine +
                                "       (select [Name] from Employee where EmpNo = a.AssignMan) AssignMan_C, IsClosed, CloseDate, CloseMan, " + Environment.NewLine +
                                "       (select [Name] from Employee where EmpNo = a.CloseMan) CloseMan_C, " + Environment.NewLine +
                                "       CaseSource, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = CAST('客服專線記錄表' AS char(16)) + CAST('fmCustomService' AS char(16)) + 'CaseSource') AND (CLASSNO = a.CaseSource)) CaseSource_C, " + Environment.NewLine +
                                "       ServiceDate, (select ClassTxt from DBDICB where FKey = '客服專線記錄表  fmCustomService IsTrue' and ClassNo = a.IsTrue) IsTrue_C " + Environment.NewLine +
                                "  from CustomService a " + Environment.NewLine +
                                " where isnull(ServiceNo, '') <> '' " + Environment.NewLine +
                                vWStr_BuildMan + vWStr_ServiceDate + vWStr_BuildDate + vWStr_ServiceType + vWStr_ServiceTypeB +
                                vWStr_ServiceTypeC + vWStr_LinesNo + vWStr_CarID + vWStr_Driver + vWStr_AthorityDep + vWStr_ClosedOnly + vWStr_CaseSource +
                                vWStr_CivicName + vWStr_CivicTel + vWStr_IsTrue +
                                //" order by a.BuildDate DESC "; //2023.08.30 修改
                                " order by a.ServiceNo DESC ";
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

        protected void bbPrint_Choised_Click(object sender, EventArgs e)
        {
            if (Session["CustomServicePrint"].ToString().Length > 0)
            {
                Session["ServicePrintMode"] = "1";
                Response.Redirect("CustomServicePrintDetail.aspx");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先挑選要列印的客訴單！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbPrint_All_Click(object sender, EventArgs e)
        {
            Session["ServicePrintMode"] = "2";
            Session["PrintAllSelected"] = "";
            string vWStr_BuildMan = (eBuildMan_Searech.Text.Trim() != "") ? " and BuildMan = '" + eBuildMan_Searech.Text.Trim() + "' " : "";
            string vWStr_BuildDate = ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                     " and BuildDate between '" + (DateTime.Parse(eBuildDate_Start_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_Start_Search.Text.Trim()).ToString("MM/dd") +
                                     "' and '" + (DateTime.Parse(eBuildDate_End_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " :
                                     ((eBuildDate_Start_Search.Text.Trim() != "") && (eBuildDate_End_Search.Text.Trim() == "")) ?
                                     " and BuildDate = '" + (DateTime.Parse(eBuildDate_Start_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' " :
                                     ((eBuildDate_Start_Search.Text.Trim() == "") && (eBuildDate_End_Search.Text.Trim() != "")) ?
                                     " and BuildDate = '" + (DateTime.Parse(eBuildDate_End_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eBuildDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " : "";
            string vWStr_ServiceDate = ((eServiceDate_Start_Search.Text.Trim() != "") && (eServiceDate_End_Search.Text.Trim() != "")) ?
                                     " and ServiceDate between '" + (DateTime.Parse(eServiceDate_Start_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_Start_Search.Text.Trim()).ToString("MM/dd") +
                                     "' and '" + (DateTime.Parse(eServiceDate_End_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " :
                                     ((eServiceDate_Start_Search.Text.Trim() != "") && (eServiceDate_End_Search.Text.Trim() == "")) ?
                                     " and ServiceDate = '" + (DateTime.Parse(eServiceDate_Start_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_Start_Search.Text.Trim()).ToString("MM/dd") + "' " :
                                     ((eServiceDate_Start_Search.Text.Trim() == "") && (eServiceDate_End_Search.Text.Trim() != "")) ?
                                     " and ServiceDate = '" + (DateTime.Parse(eServiceDate_End_Search.Text.Trim()).Year).ToString("D4") + "/" + DateTime.Parse(eServiceDate_End_Search.Text.Trim()).ToString("MM/dd") + "' " : "";
            string vWStr_ServiceType = (eServiceType_Search.Text.Trim() != "") ? " and ServiceType = '" + eServiceType_Search.Text.Trim() + "' " : "";
            string vWStr_ServiceTypeB = (eServiceTypeB_Search.Text.Trim() != "") ? " and ServiceTypeB = '" + eServiceTypeB_Search.Text.Trim().Substring(3) + "' " : "";
            string vWStr_ServiceTypeC = (eServiceTypeC_Search.Text.Trim() != "") ? " and ServiceTypeC = '" + eServiceTypeC_Search.Text.Trim().Substring(5) + "' " : "";
            string vWStr_LinesNo = (eLinesNo_Search.Text.Trim() != "") ? " and LinesNo like '" + eLinesNo_Search.Text.Trim() + "%' " : "";
            string vWStr_CarID = (eCarID_Search.Text.Trim() != "") ? " and Car_ID like '" + eCarID_Search.Text.Trim() + "' " : "";
            string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? " and (Driver like '" + eDriver_Search.Text.Trim() + "%' or DriverName like '" + eDriver_Search.Text.Trim() + "%')" : "";
            string vWStr_AthorityDep = (eAthorityDep_Search.Text.Trim() != "") ? " and AthorityDepNo = '" + eAthorityDep_Search.Text.Trim() + "' " : "";
            string vWStr_ClosedOnly = (rbIsClosed.SelectedValue == "1") ? " and isnull(IsClosed, 0) = 1 " :
                                      (rbIsClosed.SelectedValue == "2") ? " and isnull(IsClosed, 0) = 0 " :
                                      "";
            string vWStr_CivicName = (eCivicName_Search.Text.Trim() != "") ? " and CivicName like '%" + eCivicName_Search.Text.Trim() + "%'" : "";
            string vWStr_CivicTel = (eCivicTel_Search.Text.Trim() != "") ? " and (CivicTelNo like '%" + eCivicTel_Search.Text.Trim() + "%' or CivicCellPhone Like '%" + eCivicTel_Search.Text.Trim() + "%' )" : "";
            string vWStr_CaseSource = (eCaseSource_Search.Text.Trim() != "") ? " and CaseSource = '" + eCaseSource_Search.Text.Trim() + "' " : "";
            //string vOrderStr = " order by BuildDate DESC";
            string vOrderStr = " order by ServiceNo DESC"; //2023.08.30 修改
            string vSelectStr = vWStr_BuildMan + vWStr_BuildDate + vWStr_ServiceDate + vWStr_ServiceType + vWStr_ServiceTypeB + vWStr_ServiceTypeC +
                                vWStr_LinesNo + vWStr_CarID + vWStr_Driver + vWStr_AthorityDep + vWStr_ClosedOnly + vWStr_CaseSource +
                                vWStr_CivicName + vWStr_CivicTel + vOrderStr;
            Session["PriineAllSelected"] = vSelectStr;
            Response.Redirect("CustomServicePrintDetail.aspx");
        }

        protected void cbPrint_Detail_CheckedChanged(object sender, EventArgs e)
        {
            string vServiceNo = "";
            int vFieldIndex = 0;
            for (int i = 0; i < gridDataList.Rows.Count; i++)
            {
                for (int j = 0; j < gridDataList.Rows[i].Cells.Count; j++)
                {
                    if (gridDataList.Columns[j].HeaderText == "客訴單號")
                    {
                        vFieldIndex = j;
                    }
                }
                CheckBox cbTemp_Print = (CheckBox)gridDataList.Rows[i].FindControl("cbPrint_Detail");
                vServiceNo = gridDataList.Rows[i].Cells[vFieldIndex].Text.Trim();

                if (cbTemp_Print.Checked)
                {
                    if (Session["CustomServicePrint"].ToString().IndexOf(vServiceNo, 0) == -1)
                    {
                        if ((Session["CustomServicePrint"].ToString().Trim() == "-1") || (Session["CustomServicePrint"].ToString().Trim() == ""))
                        {
                            Session["CustomServicePrint"] = "'" + vServiceNo + "'";
                        }
                        else
                        {
                            Session["CustomServicePrint"] = Session["CustomServicePrint"].ToString().Trim() + "," + "'" + vServiceNo + "'";
                        }
                    }
                }
                else
                {
                    if (Session["CustomServicePrint"].ToString().Trim() != "-1")
                    {
                        Session["CustomServicePrint"] = Session["CustomServicePrint"].ToString().Trim().Replace("'" + vServiceNo + "'", "");
                        Session["CustomServicePrint"] = Session["CustomServicePrint"].ToString().Trim().Replace(",,", ",");
                        if ((Session["CustomServicePrint"].ToString().Length > 0) && (Session["CustomServicePrint"].ToString().Substring(0, 1) == ","))
                        {
                            Session["CustomServicePrint"] = Session["CustomServicePrint"].ToString().Substring(1, Session["CustomServicePrint"].ToString().Length - 1);
                        }
                        else if ((Session["CustomServicePrint"].ToString().Length > 0) && (Session["CustomServicePrint"].ToString().Substring(Session["CustomServicePrint"].ToString().Length - 2, 1) == ","))
                        {
                            Session["CustomServicePrint"] = Session["CustomServicePrint"].ToString().Substring(0, Session["CustomServicePrint"].ToString().Length - 1);
                        }
                    }
                }
            }
        }

        protected void fvServiceDataView_DataBound(object sender, EventArgs e)
        {
            string vDateURL_Temp = "";
            string vDateScript_Temp = "";
            string vSQLStr = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr((Request.ApplicationPath));
            }

            switch (fvServiceDataView.CurrentMode)
            {
                case FormViewMode.Edit:
                    plSearch.Visible = false;
                    gridDataList.Visible = true;
                    TextBox lbTemp_AssignDate_Edit = (TextBox)fvServiceDataView.FindControl("eAssignDate_Edit");
                    TextBox lbTemp_ServiceDate_Edit = (TextBox)fvServiceDataView.FindControl("eServiceDate_Edit");

                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + lbTemp_AssignDate_Edit.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    lbTemp_AssignDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + lbTemp_ServiceDate_Edit.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    lbTemp_ServiceDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                    Label lbTemp_Edit = (Label)fvServiceDataView.FindControl("eCaseSource_Edit");
                    DropDownList ddlTemp_Edit = (DropDownList)fvServiceDataView.FindControl("ddlCaseSource_Edit");
                    ddlTemp_Edit.DataBind();
                    for (int iCaseIndex = 0; iCaseIndex < ddlTemp_Edit.Items.Count; iCaseIndex++)
                    {
                        if (ddlTemp_Edit.Items[iCaseIndex].Value.Trim() == lbTemp_Edit.Text.Trim())
                        {
                            ddlTemp_Edit.SelectedIndex = iCaseIndex;
                        }
                    }

                    DropDownList ddlServiceType_Edit = (DropDownList)fvServiceDataView.FindControl("ddlServiceType_Edit");
                    if (ddlServiceType_Edit != null)
                    {
                        Label eServiceType_Edit = (Label)fvServiceDataView.FindControl("eServiceType_Edit");
                        DropDownList ddlServiceTypeB_Edit = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeB_Edit");
                        Label eServiceTypeB_Edit = (Label)fvServiceDataView.FindControl("eServiceTypeB_Edit");
                        DropDownList ddlServiceTypeC_Edit = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeC_Edit");
                        Label eServiceTypeC_Edit = (Label)fvServiceDataView.FindControl("eServiceTypeC_Edit");
                        string vTypeNo = eServiceType_Edit.Text.Trim();
                        string vTypeNoB = (eServiceTypeB_Edit.Text.Trim() != "") ? vTypeNo + eServiceTypeB_Edit.Text.Trim() : "";
                        string vTypeNoC = (eServiceTypeC_Edit.Text.Trim() != "") ? vTypeNoB + eServiceTypeC_Edit.Text.Trim() : "";
                        ddlServiceType_Edit.SelectedIndex = ddlServiceType_Edit.Items.IndexOf(ddlServiceType_Edit.Items.FindByValue(vTypeNo));
                        using (SqlConnection connTypeNo = new SqlConnection(vConnStr))
                        {
                            vSQLStr = "SELECT CAST(NULL AS varchar(6)) AS TypeNo, CAST(NULL AS varchar) AS TypeText " + Environment.NewLine +
                                      " UNION ALL " + Environment.NewLine +
                                      "SELECT LEFT (TypeNo, 6) AS TypeNo, TypeText " + Environment.NewLine +
                                      "  FROM CustomServiceType " + Environment.NewLine +
                                      " WHERE (LEFT (TypeNo, 3) = '" + vTypeNo + "') " + Environment.NewLine +
                                      "   AND (TypeStep = '2') " + Environment.NewLine +
                                      " ORDER BY TypeNo";
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTypeNo);
                            connTypeNo.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            ddlServiceTypeB_Edit.Items.Clear();
                            while (drTemp.Read())
                            {
                                ddlServiceTypeB_Edit.Items.Add(new ListItem(drTemp["TypeText"].ToString().Trim(), drTemp["TypeNo"].ToString().Trim()));
                            }
                            ddlServiceTypeB_Edit.SelectedIndex = ddlServiceTypeB_Edit.Items.IndexOf(ddlServiceTypeB_Edit.Items.FindByValue(vTypeNoB));

                            vSQLStr = "SELECT CAST(NULL AS varchar(9)) AS TypeNo, CAST(NULL AS varchar) AS TypeText " + Environment.NewLine +
                                      " UNION ALL " + Environment.NewLine +
                                      "SELECT TypeNo, TypeText " + Environment.NewLine +
                                      "  FROM CustomServiceType " + Environment.NewLine +
                                      " WHERE (LEFT (TypeNo, 6) = '" + vTypeNoB + "') " + Environment.NewLine +
                                      "   AND (TypeStep = '3') " + Environment.NewLine +
                                      " ORDER BY TypeNo";
                            drTemp.Close();
                            cmdTemp.Cancel();
                            connTypeNo.Close();
                            cmdTemp.CommandText = vSQLStr;
                            connTypeNo.Open();
                            drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ddlServiceTypeC_Edit.Items.Add(new ListItem(drTemp["TypeText"].ToString().Trim(), drTemp["TypeNo"].ToString().Trim()));
                            }
                            ddlServiceTypeC_Edit.SelectedIndex = ddlServiceTypeC_Edit.Items.IndexOf(ddlServiceTypeC_Edit.Items.FindByValue(vTypeNoC));
                        }
                    }
                    DropDownList ddlIsTrue_Edit = (DropDownList)fvServiceDataView.FindControl("ddlIsTrue_Edit");
                    Label eIsTrue_Edit = (Label)fvServiceDataView.FindControl("eIstrue_Edit");
                    if (ddlIsTrue_Edit != null)
                    {
                        using (SqlConnection connIsTrue_Edit = new SqlConnection(vConnStr))
                        {
                            vSQLStr = "select Cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                      " union all " + Environment.NewLine +
                                      "select CLassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                      " where FKey = '客服專線記錄表  fmCustomService IsTrue' " + Environment.NewLine +
                                      "   and ClassNo <> 'IT02' " + Environment.NewLine +
                                      " order by CLassNo";
                            SqlCommand cmdIsTrue_Edit = new SqlCommand(vSQLStr, connIsTrue_Edit);
                            connIsTrue_Edit.Open();
                            SqlDataReader drIsTrue_Edit = cmdIsTrue_Edit.ExecuteReader();
                            ddlIsTrue_Edit.Items.Clear();
                            while (drIsTrue_Edit.Read())
                            {
                                ddlIsTrue_Edit.Items.Add(new ListItem(drIsTrue_Edit["ClassTxt"].ToString().Trim(), drIsTrue_Edit["ClassNo"].ToString().Trim()));
                            }
                        }
                        ddlIsTrue_Edit.SelectedIndex = (eIsTrue_Edit.Text.Trim() != "") ?
                                                       ddlIsTrue_Edit.Items.IndexOf(ddlIsTrue_Edit.Items.FindByValue(eIsTrue_Edit.Text.Trim())) :
                                                       0;
                    }
                    break;

                case FormViewMode.Insert:
                    plSearch.Visible = false;
                    gridDataList.Visible = false;
                    Label lbTemp_LoginID = (Label)fvServiceDataView.FindControl("eBuildMan_Insert");
                    Label lbTemp_LoginName = (Label)fvServiceDataView.FindControl("eBuildManName_Insert");
                    TextBox lbTemp_BuildDate_INS = (TextBox)fvServiceDataView.FindControl("eBuildDate_Insert");
                    TextBox lbTemp_BuildTime = (TextBox)fvServiceDataView.FindControl("eBuildTime_Insert");
                    TextBox lbTemp_AssignDate_INS = (TextBox)fvServiceDataView.FindControl("eAssignDate_Insert");
                    TextBox lbTemp_ServiceDate_INS = (TextBox)fvServiceDataView.FindControl("eServiceDate_Insert");
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + lbTemp_BuildDate_INS.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    lbTemp_BuildDate_INS.Attributes["onClick"] = vDateScript_Temp;

                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + lbTemp_AssignDate_INS.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    lbTemp_AssignDate_INS.Attributes["onClick"] = vDateScript_Temp;

                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + lbTemp_ServiceDate_INS.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    lbTemp_ServiceDate_INS.Attributes["onClick"] = vDateScript_Temp;


                    lbTemp_LoginID.Text = vLoginID;
                    lbTemp_LoginName.Text = Session["LoginName"].ToString().Trim();
                    lbTemp_BuildDate_INS.Text = (DateTime.Today.Year - 1911).ToString("D3") + "/" + DateTime.Today.ToString("MM/dd");
                    lbTemp_ServiceDate_INS.Text = (DateTime.Today.Year - 1911).ToString("D3") + "/" + DateTime.Today.ToString("MM/dd");
                    lbTemp_BuildTime.Text = DateTime.Now.ToString("HH:mm");
                    DropDownList ddlTemp_INS = (DropDownList)fvServiceDataView.FindControl("ddlCaseSource_Insert");
                    ddlTemp_INS.DataBind();
                    Label lbTemp_INS = (Label)fvServiceDataView.FindControl("eCaseSource_Insert");
                    ddlTemp_INS.SelectedIndex = 3;
                    lbTemp_INS.Text = "CS03";
                    Label eServiceType_INS = (Label)fvServiceDataView.FindControl("eServiceType_Insert");
                    Label eServiceTypeB_INS = (Label)fvServiceDataView.FindControl("eServiceTypeB_Insert");
                    Label eServiceTypeC_INS = (Label)fvServiceDataView.FindControl("eServiceTypeC_Insert");
                    eServiceType_INS.Text = "";
                    eServiceTypeB_INS.Text = "";
                    eServiceTypeC_INS.Text = "";
                    DropDownList ddlIsTrue_Insert = (DropDownList)fvServiceDataView.FindControl("ddlIsTrue_Insert");
                    Label eIsTrue_Insert = (Label)fvServiceDataView.FindControl("eIstrue_Insert");
                    if (ddlIsTrue_Insert != null)
                    {
                        using (SqlConnection connIsTrue_Insert = new SqlConnection(vConnStr))
                        {
                            vSQLStr = "select Cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                      " union all " + Environment.NewLine +
                                      "select CLassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                      " where FKey = '客服專線記錄表  fmCustomService IsTrue' " + Environment.NewLine +
                                      "   and ClassNo <> 'IT02' " + Environment.NewLine +
                                      " order by CLassNo";
                            SqlCommand cmdIsTrue_Insert = new SqlCommand(vSQLStr, connIsTrue_Insert);
                            connIsTrue_Insert.Open();
                            SqlDataReader drIsTrue_Insert = cmdIsTrue_Insert.ExecuteReader();
                            ddlIsTrue_Insert.Items.Clear();
                            while (drIsTrue_Insert.Read())
                            {
                                ddlIsTrue_Insert.Items.Add(new ListItem(drIsTrue_Insert["ClassTxt"].ToString().Trim(), drIsTrue_Insert["ClassNo"].ToString().Trim()));
                            }
                        }
                        ddlIsTrue_Insert.SelectedIndex = 0;
                        eIsTrue_Insert.Text = "";
                    }
                    break;

                case FormViewMode.ReadOnly:
                    plSearch.Visible = true;
                    gridDataList.Visible = true;
                    Label lbTemp_IsClosed = (Label)fvServiceDataView.FindControl("eCloseDate_List");
                    Button bbTemp_Close = (Button)fvServiceDataView.FindControl("bbClosed");
                    Button bbTemp_Reopen = (Button)fvServiceDataView.FindControl("bbReopen");
                    Label eIsTrue_List = (Label)fvServiceDataView.FindControl("eIsTrue_List");
                    DropDownList ddlIsTrue_List = (DropDownList)fvServiceDataView.FindControl("ddlIsTrue_List");
                    if (bbTemp_Close != null)
                    {
                        bbTemp_Close.Visible = lbTemp_IsClosed.Text.Trim() == "";
                    }
                    if (bbTemp_Reopen != null)
                    {
                        bbTemp_Reopen.Visible = lbTemp_IsClosed.Text.Trim() != "";
                    }
                    if ((eIsTrue_List != null) && (ddlIsTrue_List != null))
                    {
                        using (SqlConnection connIsTrue_List = new SqlConnection(vConnStr))
                        {
                            vSQLStr = "select Cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                      " union all " + Environment.NewLine +
                                      "select CLassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                      " where FKey = '客服專線記錄表  fmCustomService IsTrue' " + Environment.NewLine +
                                      "   and ClassNo <> 'IT02' " + Environment.NewLine +
                                      " order by CLassNo";
                            SqlCommand cmdIsTrue_List = new SqlCommand(vSQLStr, connIsTrue_List);
                            connIsTrue_List.Open();
                            SqlDataReader drIsTrue_List = cmdIsTrue_List.ExecuteReader();
                            ddlIsTrue_List.Items.Clear();
                            while (drIsTrue_List.Read())
                            {
                                ddlIsTrue_List.Items.Add(new ListItem(drIsTrue_List["ClassTxt"].ToString().Trim(), drIsTrue_List["ClassNo"].ToString().Trim()));
                            }
                        }
                        ddlIsTrue_List.SelectedIndex = (eIsTrue_List.Text.Trim() != "") ?
                                                       ddlIsTrue_List.Items.IndexOf(ddlIsTrue_List.Items.FindByValue(eIsTrue_List.Text.Trim())) :
                                                       0;
                    }
                    break;
            }
        }

        protected void ddlCaseSource_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlCaseSource_Edit");
            Label lbTemp = (Label)fvServiceDataView.FindControl("eCaseSource_Edit");
            lbTemp.Text = ddlTemp.SelectedValue;
        }

        protected void ddlServiceType_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label lbTemp = (Label)fvServiceDataView.FindControl("eServiceType_Edit");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlServiceType_Edit");
            lbTemp.Text = ddlTemp.SelectedValue;
            //Session["TypeNo_A"] = ddlTemp.SelectedValue;
            Label lbTempB = (Label)fvServiceDataView.FindControl("eServiceTypeB_Edit");
            DropDownList ddlTempB = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeB_Edit");
            if (ddlTempB != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                ddlTempB.Items.Clear();
                lbTempB.Text = "";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    string vSQLStr = "SELECT CAST(NULL AS varchar(6)) AS TypeNo, CAST(NULL AS varchar) AS TypeText " + Environment.NewLine +
                                     " UNION ALL " + Environment.NewLine +
                                     "SELECT LEFT (TypeNo, 6) AS TypeNo, TypeText " + Environment.NewLine +
                                     "  FROM CustomServiceType " + Environment.NewLine +
                                     " WHERE (LEFT (TypeNo, 3) = '" + ddlTemp.SelectedValue.Trim() + "') " + Environment.NewLine +
                                     "   AND (TypeStep = '2') " + Environment.NewLine +
                                     " ORDER BY TypeNo";
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        ddlTempB.Items.Add(new ListItem(drTemp["TypeText"].ToString().Trim(), drTemp["TypeNo"].ToString().Trim()));
                    }
                }
            }
        }

        protected void ddlServiceTypeB_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label lbTemp = (Label)fvServiceDataView.FindControl("eServiceTypeB_Edit");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeB_Edit");
            lbTemp.Text = ddlTemp.SelectedValue.Substring(3, 3);
            //Session["TypeNo_B"] = ddlTemp.SelectedValue;
            Label lbTempC = (Label)fvServiceDataView.FindControl("eServiceTypeC_Edit");
            DropDownList ddlTempC = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeC_Edit");
            if (ddlTempC != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                ddlTempC.Items.Clear();
                lbTempC.Text = "";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    string vSQLStr = "SELECT CAST(NULL AS varchar(9)) AS TypeNo, CAST(NULL AS varchar) AS TypeText " + Environment.NewLine +
                                     " UNION ALL " + Environment.NewLine +
                                     "SELECT TypeNo, TypeText " + Environment.NewLine +
                                     "  FROM CustomServiceType " + Environment.NewLine +
                                     " WHERE (LEFT (TypeNo, 6) = '" + ddlTemp.SelectedValue.Trim() + "') " + Environment.NewLine +
                                     "   AND (TypeStep = '3') " + Environment.NewLine +
                                     " ORDER BY TypeNo";
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        ddlTempC.Items.Add(new ListItem(drTemp["TypeText"].ToString().Trim(), drTemp["TypeNo"].ToString().Trim()));
                    }
                }
            }
        }

        protected void ddlServiceTypeC_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label lbTemp = (Label)fvServiceDataView.FindControl("eServiceTypeC_Edit");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeC_Edit");
            lbTemp.Text = ddlTemp.SelectedValue.Substring(6, 3);
        }

        protected void ddlAthorityDepNo_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label tbTemp = (Label)fvServiceDataView.FindControl("eAthorityDepNo_Edit");
            TextBox tbTemp_Note = (TextBox)fvServiceDataView.FindControl("eAthorityDepNote_Edit");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlAthorityDepNo_Edit");
            tbTemp.Text = ddlTemp.SelectedValue;
            tbTemp_Note.Text = (ddlTemp.SelectedValue == "99") ? "" : ddlTemp.SelectedItem.Text.Trim();
            tbTemp_Note.Enabled = (ddlTemp.SelectedValue == "99");
        }

        protected void eDriverName_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eTempDriverName = (TextBox)fvServiceDataView.FindControl("eDriverName_Edit");
            Label eTempDriver = (Label)fvServiceDataView.FindControl("eDriver_Edit");
            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            string vSQLStr = "select EmpNo from Employee where LeaveDay is null and [Name] = '" + eTempDriverName.Text.Trim() + "' ";
            eTempDriver.Text = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
        }

        protected void ddlAthorityDepNo2_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label tbTemp = (Label)fvServiceDataView.FindControl("eAthorityDepNo2_Edit");
            TextBox tbTemp_Note = (TextBox)fvServiceDataView.FindControl("eAthorityDepNote2_Edit");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlAthorityDepNo2_Edit");
            tbTemp.Text = ddlTemp.SelectedValue;
            tbTemp_Note.Text = (ddlTemp.SelectedValue == "99") ? "" : ddlTemp.SelectedItem.Text.Trim();
            tbTemp_Note.Enabled = (ddlTemp.SelectedValue == "99");
        }

        protected void eDriverName2_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eTempDriverName = (TextBox)fvServiceDataView.FindControl("eDriverName2_Edit");
            Label eTempDriver = (Label)fvServiceDataView.FindControl("eDriver2_Edit");
            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            string vSQLStr = "select EmpNo from Employee where LeaveDay is null and [Name] = '" + eTempDriverName.Text.Trim() + "' ";
            eTempDriver.Text = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
        }

        protected void ddlIsTrue_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlIsTrue_Edit = (DropDownList)fvServiceDataView.FindControl("ddlIsTrue_Edit");
            Label eIsTrue_Edit = (Label)fvServiceDataView.FindControl("eIsTrue_Edit");
            if (ddlIsTrue_Edit != null)
            {
                eIsTrue_Edit.Text = ddlIsTrue_Edit.SelectedValue.Trim();
            }
        }

        protected void eAssignMan_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eAssignMan = (TextBox)fvServiceDataView.FindControl("eAssignMan_Edit");
            Label eAssignName = (Label)fvServiceDataView.FindControl("eAssignName_Edit");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vAssignMan = eAssignMan.Text.Trim();
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
            string vAssignName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vAssignName == "")
            {
                vAssignName = vAssignMan;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vAssignName + "' ";
                vAssignMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eAssignMan.Text = vAssignMan;
            eAssignName.Text = vAssignName;
        }

        protected void ddlCaseSource_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlCaseSource_Insert");
            Label lbTemp = (Label)fvServiceDataView.FindControl("eCaseSource_Insert");
            lbTemp.Text = ddlTemp.SelectedValue;
        }

        protected void ddlServiceType_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label tbTemp = (Label)fvServiceDataView.FindControl("eServiceType_Insert");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlServiceType_Insert");
            tbTemp.Text = ddlTemp.SelectedValue;
            //Session["TypeNo_A"] = ddlTemp.SelectedValue;
            Label lbTempB = (Label)fvServiceDataView.FindControl("eServiceTypeB_Insert");
            DropDownList ddlTempB = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeB_Insert");
            if (ddlTempB != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                ddlTempB.Items.Clear();
                lbTempB.Text = "";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    string vSQLStr = "SELECT CAST(NULL AS varchar(6)) AS TypeNo, CAST(NULL AS varchar) AS TypeText " + Environment.NewLine +
                                     " UNION ALL " + Environment.NewLine +
                                     "SELECT LEFT (TypeNo, 6) AS TypeNo, TypeText " + Environment.NewLine +
                                     "  FROM CustomServiceType " + Environment.NewLine +
                                     " WHERE (LEFT (TypeNo, 3) = '" + ddlTemp.SelectedValue.Trim() + "') " + Environment.NewLine +
                                     "   AND (TypeStep = '2') " + Environment.NewLine +
                                     " ORDER BY TypeNo";
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        ddlTempB.Items.Add(new ListItem(drTemp["TypeText"].ToString().Trim(), drTemp["TypeNo"].ToString().Trim()));
                    }
                }
            }
        }

        protected void ddlServiceTypeB_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label tbTemp = (Label)fvServiceDataView.FindControl("eServiceTypeB_Insert");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeB_Insert");
            tbTemp.Text = ddlTemp.SelectedValue.Substring(3, 3);
            //Session["TypeNo_B"] = ddlTemp.SelectedValue;
            Label lbTempC = (Label)fvServiceDataView.FindControl("eServiceTypeC_Insert");
            DropDownList ddlTempC = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeC_Insert");
            if (ddlTempC != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                ddlTempC.Items.Clear();
                lbTempC.Text = "";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    string vSQLStr = "SELECT CAST(NULL AS varchar(9)) AS TypeNo, CAST(NULL AS varchar) AS TypeText " + Environment.NewLine +
                                     " UNION ALL " + Environment.NewLine +
                                     "SELECT TypeNo, TypeText " + Environment.NewLine +
                                     "  FROM CustomServiceType " + Environment.NewLine +
                                     " WHERE (LEFT (TypeNo, 6) = '" + ddlTemp.SelectedValue.Trim() + "') " + Environment.NewLine +
                                     "   AND (TypeStep = '3') " + Environment.NewLine +
                                     " ORDER BY TypeNo";
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        ddlTempC.Items.Add(new ListItem(drTemp["TypeText"].ToString().Trim(), drTemp["TypeNo"].ToString().Trim()));
                    }
                }
            }
        }

        protected void ddlServiceTypeC_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label tbTemp = (Label)fvServiceDataView.FindControl("eServiceTypeC_Insert");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlServiceTypeC_Insert");
            tbTemp.Text = ddlTemp.SelectedValue.Substring(6, 3);
        }

        protected void ddlAthorityDepNo_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label tbTemp = (Label)fvServiceDataView.FindControl("eAthorityDepNo_Insert");
            TextBox tbTemp_Note = (TextBox)fvServiceDataView.FindControl("eAthorityDepNote_Insert");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlAthorityDepNo_Insert");
            tbTemp.Text = ddlTemp.SelectedValue;
            tbTemp_Note.Text = (ddlTemp.SelectedValue == "99") ? "" : ddlTemp.SelectedItem.Text.Trim();
            tbTemp_Note.Enabled = (ddlTemp.SelectedValue == "99");
        }

        protected void eDriverName_Insert_TextChanged(object sender, EventArgs e)
        {
            TextBox eTempDriverName = (TextBox)fvServiceDataView.FindControl("eDriverName_Insert");
            Label eTempDriver = (Label)fvServiceDataView.FindControl("eDriver_Insert");
            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            string vSQLStr = "select EmpNo from Employee where LeaveDay is null and [Name] = '" + eTempDriverName.Text.Trim() + "' ";
            eTempDriver.Text = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
        }

        protected void ddlAthorityDepNo2_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label tbTemp = (Label)fvServiceDataView.FindControl("eAthorityDepNo2_Insert");
            TextBox tbTemp_Note = (TextBox)fvServiceDataView.FindControl("eAthorityDepNote2_Insert");
            DropDownList ddlTemp = (DropDownList)fvServiceDataView.FindControl("ddlAthorityDepNo2_Insert");
            tbTemp.Text = ddlTemp.SelectedValue;
            tbTemp_Note.Text = (ddlTemp.SelectedValue == "99") ? "" : ddlTemp.SelectedItem.Text.Trim();
            tbTemp_Note.Enabled = (ddlTemp.SelectedValue == "99");
        }

        protected void eDriverName2_Insert_TextChanged(object sender, EventArgs e)
        {
            TextBox eTempDriverName = (TextBox)fvServiceDataView.FindControl("eDriverName2_Insert");
            Label eTempDriver = (Label)fvServiceDataView.FindControl("eDriver2_Insert");
            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            string vSQLStr = "select EmpNo from Employee where LeaveDay is null and [Name] = '" + eTempDriverName.Text.Trim() + "' ";
            eTempDriver.Text = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
        }

        protected void ddlIsTrue_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlIstrue_Insert = (DropDownList)fvServiceDataView.FindControl("ddlIsTrue_Insert");
            Label eIsTrue_Insert = (Label)fvServiceDataView.FindControl("eIsTrue_Insert");
            eIsTrue_Insert.Text = ddlIstrue_Insert.SelectedValue.Trim();
        }

        protected void eAssignMan_Insert_TextChanged(object sender, EventArgs e)
        {
            TextBox eAssignMan = (TextBox)fvServiceDataView.FindControl("eAssignMan_Insert");
            Label eAssignName = (Label)fvServiceDataView.FindControl("eAssignName_Insert");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vAssignMan = eAssignMan.Text.Trim();
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
            string vAssignName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vAssignName == "")
            {
                vAssignName = vAssignMan;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vAssignName + "' ";
                vAssignMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eAssignMan.Text = vAssignMan;
            eAssignName.Text = vAssignName;
        }

        protected void bbDelete_Click(object sender, EventArgs e)
        {
            Label lbTemp_ClosedDate = (Label)fvServiceDataView.FindControl("eCloseDate_List");
            string vClosedDate = lbTemp_ClosedDate.Text.Trim();
            Label lbTemp_BuildMan = (Label)fvServiceDataView.FindControl("eBuildMan_List");
            string vBuildMan = lbTemp_BuildMan.Text.Trim();
            Label lbTemp_ServiceNo = (Label)fvServiceDataView.FindControl("eServiceNo_List");
            string vServiceNo = lbTemp_ServiceNo.Text.Trim();
            string vDelConnection = PF.GetConnectionStr(Request.ApplicationPath);
            //刪除資料前先複製一份到 History 區去
            string vDelCommandStr = "INSERT INTO [dbo].[CustomServiceHistory] " + Environment.NewLine +
                                    "           ([ServiceNo],[BuildDate],[BuildTime],[BuildMan],[ServiceType],[ServiceTypeB],[ServiceTypeC]," + Environment.NewLine +
                                    "            [LinesNo],[Car_ID],[Driver],[DriverName],[BoardTime],[BoardStation],[GetoffTime],[GetoffStation]," + Environment.NewLine +
                                    "            [ServiceNote],[IsNoContect],[CivicName],[CivicTelNo],[CivicCellPhone],[CivicAddress],[CivicEMail]," + Environment.NewLine +
                                    "            [AthorityDepNo],[AthorityDepNote],[IsReplied],[Remark],[IsPending],[AssignDate],[AssignMan]," + Environment.NewLine +
                                    "            [IsClosed],[CloseDate],[CloseMan],[ModifyType],[ModifyDate],[ModifyMan],[CivicTelExtNo],[LinesNo2]," + Environment.NewLine +
                                    "            [Car_ID2],[Driver2],[DriverName2],[AthorityDepNo2],[AthorityDepNote2],[CaseSource],[ServiceDate],[IsTrue])" + Environment.NewLine +
                                    "SELECT [ServiceNo],[BuildDate],[BuildTime],[BuildMan],[ServiceType],[ServiceTypeB],[ServiceTypeC]," + Environment.NewLine +
                                    "       [LinesNo],[Car_ID],[Driver],[DriverName],[BoardTime],[BoardStation],[GetoffTime],[GetoffStation]," + Environment.NewLine +
                                    "       [ServiceNote],[IsNoContect],[CivicName],[CivicTelNo],[CivicCellPhone],[CivicAddress],[CivicEMail]," + Environment.NewLine +
                                    "       [AthorityDepNo],[AthorityDepNote],[IsReplied],[Remark],[IsPending],[AssignDate],[AssignMan]," + Environment.NewLine +
                                    "       [IsClosed],[CloseDate],[CloseMan],'DEL',GetDate(),'" + vLoginID + "',[CivicTelExtNo],[LinesNo2], " + Environment.NewLine +
                                    "       [Car_ID2],[Driver2],[DriverName2],[AthorityDepNo2],[AthorityDepNote2],[CaseSource],[ServiceDate],[IsTrue] " + Environment.NewLine +
                                    "  from CustomService " + Environment.NewLine +
                                    " where ServiceNo = '" + vServiceNo + "'";
            try
            {
                PF.ExecSQL(vDelConnection, vDelCommandStr);
                dsServiceDataDetail.DeleteParameters.Clear();
                dsServiceDataDetail.DeleteParameters.Add(new Parameter("ServiceNo", DbType.String, vServiceNo));
                int vIsDeleted = dsServiceDataDetail.Delete();
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        protected void bbClosed_Click(object sender, EventArgs e)
        {
            string vServiceNo = "";
            string vSQLStr = "";
            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            Label lbTemp = (Label)fvServiceDataView.FindControl("eServiceNo_List");
            Label lbTemp_IsClosed = (Label)fvServiceDataView.FindControl("eCloseDate_List");
            if (lbTemp_IsClosed.Text.Trim() == "")
            {
                vServiceNo = lbTemp.Text.Trim();
                vSQLStr = "update CustomService " + Environment.NewLine +
                          "   set IsClosed = 1, CloseDate = '" + PF.TransDateString(DateTime.Today, "B") + "', " + Environment.NewLine +
                          "       CloseMan = '" + vLoginID + "' " + Environment.NewLine +
                          " where ServiceNo = '" + vServiceNo + "' ";
                PF.ExecSQL(vConnStr, vSQLStr);
                this.gridDataList.DataBind();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('本單己經結案，請勿重複結案')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbReopen_Click(object sender, EventArgs e)
        {
            string vServiceNo = "";
            string vSQLStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label lbTemp = (Label)fvServiceDataView.FindControl("eServiceNo_List");
            Label lbTemp_IsClosed = (Label)fvServiceDataView.FindControl("eCloseDate_List");
            if (lbTemp_IsClosed.Text.Trim() != "")
            {
                vServiceNo = lbTemp.Text.Trim();
                vSQLStr = "update CustomService " + Environment.NewLine +
                          "   set IsClosed = 0, CloseDate = null, CloseMan = null " + Environment.NewLine +
                          " where ServiceNo = '" + vServiceNo + "' ";
                PF.ExecSQL(vConnStr, vSQLStr);
                this.gridDataList.DataBind();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('本單尚未結案，請先進行結案！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbPrint_Single_Click(object sender, EventArgs e)
        {
            Label eServiceNo = (Label)fvServiceDataView.FindControl("eServiceNo_List");
            Session["CustomServicePrint"] = "'" + eServiceNo.Text.Trim() + "'";
            Session["ServicePrintMode"] = "1";
            Response.Redirect("CustomServicePrintDetail.aspx");
        }

        protected void dsServiceDataDetail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            /*改用全手寫，不呼叫內建命令
            string vConnStrTemp = PF.GetConnectionStr(Request.ApplicationPath);
            string vIndex = "";
            string vTempStr = "";
            string vServiceNo = "";
            TextBox tbTemp = (TextBox)fvServiceDataView.FindControl("eServiceType_Insert");
            TextBox tbTemp_BuildDate = (TextBox)fvServiceDataView.FindControl("eBuildDate_Insert");
            TextBox tbTemp_AssignDate = (TextBox)fvServiceDataView.FindControl("eAssignDate_Insert");
            TextBox tbTemp_ServiceDate = (TextBox)fvServiceDataView.FindControl("eServiceDate_Insert");
            DateTime vDateResult;
            DateTime vToday = DateTime.Today;
            string vSeriveNo_TypeA = (DateTime.TryParse(tbTemp_BuildDate.Text.Trim(), out vDateResult)) ?
                                     vDateResult.Year.ToString("D4") + vDateResult.ToString("MMdd") + tbTemp.Text.Trim() :
                                     vToday.Year.ToString("D4") + vToday.ToString("MMdd") + tbTemp.Text.Trim();
            vTempStr = "select RIGHT(MAX(ServiceNo), 3) [Index] from CustomService where ServiceNo like '" + vSeriveNo_TypeA + "%' ";
            vIndex = PF.GetValue(vConnStrTemp, vTempStr, "Index");
            vIndex = (vIndex != "") ? (int.Parse(vIndex) + 1).ToString("D3") : "001";
            vServiceNo = vSeriveNo_TypeA + vIndex;
            e.Command.Parameters["@ServiceNo"].Value = vServiceNo;
            e.Command.Parameters["@BuildDate"].Value = (tbTemp_BuildDate.Text.Trim() != "") ? DateTime.Parse(tbTemp_BuildDate.Text.Trim()) : DateTime.Today;
            e.Command.Parameters["@ServiceDate"].Value = (tbTemp_ServiceDate.Text.Trim() != "") ? DateTime.Parse(tbTemp_ServiceDate.Text.Trim()) : DateTime.Today;
            if ((tbTemp_AssignDate != null) && (tbTemp_AssignDate.Text.Trim() != ""))
            {
                e.Command.Parameters["@AssignDate"].Value = DateTime.Parse(tbTemp_AssignDate.Text.Trim());
            }
            */
        }

        protected void dsServiceDataDetail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridDataList.DataBind();
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

        protected void dsServiceDataDetail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridDataList.DataBind();
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

        protected void dsServiceDataDetail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridDataList.DataBind();
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

        protected void dsServiceDataDetail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            /* 全部改手寫，不使用 FormView 內建命令
            TextBox tbTemp_AssignDate = (TextBox)fvServiceDataView.FindControl("eAssignDate_Edit");
            TextBox tbTemp_ServiceDate = (TextBox)fvServiceDataView.FindControl("eServiceDate_Edit");

            if ((tbTemp_AssignDate != null) && (tbTemp_AssignDate.Text.Trim() != ""))
            {
                e.Command.Parameters["@AssignDate"].Value = DateTime.Parse(tbTemp_AssignDate.Text.Trim());
            }
            if ((tbTemp_ServiceDate != null) && (tbTemp_ServiceDate.Text.Trim() != ""))
            {
                e.Command.Parameters["@ServiceDate"].Value = DateTime.Parse(tbTemp_ServiceDate.Text.Trim());
            }
            */
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
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eServiceNo_Edit = (Label)fvServiceDataView.FindControl("eServiceNo_Edit");
            if ((eServiceNo_Edit != null) && (eServiceNo_Edit.Text.Trim() != ""))
            {
                Label eServiceType_Edit = (Label)fvServiceDataView.FindControl("eServiceType_Edit");
                Label eServiceTypeB_Edit = (Label)fvServiceDataView.FindControl("eServiceTypeB_Edit");
                Label eServiceTypeC_Edit = (Label)fvServiceDataView.FindControl("eServiceTypeC_Edit");
                TextBox eBoardTime_Edit = (TextBox)fvServiceDataView.FindControl("eBoardTime_Edit");
                TextBox eBoardStation_Edit = (TextBox)fvServiceDataView.FindControl("eBoardStation_Edit");
                TextBox eGetoffTime_Edit = (TextBox)fvServiceDataView.FindControl("eGetoffTime_Edit");
                TextBox eGetoffStation_Edit = (TextBox)fvServiceDataView.FindControl("eGetoffStation_Edit");
                TextBox eServiceNote_Edit = (TextBox)fvServiceDataView.FindControl("eServiceNote_Edit");
                CheckBox eIsNoContect_Edit = (CheckBox)fvServiceDataView.FindControl("eIsNoContect_Edit");
                TextBox eCivicName_Edit = (TextBox)fvServiceDataView.FindControl("eCivicName_Edit");
                TextBox eCivicTelNo_Edit = (TextBox)fvServiceDataView.FindControl("eCivicTelNo_Edit");
                TextBox eCivicCellPhone_Edit = (TextBox)fvServiceDataView.FindControl("eCivicCellPhone_Edit");
                TextBox eCivicAddress_Edit = (TextBox)fvServiceDataView.FindControl("eCivicAddress_Edit");
                TextBox eCivicEMail_Edit = (TextBox)fvServiceDataView.FindControl("eCivicEMail_Edit");
                TextBox eCivicTelExtNo_Edit = (TextBox)fvServiceDataView.FindControl("eCivicTelExtNo_Edit");
                TextBox eLinesNo_Edit = (TextBox)fvServiceDataView.FindControl("eLinesNo_Edit");
                TextBox eCar_ID_Edit = (TextBox)fvServiceDataView.FindControl("eCar_ID_Edit");
                Label eDriver_Edit = (Label)fvServiceDataView.FindControl("eDriver_Edit");
                TextBox eDriverName_Edit = (TextBox)fvServiceDataView.FindControl("eDriverName_Edit");
                Label eAthorityDepNo_Edit = (Label)fvServiceDataView.FindControl("eAthorityDepNo_Edit");
                TextBox eAthorityDepNote_Edit = (TextBox)fvServiceDataView.FindControl("eAthorityDepNote_Edit");
                TextBox eLinesNo2_Edit = (TextBox)fvServiceDataView.FindControl("eLinesNo2_Edit");
                TextBox eCar_ID2_Edit = (TextBox)fvServiceDataView.FindControl("eCar_ID2_Edit");
                Label eDriver2_Edit = (Label)fvServiceDataView.FindControl("eDriver2_Edit");
                TextBox eDriverName2_Edit = (TextBox)fvServiceDataView.FindControl("eDriverName2_Edit");
                Label eAthorityDepNo2_Edit = (Label)fvServiceDataView.FindControl("eAthorityDepNo2_Edit");
                TextBox eAthorityDepNote2_Edit = (TextBox)fvServiceDataView.FindControl("eAthorityDepNote2_Edit");
                CheckBox eIsReplied_Edit = (CheckBox)fvServiceDataView.FindControl("eIsReplied_Edit");
                TextBox eRemark_Edit = (TextBox)fvServiceDataView.FindControl("eRemark_Edit");
                CheckBox eIsPending_Edit = (CheckBox)fvServiceDataView.FindControl("eIsPending_Edit");
                TextBox eAssignDate_Edit = (TextBox)fvServiceDataView.FindControl("eAssignDate_Edit");
                TextBox eAssignMan_Edit = (TextBox)fvServiceDataView.FindControl("eAssignMan_Edit");
                Label eCaseSource_Edit = (Label)fvServiceDataView.FindControl("eCaseSource_Edit");
                TextBox eServiceDate_Edit = (TextBox)fvServiceDataView.FindControl("eServiceDate_Edit");
                Label eIsTrue_Edit = (Label)fvServiceDataView.FindControl("eIsTrue_Edit");

                dsServiceDataDetail.UpdateCommand = "UPDATE CustomService " + Environment.NewLine +
                                                    "   SET ServiceType = @ServiceType, ServiceTypeB = @ServiceTypeB, ServiceTypeC = @ServiceTypeC, " + Environment.NewLine +
                                                    "       BoardTime = @BoardTime, BoardStation = @BoardStation, GetoffTime = @GetoffTime, " + Environment.NewLine +
                                                    "       GetoffStation = @GetoffStation, ServiceNote = @ServiceNote, IsNoContect = @IsNoContect, " + Environment.NewLine +
                                                    "       CivicName = @CivicName, CivicTelNo = @CivicTelNo, CivicCellPhone = @CivicCellPhone, " + Environment.NewLine +
                                                    "       CivicAddress = @CivicAddress, CivicEMail = @CivicEMail, CivicTelExtNo = @CivicTelExtNo, " + Environment.NewLine +
                                                    "       LinesNo = @LinesNo, Car_ID = @Car_ID, Driver = @Driver, DriverName = @DriverName, " + Environment.NewLine +
                                                    "       AthorityDepNo = @AthorityDepNo, AthorityDepNote = @AthorityDepNote, " + Environment.NewLine +
                                                    "       LinesNo2 = @LinesNo2, Car_ID2 = @Car_ID2, Driver2 = @Driver2, DriverName2 = @DriverName2, " + Environment.NewLine +
                                                    "       AthorityDepNo2 = @AthorityDepNo2, AthorityDepNote2 = @AthorityDepNote2, " + Environment.NewLine +
                                                    "       IsReplied = @IsReplied, Remark = @Remark, IsPending = @IsPending, " + Environment.NewLine +
                                                    "       AssignDate = @AssignDate, AssignMan = @AssignMan, " + Environment.NewLine +
                                                    "       CaseSource = @CaseSource, ServiceDate = @ServiceDate, IsTrue = @IsTrue " + Environment.NewLine +
                                                    " WHERE (ServiceNo = @ServiceNo)";
                dsServiceDataDetail.UpdateParameters.Clear();
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("ServiceType", DbType.String, (eServiceType_Edit.Text.Trim() != "") ? eServiceType_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("ServiceTypeB", DbType.String, (eServiceTypeB_Edit.Text.Trim() != "") ? eServiceTypeB_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("ServiceTypeC", DbType.String, (eServiceTypeC_Edit.Text.Trim() != "") ? eServiceTypeC_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("BoardTime", DbType.String, (eBoardTime_Edit.Text.Trim() != "") ? eBoardTime_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("BoardStation", DbType.String, (eBoardStation_Edit.Text.Trim() != "") ? eBoardStation_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("GetoffTime", DbType.String, (eGetoffTime_Edit.Text.Trim() != "") ? eGetoffTime_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("GetoffStation", DbType.String, (eGetoffStation_Edit.Text.Trim() != "") ? eGetoffStation_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("ServiceNote", DbType.String, (eServiceNote_Edit.Text.Trim() != "") ? eServiceNote_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("IsNoContect", DbType.Boolean, (eIsNoContect_Edit.Checked == true) ? "true" : "false"));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("CivicName", DbType.String, (eCivicName_Edit.Text.Trim() != "") ? eCivicName_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("CivicTelNo", DbType.String, (eCivicTelNo_Edit.Text.Trim() != "") ? eCivicTelNo_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("CivicCellPhone", DbType.String, (eCivicCellPhone_Edit.Text.Trim() != "") ? eCivicCellPhone_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("CivicAddress", DbType.String, (eCivicAddress_Edit.Text.Trim() != "") ? eCivicAddress_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("CivicEMail", DbType.String, (eCivicEMail_Edit.Text.Trim() != "") ? eCivicEMail_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("CivicTelExtNo", DbType.String, (eCivicTelExtNo_Edit.Text.Trim() != "") ? eCivicTelExtNo_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("LinesNo", DbType.String, (eLinesNo_Edit.Text.Trim() != "") ? eLinesNo_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("Car_ID", DbType.String, (eCar_ID_Edit.Text.Trim() != "") ? eCar_ID_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("Driver", DbType.String, (eDriver_Edit.Text.Trim() != "") ? eDriver_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("DriverName", DbType.String, (eDriverName_Edit.Text.Trim() != "") ? eDriverName_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("AthorityDepNo", DbType.String, (eAthorityDepNo_Edit.Text.Trim() != "") ? eAthorityDepNo_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("AthorityDepNote", DbType.String, (eAthorityDepNote_Edit.Text.Trim() != "") ? eAthorityDepNote_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("LinesNo2", DbType.String, (eLinesNo2_Edit.Text.Trim() != "") ? eLinesNo2_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("Car_ID2", DbType.String, (eCar_ID2_Edit.Text.Trim() != "") ? eCar_ID2_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("Driver2", DbType.String, (eDriver2_Edit.Text.Trim() != "") ? eDriver2_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("DriverName2", DbType.String, (eDriverName2_Edit.Text.Trim() != "") ? eDriverName2_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("AthorityDepNo2", DbType.String, (eAthorityDepNo2_Edit.Text.Trim() != "") ? eAthorityDepNo2_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("AthorityDepNote2", DbType.String, (eAthorityDepNote2_Edit.Text.Trim() != "") ? eAthorityDepNote2_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("IsReplied", DbType.Boolean, (eIsReplied_Edit.Checked == true) ? "true" : "false"));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("IsPending", DbType.Boolean, (eIsPending_Edit.Checked == true) ? "true" : "false"));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("AssignDate", DbType.DateTime, (eAssignDate_Edit.Text.Trim() != "") ? eAssignDate_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan_Edit.Text.Trim() != "") ? eAssignMan_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("CaseSource", DbType.String, (eCaseSource_Edit.Text.Trim() != "") ? eCaseSource_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("ServiceDate", DbType.DateTime, (eServiceDate_Edit.Text.Trim() != "") ? eServiceDate_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("IsTrue", DbType.String, (eIsTrue_Edit.Text.Trim() != "") ? eIsTrue_Edit.Text.Trim() : String.Empty));
                dsServiceDataDetail.UpdateParameters.Add(new Parameter("ServiceNo", DbType.String, eServiceNo_Edit.Text.Trim()));
                dsServiceDataDetail.Update();
                fvServiceDataView.ChangeMode(FormViewMode.ReadOnly);
                fvServiceDataView.DataBind();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('客服單號資料不正確！')");
                Response.Write("</" + "Script>");
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
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eServiceNo_Insert = (Label)fvServiceDataView.FindControl("eServiceNo_Insert");
            if (eServiceNo_Insert != null)
            {
                string vConnStrTemp = PF.GetConnectionStr(Request.ApplicationPath);
                string vIndex = "";
                string vTempStr = "";
                string vServiceNo = "";

                TextBox eBuildDate_Insert = (TextBox)fvServiceDataView.FindControl("eBuildDate_Insert");
                TextBox eBuildTime_Insert = (TextBox)fvServiceDataView.FindControl("eBuildTime_Insert");
                Label eBuildMan_Insert = (Label)fvServiceDataView.FindControl("eBuildMan_Insert");
                Label eServiceType_Insert = (Label)fvServiceDataView.FindControl("eServiceType_Insert");
                Label eServiceTypeB_Insert = (Label)fvServiceDataView.FindControl("eServiceTypeB_Insert");
                Label eServiceTypeC_Insert = (Label)fvServiceDataView.FindControl("eServiceTypeC_Insert");
                TextBox eLinesNo_Insert = (TextBox)fvServiceDataView.FindControl("eLinesNo_Insert");
                TextBox eCar_ID_Insert = (TextBox)fvServiceDataView.FindControl("eCar_ID_Insert");
                Label eDriver_Insert = (Label)fvServiceDataView.FindControl("eDriver_Insert");
                TextBox eDriverName_Insert = (TextBox)fvServiceDataView.FindControl("eDriverName_Insert");
                Label eAthorityDepNo_Insert = (Label)fvServiceDataView.FindControl("eAthorityDepNo_Insert");
                TextBox eAthorityDepNote_Insert = (TextBox)fvServiceDataView.FindControl("eAthorityDepNote_Insert");
                TextBox eLinesNo2_Insert = (TextBox)fvServiceDataView.FindControl("eLinesNo2_Insert");
                TextBox eCar_ID2_Insert = (TextBox)fvServiceDataView.FindControl("eCar_ID2_Insert");
                Label eDriver2_Insert = (Label)fvServiceDataView.FindControl("eDriver2_Insert");
                TextBox eDriverName2_Insert = (TextBox)fvServiceDataView.FindControl("eDriverName2_Insert");
                Label eAthorityDepNo2_Insert = (Label)fvServiceDataView.FindControl("eAthorityDepNo2_Insert");
                TextBox eAthorityDepNote2_Insert = (TextBox)fvServiceDataView.FindControl("eAthorityDepNote2_Insert");
                TextBox eBoardTime_Insert = (TextBox)fvServiceDataView.FindControl("eBoardTime_Insert");
                TextBox eBoardStation_Insert = (TextBox)fvServiceDataView.FindControl("eBoardStation_Insert");
                TextBox eGetoffTime_Insert = (TextBox)fvServiceDataView.FindControl("eGetoffTime_Insert");
                TextBox eGetoffStation_Insert = (TextBox)fvServiceDataView.FindControl("eGetoffStation_Insert");
                TextBox eServiceNote_Insert = (TextBox)fvServiceDataView.FindControl("eServiceNote_Insert");
                CheckBox eIsNoContect_Insert = (CheckBox)fvServiceDataView.FindControl("eIsNoContect_Insert");
                TextBox eCivicName_Insert = (TextBox)fvServiceDataView.FindControl("eCivicName_Insert");
                TextBox eCivicTelNo_Insert = (TextBox)fvServiceDataView.FindControl("eCivicTelNo_Insert");
                TextBox eCivicTelExtNo_Insert = (TextBox)fvServiceDataView.FindControl("eCivicTelExtNo_Insert");
                TextBox eCivicCellPhone_Insert = (TextBox)fvServiceDataView.FindControl("eCivicCellPhone_Insert");
                TextBox eCivicAddress_Insert = (TextBox)fvServiceDataView.FindControl("eCivicAddress_Insert");
                TextBox eCivicEMail_Insert = (TextBox)fvServiceDataView.FindControl("eCivicEMail_Insert");
                CheckBox eIsReplied_Insert = (CheckBox)fvServiceDataView.FindControl("eIsReplied_Insert");
                TextBox eRemark_Insert = (TextBox)fvServiceDataView.FindControl("eRemark_Insert");
                CheckBox eIsPending_Insert = (CheckBox)fvServiceDataView.FindControl("eIsPending_Insert");
                TextBox eAssignDate_Insert = (TextBox)fvServiceDataView.FindControl("eAssignDate_Insert");
                TextBox eAssignMan_Insert = (TextBox)fvServiceDataView.FindControl("eAssignMan_Insert");
                Label eCaseSource_Insert = (Label)fvServiceDataView.FindControl("eCaseSource_Insert");
                TextBox eServiceDate_Insert = (TextBox)fvServiceDataView.FindControl("eServiceDate_Insert");
                Label eIsTrue_Insert = (Label)fvServiceDataView.FindControl("eIsTrue_Insert");

                DateTime vDateResult;
                DateTime vToday = DateTime.Today;
                try
                {
                    string vSeriveNo_TypeA = (DateTime.TryParse(eBuildDate_Insert.Text.Trim(), out vDateResult)) ?
                                             vDateResult.Year.ToString("D4") + vDateResult.ToString("MMdd") + eServiceType_Insert.Text.Trim() :
                                             vToday.Year.ToString("D4") + vToday.ToString("MMdd") + eServiceType_Insert.Text.Trim();
                    vTempStr = "select RIGHT(MAX(ServiceNo), 3) [Index] from CustomService where ServiceNo like '" + vSeriveNo_TypeA + "%' ";
                    vIndex = PF.GetValue(vConnStrTemp, vTempStr, "Index");
                    vIndex = (vIndex != "") ? (int.Parse(vIndex) + 1).ToString("D3") : "001";
                    vServiceNo = vSeriveNo_TypeA + vIndex;

                    dsServiceDataDetail.InsertCommand = "INSERT INTO CustomService " + Environment.NewLine +
                                                        "       (ServiceNo, BuildDate, BuildTime, BuildMan, ServiceType, ServiceTypeB, ServiceTypeC, " + Environment.NewLine +
                                                        "        LinesNo, Car_ID, Driver, DriverName, AthorityDepNo, AthorityDepNote, " + Environment.NewLine +
                                                        "        LinesNo2, Car_ID2, Driver2, DriverName2, AthorityDepNo2, AthorityDepNote2, " + Environment.NewLine +
                                                        "        BoardTime, BoardStation, GetoffTime, GetoffStation, ServiceNote, " + Environment.NewLine +
                                                        "        IsNoContect, CivicName, CivicTelNo, CivicTelExtNo, CivicCellPhone, CivicAddress, CivicEMail, " + Environment.NewLine +
                                                        "        IsReplied, Remark, IsPending, AssignDate, AssignMan, CaseSource, ServiceDate, IsTrue) " + Environment.NewLine +
                                                        "VALUES (@ServiceNo, @BuildDate, @BuildTime, @BuildMan, @ServiceType, @ServiceTypeB, @ServiceTypeC, " + Environment.NewLine +
                                                        "        @LinesNo, @Car_ID, @Driver, @DriverName, @AthorityDepNo, @AthorityDepNote, " + Environment.NewLine +
                                                        "        @LinesNo2, @Car_ID2, @Driver2, @DriverName2, @AthorityDepNo2, @AthorityDepNote2, " + Environment.NewLine +
                                                        "        @BoardTime, @BoardStation, @GetoffTime, @GetoffStation, @ServiceNote, " + Environment.NewLine +
                                                        "        @IsNoContect, @CivicName, @CivicTelNo, @CivicTelExtNo, @CivicCellPhone, @CivicAddress, @CivicEMail, " + Environment.NewLine +
                                                        "        @IsReplied, @Remark, @IsPending, @AssignDate, @AssignMan, @CaseSource, @ServiceDate, @IsTrue)";
                    dsServiceDataDetail.InsertParameters.Clear();
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("ServiceNo", DbType.String, vServiceNo.Trim()));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("BuildDate", DbType.DateTime, (eBuildDate_Insert.Text.Trim() != "") ? eBuildDate_Insert.Text.Trim() : DateTime.Today.ToShortDateString()));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("BuildTime", DbType.String, (eBuildTime_Insert.Text.Trim() != "") ? eBuildTime_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("BuildMan", DbType.String, (eBuildMan_Insert.Text.Trim() != "") ? eBuildMan_Insert.Text.Trim() : vLoginID));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("ServiceType", DbType.String, (eServiceType_Insert.Text.Trim() != "") ? eServiceType_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("ServiceTypeB", DbType.String, (eServiceTypeB_Insert.Text.Trim() != "") ? eServiceTypeB_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("ServiceTypeC", DbType.String, (eServiceTypeC_Insert.Text.Trim() != "") ? eServiceTypeC_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("LinesNo", DbType.String, (eLinesNo_Insert.Text.Trim() != "") ? eLinesNo_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("Car_ID", DbType.String, (eCar_ID_Insert.Text.Trim() != "") ? eCar_ID_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("Driver", DbType.String, (eDriver_Insert.Text.Trim() != "") ? eDriver_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("DriverName", DbType.String, (eDriverName_Insert.Text.Trim() != "") ? eDriverName_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("AthorityDepNo", DbType.String, (eAthorityDepNo_Insert.Text.Trim() != "") ? eAthorityDepNo_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("AthorityDepNote", DbType.String, (eAthorityDepNote_Insert.Text.Trim() != "") ? eAthorityDepNote_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("LinesNo2", DbType.String, (eLinesNo2_Insert.Text.Trim() != "") ? eLinesNo2_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("Car_ID2", DbType.String, (eCar_ID2_Insert.Text.Trim() != "") ? eCar_ID2_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("Driver2", DbType.String, (eDriver2_Insert.Text.Trim() != "") ? eDriver2_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("DriverName2", DbType.String, (eDriverName2_Insert.Text.Trim() != "") ? eDriverName2_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("AthorityDepNo2", DbType.String, (eAthorityDepNo2_Insert.Text.Trim() != "") ? eAthorityDepNo2_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("AthorityDepNote2", DbType.String, (eAthorityDepNote2_Insert.Text.Trim() != "") ? eAthorityDepNote2_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("BoardTime", DbType.String, (eBoardTime_Insert.Text.Trim() != "") ? eBoardTime_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("BoardStation", DbType.String, (eBoardStation_Insert.Text.Trim() != "") ? eBoardStation_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("GetoffTime", DbType.String, (eGetoffTime_Insert.Text.Trim() != "") ? eGetoffTime_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("GetoffStation", DbType.String, (eGetoffStation_Insert.Text.Trim() != "") ? eGetoffStation_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("ServiceNote", DbType.String, (eServiceNote_Insert.Text.Trim() != "") ? eServiceNote_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("IsNoContect", DbType.String, (eIsNoContect_Insert.Checked == true) ? "true" : "false"));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("CivicName", DbType.String, (eCivicName_Insert.Text.Trim() != "") ? eCivicName_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("CivicTelNo", DbType.String, (eCivicTelNo_Insert.Text.Trim() != "") ? eCivicTelNo_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("CivicTelExtNo", DbType.String, (eCivicTelExtNo_Insert.Text.Trim() != "") ? eCivicTelExtNo_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("CivicCellPhone", DbType.String, (eCivicCellPhone_Insert.Text.Trim() != "") ? eCivicCellPhone_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("CivicAddress", DbType.String, (eCivicAddress_Insert.Text.Trim() != "") ? eCivicAddress_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("CivicEMail", DbType.String, (eCivicEMail_Insert.Text.Trim() != "") ? eCivicEMail_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("IsReplied", DbType.Boolean, (eIsReplied_Insert.Checked == true) ? "true" : "false"));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Insert.Text.Trim() != "") ? eRemark_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("IsPending", DbType.Boolean, (eIsPending_Insert.Checked == true) ? "true" : "false"));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("AssignDate", DbType.DateTime, (eAssignDate_Insert.Text.Trim() != "") ? eAssignDate_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan_Insert.Text.Trim() != "") ? eAssignMan_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("CaseSource", DbType.String, (eCaseSource_Insert.Text.Trim() != "") ? eCaseSource_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("ServiceDate", DbType.DateTime, (eServiceDate_Insert.Text.Trim() != "") ? eServiceDate_Insert.Text.Trim() : String.Empty));
                    dsServiceDataDetail.InsertParameters.Add(new Parameter("IsTrue", DbType.String, (eIsTrue_Insert.Text.Trim() != "") ? eIsTrue_Insert.Text.Trim() : String.Empty));

                    dsServiceDataDetail.Insert();
                    fvServiceDataView.ChangeMode(FormViewMode.ReadOnly);
                    fvServiceDataView.DataBind();
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

        protected void ddlIsTrue_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eIsTrue_Search.Text = ddlIsTrue_Search.SelectedValue.Trim();
        }
    }
}