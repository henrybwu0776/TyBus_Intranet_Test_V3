using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Threading;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class RescueCarAssignList : System.Web.UI.Page
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
                        //Session["RescueNo"] = "";
                        DateTime vCalDate = DateTime.Today.AddMonths(-1);
                        eListYear_Search.Text = (vCalDate.Year - 1911).ToString().Trim();
                        eListMonth_Search.Text = vCalDate.Month.ToString().Trim();
                        bbPrintCarAssign.Visible = false;
                        bbPrintMonthList.Visible = false;
                        bbPrintPersonList.Visible = false;
                        /*
                        if (ddlDepNo_Search.Items.Count > 0)
                        {
                            ddlDepNo_Search.SelectedIndex = 0;
                            eDepNo_Search.Text = ddlDepNo_Search.Items[0].Value.Trim();
                        } //*/

                        if (vLoginDepNo != "09")
                        {
                            string vSQLStr_Temp = "SELECT d.DEPNO + '-' + d.NAME AS ListTitle, d.DepNo " + Environment.NewLine +
                                                  "  FROM DEPARTMENT d " + Environment.NewLine +
                                                  " WHERE (ISNULL(InSHReport, 'X') = 'V') " + Environment.NewLine +
                                                  "   AND d.DepNo in (SELECT DepNo FROM EmployeeDepNo e where e.EmpNo = '" + vLoginID + "' and isnull(e.IsUsed, 0) = 1)";
                            sdsDepNo_Search.SelectCommand = vSQLStr_Temp;
                            sdsDepNo_Search.DataBind();
                        }
                        ddlDepNo_Search.SelectedIndex = 0;
                        if (ddlDepNo_Search.Items.Count > 0)
                        {
                            eDepNo_Search.Text = ddlDepNo_Search.Items[0].Value.Trim();
                        }
                        //===================================================================================================================================================
                        plShowData.Visible = true;
                        plPrint.Visible = false;
                    }
                    else
                    {
                        //OpenData(Session["RescueNo"].ToString());
                        OpenData();
                        bbPrintCarAssign.Visible = (gridRescueCarAssignList.Rows.Count > 0);
                        bbPrintMonthList.Visible = (gridRescueCarAssignList.Rows.Count > 0);
                        bbPrintPersonList.Visible = (gridRescueCarAssignList.Rows.Count > 0);
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
            if ((vLoginDepNo != "09") && (eDepNo_Search.Text.Trim() == ""))
            {
                eDepNo_Search.Text = (ddlDepNo_Search.Items.Count > 0) ? ddlDepNo_Search.Items[0].Value.Trim() : vLoginDepNo;
            }
            string vCaseYear = (eListYear_Search.Text.Trim() != "") ? (Int32.Parse(eListYear_Search.Text.Trim()) + 1911).ToString("D4") : DateTime.Today.Year.ToString("D4");
            string vCaseMonth = (eListMonth_Search.Text.Trim() != "") ? eListMonth_Search.Text.Trim() : DateTime.Today.Month.ToString("D2");
            string vCaseDate_Temp = vCaseYear + "/" + vCaseMonth + "/01";
            string vCaseDateS = PF.GetMonthFirstDay(DateTime.Parse(vCaseDate_Temp), "B");
            string vCaseDateE = PF.GetMonthLastDay(DateTime.Parse(vCaseDate_Temp), "B");
            string vWStr_ListDate = ((eListYear_Search.Text.Trim() != "") && (eListMonth_Search.Text.Trim() != "")) ? "   and r.CaseDate between '" + vCaseDateS.Trim() + "' and '" + vCaseDateE.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and r.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_AssignMan = (eAssignMan_Search.Text.Trim() != "") ?
                                     "   and (r.AssignMan = '" + eAssignMan_Search.Text.Trim() + "' " + Environment.NewLine +
                                     "        or r.FirstSupportMan = '" + eAssignMan_Search.Text.Trim() + "' " + Environment.NewLine +
                                     "        or r.SecondSupportMan = '" + eAssignMan_Search.Text.Trim() + "' )" + Environment.NewLine :
                                     "";
            //2020.08.27 增加起迄日期
            string vReturnStr = "select ListNo, CaseDate, Car_ID, StartPosition, StartDate, StartTime, StartMeterMiles, " + Environment.NewLine +
                                "       FinishPosition, FinishDate, FinishTime, FinishMeterMiles, " + Environment.NewLine +
                                "       AssignMan, (select [Name] from Employee where EmpNo = r.AssignMan) AssignManName, " + Environment.NewLine +
                                "       DepNo, (select [Name] from Department where DepNo = r.DepNo) DepName, " + Environment.NewLine +
                                "       FirstSupportMan, (select [Name] from Employee where EmpNo = r.FirstSupportMan) FirstSupportManName, " + Environment.NewLine +
                                "       SecondSupportMan, (select [Name] from Employee where EmpNo = r.SecondSupportMan) SecondSupportManName, " + Environment.NewLine +
                                "       AssignReason, ReasonNote, UsedMiles, UsedMin, Remark, " + Environment.NewLine +
                                "       BuMan, (select [Name] from Employee where EmpNo = r.BuMan) BuManName, BuDate, " + Environment.NewLine +
                                "       ModifyMan, (select [Name] from Employee where EmpNo = r.ModifyMan) ModifyManName, ModifyDate, TargetCarID " + Environment.NewLine +
                                "  from RescueCarAssignList r " + Environment.NewLine +
                                " where 1 = 1 " + Environment.NewLine +
                                vWStr_AssignMan + vWStr_DepNo + vWStr_ListDate +
                                " order by CaseDate, StartTime ";
            return vReturnStr;
        }

        private string GetSelectStrByRescueNo(string fRescueNo)
        {
            string vReturnStr = "select ListNo, CaseDate, Car_ID, StartPosition, StartDate, StartTime, StartMeterMiles, " + Environment.NewLine +
                                "       FinishPosition, FinishDate, FinishTime, FinishMeterMiles, " + Environment.NewLine +
                                "       AssignMan, (select [Name] from Employee where EmpNo = r.AssignMan) AssignManName, " + Environment.NewLine +
                                "       DepNo, (select [Name] from Department where DepNo = r.DepNo) DepName, " + Environment.NewLine +
                                "       FirstSupportMan, (select [Name] from Employee where EmpNo = r.FirstSupportMan) FirstSupportManName, " + Environment.NewLine +
                                "       SecondSupportMan, (select [Name] from Employee where EmpNo = r.SecondSupportMan) SecondSupportManName, " + Environment.NewLine +
                                "       AssignReason, ReasonNote, UsedMiles, UsedMin, Remark, " + Environment.NewLine +
                                "       BuMan, (select [Name] from Employee where EmpNo = r.BuMan) BuManName, BuDate, " + Environment.NewLine +
                                "       ModifyMan, (select [Name] from Employee where EmpNo = r.ModifyMan) ModifyManName, ModifyDate, TargetCarID " + Environment.NewLine +
                                "  from RescueCarAssignList r " + Environment.NewLine +
                                " where ListNo = '" + fRescueNo + "' ";
            return vReturnStr;
        }

        private void OpenData(string fListNo)
        {
            string vSelStr = (fListNo == "") ? GetSelectStr() : GetSelectStrByRescueNo(fListNo);
            sdsRescueCarAssignList.SelectCommand = vSelStr;
            gridRescueCarAssignList.DataBind();
        }

        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            sdsRescueCarAssignList.SelectCommand = vSelStr;
            gridRescueCarAssignList.DataBind();
        }

        /// <summary>
        /// 取回人員公出記錄單用的語法
        /// </summary>
        /// <returns></returns>
        private string GetPersonListStr()
        {
            string vCaseYear = (eListYear_Search.Text.Trim() != "") ? (Int32.Parse(eListYear_Search.Text.Trim()) + 1911).ToString("D4") : DateTime.Today.Year.ToString("D4");
            string vCaseMonth = (eListMonth_Search.Text.Trim() != "") ? eListMonth_Search.Text.Trim() : DateTime.Today.Month.ToString("D2");
            string vCaseDate_Temp = vCaseYear + "/" + vCaseMonth + "/01";
            string vCaseDateS = PF.GetMonthFirstDay(DateTime.Parse(vCaseDate_Temp), "B");
            string vCaseDateE = PF.GetMonthLastDay(DateTime.Parse(vCaseDate_Temp), "B");
            string vWStr_ListDate = ((eListYear_Search.Text.Trim() != "") && (eListMonth_Search.Text.Trim() != "")) ? "   and r.CaseDate between '" + vCaseDateS.Trim() + "' and '" + vCaseDateE.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and r.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_AssignMan = (eAssignMan_Search.Text.Trim() != "") ?
                                     "   and (r.AssignMan = '" + eAssignMan_Search.Text.Trim() + "' " + Environment.NewLine +
                                     "        or r.FirstSupportMan = '" + eAssignMan_Search.Text.Trim() + "' " + Environment.NewLine +
                                     "        or r.SecondSupportMan = '" + eAssignMan_Search.Text.Trim() + "' )" + Environment.NewLine :
                                     "";
            string vReturnStr = "select DepNo, CaseDate, case when AssignReason <> '其他' then AssignReason else ReasonNote end Reason, Remark, " + Environment.NewLine +
                                //2020.08.27 增加起迄日期
                                "       FinishPosition, StartDate, StartTime, FinishDate, FinishTime, (select[Name] from Employee where EmpNo = r.AssignMan) UseMan, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '員工資料        EMPLOYEE        TITLE' and ClassNo = (select Title from Employee where EmpNo = r.AssignMan)) TitleName " + Environment.NewLine +
                                "  from RescueCarAssignList r " + Environment.NewLine +
                                " where isnull(r.AssignMan, '') <> '' " + Environment.NewLine +
                                vWStr_AssignMan + vWStr_DepNo + vWStr_ListDate +
                                " union all " + Environment.NewLine +
                                "select DepNo, CaseDate, case when AssignReason <> '其他' then AssignReason else ReasonNote end Reason, Remark, " + Environment.NewLine +
                                "       FinishPosition, StartDate, StartTime, FinishDate, FinishTime, (select[Name] from Employee where EmpNo = r.FirstSupportMan) UseMan, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '員工資料        EMPLOYEE        TITLE' and ClassNo = (select Title from Employee where EmpNo = r.FirstSupportMan)) TitleName " + Environment.NewLine +
                                "  from RescueCarAssignList r " + Environment.NewLine +
                                " where isnull(r.FirstSupportMan,'') <> '' " + Environment.NewLine +
                                vWStr_AssignMan + vWStr_DepNo + vWStr_ListDate +
                                " union all " + Environment.NewLine +
                                "select DepNo, CaseDate, case when AssignReason<> '其他' then AssignReason else ReasonNote end Reason, Remark, " + Environment.NewLine +
                                "       FinishPosition, StartDate, StartTime, FinishDate, FinishTime, (select[Name] from Employee where EmpNo = r.SecondSupportMan) UseMan, " + Environment.NewLine +
                                "       (select ClassTxt from DBDICB where FKey = '員工資料        EMPLOYEE        TITLE' and ClassNo = (select Title from Employee where EmpNo = r.SecondSupportMan)) TitleName " + Environment.NewLine +
                                "  from RescueCarAssignList r " + Environment.NewLine +
                                " where isnull(r.SecondSupportMan,'') <> '' " + Environment.NewLine +
                                vWStr_AssignMan + vWStr_DepNo + vWStr_ListDate;
            return vReturnStr;
        }

        protected void ddlDepNo_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDepNo_Search.Text = ddlDepNo_Search.SelectedValue.Trim();
        }

        protected void eAssignMan_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vAssignMan_Temp = eAssignMan_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from EMployee where EmpNo = '" + vAssignMan_Temp + "' ";
            string vAssignManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vAssignManName_Temp == "")
            {
                vAssignManName_Temp = vAssignMan_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vAssignManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vAssignMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eAssignMan_Search.Text = vAssignMan_Temp.Trim();
            eAssignManNAme_Search.Text = vAssignManName_Temp.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            //Session["RescueNo"] = "";
            //OpenData(Session["RescueNo"].ToString());
            OpenData();
        }

        /// <summary>
        /// 預覽行車憑單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintMonthList_Click(object sender, EventArgs e)
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
                DataTable dtPrintPoint = new DataTable("RescueCarAssignListP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("RescueCarAssignListP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\RescueCarMonthListP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "公務車行車憑單"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;

                    string vListYMStr = eListYear_Search.Text.Trim() + " 年 " + eListMonth_Search.Text.Trim() + " 月";
                    string vDepNoStr = ddlDepNo_Search.SelectedItem.Text.Trim();
                    string vAssignManStr = (eAssignMan_Search.Text.Trim() != "") ? eAssignMan_Search.Text.Trim() : "全部";
                    string vRecordNote = "預覽報表_救援車勤務派遣登記表_公務車行車憑單" + Environment.NewLine +
                                         "RescueCarAssignList.aspx" + Environment.NewLine +
                                         "派車年月：" + vListYMStr + Environment.NewLine +
                                         "派車單位：" + vDepNoStr + Environment.NewLine +
                                         "用車人：" + vAssignManStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        /// <summary>
        /// 預覽人員公出記錄
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintPersonList_Click(object sender, EventArgs e)
        {
            string vSelStr = GetPersonListStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintPoint = new SqlDataAdapter(vSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintPoint = new DataTable("RescueCarPersonListP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("RescueCarPersonListP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\RescueCarPersonListP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "人員公出登記簿"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;

                    string vListYMStr = eListYear_Search.Text.Trim() + " 年 " + eListMonth_Search.Text.Trim() + " 月";
                    string vDepNoStr = ddlDepNo_Search.SelectedItem.Text.Trim();
                    string vAssignManStr = (eAssignMan_Search.Text.Trim() != "") ? eAssignMan_Search.Text.Trim() : "全部";
                    string vRecordNote = "預覽報表_救援車勤務派遣登記表_人員公出登記簿" + Environment.NewLine +
                                         "RescueCarAssignList.aspx" + Environment.NewLine +
                                         "派車年月：" + vListYMStr + Environment.NewLine +
                                         "派車單位：" + vDepNoStr + Environment.NewLine +
                                         "用車人：" + vAssignManStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        /// <summary>
        /// 預覽派車單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintCarAssign_Click(object sender, EventArgs e)
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
                DataTable dtPrintPoint = new DataTable("RescueCarAssignListP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("RescueCarAssignListP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\RescueCarAssignSheetP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "桃園汽車客運公司救援車派車單"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;

                    string vListYMStr = eListYear_Search.Text.Trim() + " 年 " + eListMonth_Search.Text.Trim() + " 月";
                    string vDepNoStr = ddlDepNo_Search.SelectedItem.Text.Trim();
                    string vAssignManStr = (eAssignMan_Search.Text.Trim() != "") ? eAssignMan_Search.Text.Trim() : "全部";
                    string vRecordNote = "預覽報表_救援車勤務派遣登記表_桃園汽車客運公司救援車派車單" + Environment.NewLine +
                                         "RescueCarAssignList.aspx" + Environment.NewLine +
                                         "派車年月：" + vListYMStr + Environment.NewLine +
                                         "派車單位：" + vDepNoStr + Environment.NewLine +
                                         "用車人：" + vAssignManStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvRescueCarAssignDetail_DataBound(object sender, EventArgs e)
        {
            string vDateURL_Temp = "";
            string vDateScript_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvRescueCarAssignDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    //Session["RescueNo"] = " ";
                    TextBox eCaseDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eCaseDate_Edit");
                    if (eCaseDate_Edit != null)
                    {
                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Edit.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eCaseDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                        //2020.08.27 增加起迄日期
                        TextBox eStartDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartDate_Edit");
                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eStartDate_Edit.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eStartDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                        TextBox eFinishDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishDate_Edit");
                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eFinishDate_Edit.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eFinishDate_Edit.Attributes["onClick"] = vDateScript_Temp;


                        Label eModifyMan_Edit = (Label)fvRescueCarAssignDetail.FindControl("eModifyMan_Edit");
                        eModifyMan_Edit.Text = vLoginID;
                        Label eModifyManName_Edit = (Label)fvRescueCarAssignDetail.FindControl("eModifyManName_Edit");
                        eModifyManName_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                        Label eModifyDate_Edit = (Label)fvRescueCarAssignDetail.FindControl("eModifyDate_Edit");
                        eModifyDate_Edit.Text = PF.GetChinsesDate(DateTime.Today);
                        RadioButtonList rbAssignReason_Edit = (RadioButtonList)fvRescueCarAssignDetail.FindControl("rbAssignReason_Edit");
                        Label eAssignReason_Edit = (Label)fvRescueCarAssignDetail.FindControl("eAssignReason_Edit");
                        rbAssignReason_Edit.SelectedIndex = (eAssignReason_Edit.Text.Trim() != "") ? rbAssignReason_Edit.Items.IndexOf(new ListItem(eAssignReason_Edit.Text.Trim())) : 0;
                    }
                    break;
                case FormViewMode.Insert:
                    //Session["RescueNo"] = " ";
                    TextBox eCaseDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eCaseDate_INS");
                    if (eCaseDate_INS != null)
                    {
                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_INS.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eCaseDate_INS.Attributes["onClick"] = vDateScript_Temp;
                        eCaseDate_INS.Text = PF.GetChinsesDate(DateTime.Today);

                        //2020.08.27 增加起迄日期
                        TextBox eStartDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eStartDate_INS");
                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eStartDate_INS.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eStartDate_INS.Attributes["onClick"] = vDateScript_Temp;
                        eStartDate_INS.Text = PF.GetChinsesDate(DateTime.Today);

                        TextBox eFinishDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishDate_INS");
                        vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eFinishDate_INS.ClientID;
                        vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eFinishDate_INS.Attributes["onClick"] = vDateScript_Temp;
                        eFinishDate_INS.Text = PF.GetChinsesDate(DateTime.Today);

                        Label eBuMan_INS = (Label)fvRescueCarAssignDetail.FindControl("eBuMan_INS");
                        eBuMan_INS.Text = vLoginID;
                        Label eBuManName_INS = (Label)fvRescueCarAssignDetail.FindControl("eBuManName_INS");
                        eBuManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");

                        Label eBuDate_INS = (Label)fvRescueCarAssignDetail.FindControl("eBuDate_INS");
                        eBuDate_INS.Text = PF.GetChinsesDate(DateTime.Today);

                        /* 2020.03.19 先不預帶單位
                        TextBox eDepNo_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eDepNo_INS");
                        Label eDepName_INS = (Label)fvRescueCarAssignDetail.FindControl("eDepName_INS");
                        eDepNo_INS.Text = (eDepNo_Search.Text.Trim() != "") ? eDepNo_Search.Text.Trim() : PF.GetValue(vConnStr, "select DepNo from Employee where EmpNo = '" + eBuMan_INS.Text.Trim() + "' ", "DepNo");
                        eDepName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Department where DepNo = '" + eDepNo_INS.Text.Trim() + "' ", "Name");
                        //*/

                        RadioButtonList rbAssignReason_INS = (RadioButtonList)fvRescueCarAssignDetail.FindControl("rbAssignReason_INS");
                        rbAssignReason_INS.SelectedIndex = 0;
                        TextBox eAssignReason_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eAssignReason_INS");
                        eAssignReason_INS.Text = rbAssignReason_INS.SelectedValue;
                    }
                    break;
            }
        }

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eListNo_Edit = (Label)fvRescueCarAssignDetail.FindControl("eListNo_Edit");
            string vListNo = (eListNo_Edit != null) ? eListNo_Edit.Text.Trim() : "";
            if (vListNo != "")
            {
                string vSQLStr_Temp = "select max(Items) MaxItems from RescueCarAssignList_History where ListNo = '" + vListNo.Trim() + "' ";
                string vMaxItemsStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems");
                int vMaxItems = (vMaxItemsStr != "") ? Int32.Parse(vMaxItemsStr.Trim()) + 1 : 1;
                string vListNoItem = vListNo.Trim() + vMaxItems.ToString("D4");
                string vExecSQL_Temp = "insert into RescueCarAssignList_History " + Environment.NewLine +
                                       "       (ListNoItem, Items, ModifyMode, ListNo, CaseDate, Car_ID, StartPosition, StartDate, StartTime, StartMeterMiles, " + Environment.NewLine +
                                       "        FinishPosition, FinishDate, FinishTime, FinishMeterMiles, AssignMan, DepNo, FirstSupportMan, SecondSupportMan, " + Environment.NewLine +
                                       "        AssignReason, ReasonNote, UsedMiles, UsedMin, Remark, BuMan, BuDate, ModifyMan, ModifyDate, TargetCarID) " + Environment.NewLine +
                                       "select '" + vListNoItem + "', '" + vMaxItems.ToString("D4") + "', 'EDIT',ListNo, CaseDate, Car_ID, " + Environment.NewLine +
                                       "       StartPosition, StartDate, StartTime, StartMeterMiles, FinishPosition, FinishDate, FinishTime, FinishMeterMiles, " + Environment.NewLine +
                                       "       AssignMan, DepNo, FirstSupportMan, SecondSupportMan, AssignReason, ReasonNote, UsedMiles, UsedMin, " + Environment.NewLine +
                                       "       Remark, BuMan, BuDate, ModifyMan, ModifyDate, TargetCarID " + Environment.NewLine +
                                       "  from RescueCarAssignList " + Environment.NewLine +
                                       " where ListNo = '" + vListNo + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vExecSQL_Temp);

                    sdsRescueCarAssignDetail.UpdateParameters.Clear();

                    TextBox eCaseDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eCaseDate_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("CaseDate", DbType.DateTime, (eCaseDate_Edit.Text.Trim() != "") ? eCaseDate_Edit.Text.Trim() : String.Empty));

                    TextBox eCar_ID_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eCar_ID_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("Car_ID", DbType.String, (eCar_ID_Edit.Text.Trim() != "") ? eCar_ID_Edit.Text.Trim() : String.Empty));

                    TextBox eStartPosition_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartPosition_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("StartPosition", DbType.String, (eStartPosition_Edit.Text.Trim() != "") ? eStartPosition_Edit.Text.Trim() : String.Empty));

                    TextBox eStartDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartDate_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("StartDate", DbType.Date, (eStartDate_Edit.Text.Trim() != "") ? eStartDate_Edit.Text.Trim() : String.Empty));

                    TextBox eStartTime_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartTime_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("StartTime", DbType.String, (eStartTime_Edit.Text.Trim() != "") ? eStartTime_Edit.Text.Trim() : String.Empty));

                    TextBox eStartMeterMiles_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartMeterMiles_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("StartMeterMiles", DbType.String, (eStartMeterMiles_Edit.Text.Trim() != "") ? eStartMeterMiles_Edit.Text.Trim() : String.Empty));

                    TextBox eFinishPosition_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishPosition_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("FinishPosition", DbType.String, (eFinishPosition_Edit.Text.Trim() != "") ? eFinishPosition_Edit.Text.Trim() : String.Empty));

                    TextBox eFinishDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishDate_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("FinishDate", DbType.Date, (eFinishDate_Edit.Text.Trim() != "") ? eFinishDate_Edit.Text.Trim() : String.Empty));

                    TextBox eFinishTime_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishTime_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("FinishTime", DbType.String, (eFinishTime_Edit.Text.Trim() != "") ? eFinishTime_Edit.Text.Trim() : String.Empty));

                    TextBox eFinishMeterMiles_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishMeterMiles_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("FinishMeterMiles", DbType.String, (eFinishMeterMiles_Edit.Text.Trim() != "") ? eFinishMeterMiles_Edit.Text.Trim() : String.Empty));

                    TextBox eAssignMan_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eAssignMan_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan_Edit.Text.Trim() != "") ? eAssignMan_Edit.Text.Trim() : String.Empty));

                    TextBox eDepNo_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eDepNo_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_Edit.Text.Trim() != "") ? eDepNo_Edit.Text.Trim() : String.Empty));

                    TextBox eFirstSupportMan_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFirstSupportMan_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("FirstSupportMan", DbType.String, (eFirstSupportMan_Edit.Text.Trim() != "") ? eFirstSupportMan_Edit.Text.Trim() : String.Empty));

                    TextBox eSecondSupportMan_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eSecondSupportMan_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("SecondSupportMan", DbType.String, (eSecondSupportMan_Edit.Text.Trim() != "") ? eSecondSupportMan_Edit.Text.Trim() : String.Empty));

                    Label eAssignReason_Edit = (Label)fvRescueCarAssignDetail.FindControl("eAssignReason_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("AssignReason", DbType.String, (eAssignReason_Edit.Text.Trim() != "") ? eAssignReason_Edit.Text.Trim() : String.Empty));

                    TextBox eReasonNote_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eReasonNote_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("ReasonNote", DbType.String, (eReasonNote_Edit.Text.Trim() != "") ? eReasonNote_Edit.Text.Trim() : String.Empty));

                    Label eUsedMiles_Edit = (Label)fvRescueCarAssignDetail.FindControl("eUsedMiles_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("UsedMiles", DbType.Int32, (eUsedMiles_Edit.Text.Trim() != "") ? eUsedMiles_Edit.Text.Trim() : String.Empty));

                    Label eUsedMin_Edit = (Label)fvRescueCarAssignDetail.FindControl("eUsedMin_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("UsedMin", DbType.Int32, (eUsedMin_Edit.Text.Trim() != "") ? eUsedMin_Edit.Text.Trim() : String.Empty));

                    TextBox eRemark_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eRemark_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));

                    Label eModifyMan_Edit = (Label)fvRescueCarAssignDetail.FindControl("eModifyMan_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, (eModifyMan_Edit.Text.Trim() != "") ? eModifyMan_Edit.Text.Trim() : String.Empty));

                    Label eModifyDate_Edit = (Label)fvRescueCarAssignDetail.FindControl("eModifyDate_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.DateTime, (eModifyDate_Edit.Text.Trim() != "") ? eModifyDate_Edit.Text.Trim() : String.Empty));

                    TextBox eTargetCarID_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eTargetCarID_Edit");
                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("TargetCarID", DbType.String, (eTargetCarID_Edit.Text.Trim() != "") ? eTargetCarID_Edit.Text.Trim() : String.Empty));

                    sdsRescueCarAssignDetail.UpdateParameters.Add(new Parameter("ListNo", DbType.String, vListNo.Trim()));
                    sdsRescueCarAssignDetail.Update();
                    fvRescueCarAssignDetail.ChangeMode(FormViewMode.ReadOnly);
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
                Response.Write("alert('請先選擇要修改的派車單！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void eCar_ID_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eCar_ID_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eCar_ID_Edit");
            TextBox eStartPosition_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartPosition_Edit");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "select (select [Name] from Department where DepNo = c.CompanyNo) DepName from Car_InfoA c where Car_ID = '" + eCar_ID_Edit.Text.Trim() + "' ";
            eStartPosition_Edit.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "DepName").Trim();
        }

        protected void eStartTime_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eStartDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartDate_Edit");
            TextBox eStartTime_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartTime_Edit");
            TextBox eFinishDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishDate_Edit");
            TextBox eFinishTime_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishTime_Edit");
            Label eUsedMin_Edit = (Label)fvRescueCarAssignDetail.FindControl("eUsedMin_Edit");
            int vDateDiff = 0;
            if (eStartDate_Edit != null)
            {
                DateTime vStartDate_Edit = (eStartDate_Edit.Text.Trim() != "") ? DateTime.Parse(eStartDate_Edit.Text.Trim()) : DateTime.Today;
                DateTime vFinishDate_Edit = (eFinishDate_Edit.Text.Trim() != "") ? DateTime.Parse(eFinishDate_Edit.Text.Trim()) : DateTime.Today;
                TimeSpan vSpanDate = vFinishDate_Edit.Subtract(vStartDate_Edit);
                vDateDiff = vSpanDate.Days;
                if ((eStartTime_Edit.Text.Trim() != "") && (eFinishTime_Edit.Text.Trim() != ""))
                {
                    string vStartTime_Temp = eStartTime_Edit.Text.Trim();
                    string vFinishTime_Temp = eFinishTime_Edit.Text.Trim();
                    if ((vStartTime_Temp.IndexOf(':') < 0) || (vFinishTime_Temp.IndexOf(':') < 0))
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('時間格式輸入錯誤，請重新輸入！')");
                        Response.Write("</" + "Script>");
                        eStartTime_Edit.Focus();
                    }
                    else
                    {
                        int vStartMin = PF.TimeStrToMins(vStartTime_Temp);
                        int vFinishMin = PF.TimeStrToMins(vFinishTime_Temp);
                        eUsedMin_Edit.Text = (vFinishMin - vStartMin + (vDateDiff * 24 * 60)).ToString();
                    }
                }
            }
        }

        protected void eStartMeterMiles_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eStartMeterMiles_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartMeterMiles_Edit");
            TextBox eFinishMeterMiles_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishMeterMiles_Edit");
            Label eUsedMiles_Edit = (Label)fvRescueCarAssignDetail.FindControl("eUsedMiles_Edit");
            if (eStartMeterMiles_Edit != null)
            {
                if ((eStartMeterMiles_Edit.Text.Trim() != "") && (eFinishMeterMiles_Edit.Text.Trim() != ""))
                {
                    int vStartMeterMiles_Temp = Int32.Parse(eStartMeterMiles_Edit.Text.Trim());
                    int vFinishMeterMiles_Temp = Int32.Parse(eFinishMeterMiles_Edit.Text.Trim());
                    eUsedMiles_Edit.Text = (vFinishMeterMiles_Temp - vStartMeterMiles_Temp).ToString();
                }
            }
        }

        protected void eAssignMan_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eAssignMan_Temp = (TextBox)fvRescueCarAssignDetail.FindControl("eAssignMan_Edit");
            Label eAssignManName_Temp = (Label)fvRescueCarAssignDetail.FindControl("eAssignManName_Edit");
            if (eAssignMan_Temp != null)
            {
                string vAssignMan_Temp = eAssignMan_Temp.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan_Temp.Trim() + "' ";
                string vAssignManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vAssignManName_Temp == "")
                {
                    vAssignManName_Temp = vAssignMan_Temp.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vAssignManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                    vAssignMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eAssignMan_Temp.Text = vAssignMan_Temp.Trim();
                eAssignManName_Temp.Text = vAssignManName_Temp.Trim();
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Temp = (TextBox)fvRescueCarAssignDetail.FindControl("eDepNo_Edit");
            Label eDepName_Temp = (Label)fvRescueCarAssignDetail.FindControl("eDepName_Edit");
            if (eDepNo_Temp != null)
            {
                string vDepNo_Temp = eDepNo_Temp.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
                string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName_Temp == "")
                {
                    vDepName_Temp = vDepNo_Temp.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                    vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_Temp.Text = vDepNo_Temp.Trim();
                eDepName_Temp.Text = vDepName_Temp.Trim();
            }
        }

        protected void eFirstSupportMan_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eFirstSupportMan_Temp = (TextBox)fvRescueCarAssignDetail.FindControl("eFirstSupportMan_Edit");
            Label eFirstSupportManName_Temp = (Label)fvRescueCarAssignDetail.FindControl("eFirstSupportManName_Edit");
            if (eFirstSupportMan_Temp != null)
            {
                string vFirstSupportMan_Temp = eFirstSupportMan_Temp.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vFirstSupportMan_Temp.Trim() + "' ";
                string vFirstSupportManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vFirstSupportManName_Temp == "")
                {
                    vFirstSupportManName_Temp = vFirstSupportMan_Temp.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vFirstSupportManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                    vFirstSupportMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eFirstSupportMan_Temp.Text = vFirstSupportMan_Temp.Trim();
                eFirstSupportManName_Temp.Text = vFirstSupportManName_Temp.Trim();
            }
        }

        protected void eSecondSupportMan_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eSecondSupportMan_Temp = (TextBox)fvRescueCarAssignDetail.FindControl("eSecondSupportMan_Edit");
            Label eSecondSupportManName_Temp = (Label)fvRescueCarAssignDetail.FindControl("eSecondSupportManName_Edit");
            if (eSecondSupportMan_Temp != null)
            {
                string vSecondSupportMan_Temp = eSecondSupportMan_Temp.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vSecondSupportMan_Temp.Trim() + "' ";
                string vSecondSupportManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vSecondSupportManName_Temp == "")
                {
                    vSecondSupportManName_Temp = vSecondSupportMan_Temp.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vSecondSupportManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                    vSecondSupportMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eSecondSupportMan_Temp.Text = vSecondSupportMan_Temp.Trim();
                eSecondSupportManName_Temp.Text = vSecondSupportManName_Temp.Trim();
            }
        }

        protected void rbAssignReason_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            RadioButtonList rbAssignReason_Edit = (RadioButtonList)fvRescueCarAssignDetail.FindControl("rbAssignReason_Edit");
            Label eAssignReason_Edit = (Label)fvRescueCarAssignDetail.FindControl("eAssignReason_Edit");
            eAssignReason_Edit.Text = rbAssignReason_Edit.SelectedValue;
        }

        protected void eCar_ID_INS_TextChanged(object sender, EventArgs e)
        {
            //2020.03.25 修車廠主任提出要求不要預帶出發地點
            /*
            TextBox eCar_ID_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eCar_ID_INS");
            TextBox eStartPosition_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eStartPosition_INS");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "select (select [Name] from Department where DepNo = c.CompanyNo) DepName from Car_InfoA c where Car_ID = '" + eCar_ID_INS.Text.Trim() + "' ";
            eStartPosition_INS.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "DepName").Trim();
            */
        }

        protected void eStartTime_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eStartDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eStartDate_INS");
            TextBox eStartTime_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eStartTime_INS");
            TextBox eFinishDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishDate_INS");
            TextBox eFinishTime_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishTime_INS");
            Label eUsedMin_INS = (Label)fvRescueCarAssignDetail.FindControl("eUsedMin_INS");
            int vDateDiff = 0;
            if (eStartDate_INS != null)
            {
                DateTime vStartDate_INS = (eStartDate_INS.Text.Trim() != "") ? DateTime.Parse(eStartDate_INS.Text.Trim()) : DateTime.Today;
                DateTime vFinishDate_INS = (eFinishDate_INS.Text.Trim() != "") ? DateTime.Parse(eFinishDate_INS.Text.Trim()) : DateTime.Today;
                TimeSpan vSpanDate = vFinishDate_INS.Subtract(vStartDate_INS);
                vDateDiff = vSpanDate.Days;
                if ((eStartTime_INS.Text.Trim() != "") && (eFinishTime_INS.Text.Trim() != ""))
                {
                    string vStartTime_Temp = eStartTime_INS.Text.Trim();
                    string vFinishTime_Temp = eFinishTime_INS.Text.Trim();
                    if ((vStartTime_Temp.IndexOf(':') < 0) || (vFinishTime_Temp.IndexOf(':') < 0))
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('時間格式輸入錯誤，請重新輸入！')");
                        Response.Write("</" + "Script>");
                        eStartTime_INS.Focus();
                    }
                    else
                    {
                        int vStartMin = PF.TimeStrToMins(vStartTime_Temp);
                        int vFinishMin = PF.TimeStrToMins(vFinishTime_Temp);
                        eUsedMin_INS.Text = (vFinishMin - vStartMin + (vDateDiff * 24 * 60)).ToString();
                    }
                }
            }
        }

        protected void eStartMeterMiles_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eStartMeterMiles_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eStartMeterMiles_INS");
            TextBox eFinishMeterMiles_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishMeterMiles_INS");
            Label eUsedMiles_INS = (Label)fvRescueCarAssignDetail.FindControl("eUsedMiles_INS");
            if (eStartMeterMiles_INS != null)
            {
                if ((eStartMeterMiles_INS.Text.Trim() != "") && (eFinishMeterMiles_INS.Text.Trim() != ""))
                {
                    int vStartMeterMiles_Temp = Int32.Parse(eStartMeterMiles_INS.Text.Trim());
                    int vFinishMeterMiles_Temp = Int32.Parse(eFinishMeterMiles_INS.Text.Trim());
                    eUsedMiles_INS.Text = (vFinishMeterMiles_Temp - vStartMeterMiles_Temp).ToString();
                }
            }
        }

        protected void eAssignMan_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eAssignMan_Temp = (TextBox)fvRescueCarAssignDetail.FindControl("eAssignMan_INS");
            Label eAssignManName_Temp = (Label)fvRescueCarAssignDetail.FindControl("eAssignManName_INS");
            if (eAssignMan_Temp != null)
            {
                string vAssignMan_Temp = eAssignMan_Temp.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan_Temp.Trim() + "' ";
                string vAssignManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vAssignManName_Temp == "")
                {
                    vAssignManName_Temp = vAssignMan_Temp.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vAssignManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                    vAssignMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eAssignMan_Temp.Text = vAssignMan_Temp.Trim();
                eAssignManName_Temp.Text = vAssignManName_Temp.Trim();
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Temp = (TextBox)fvRescueCarAssignDetail.FindControl("eDepNo_INS");
            Label eDepName_Temp = (Label)fvRescueCarAssignDetail.FindControl("eDepName_INS");
            if (eDepNo_Temp != null)
            {
                string vDepNo_Temp = eDepNo_Temp.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
                string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName_Temp == "")
                {
                    vDepName_Temp = vDepNo_Temp.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                    vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_Temp.Text = vDepNo_Temp.Trim();
                eDepName_Temp.Text = vDepName_Temp.Trim();
            }
        }

        protected void eFirstSupportMan_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eFirstSupportMan_Temp = (TextBox)fvRescueCarAssignDetail.FindControl("eFirstSupportMan_INS");
            Label eFirstSupportManName_Temp = (Label)fvRescueCarAssignDetail.FindControl("eFirstSupportManName_INS");
            if (eFirstSupportMan_Temp != null)
            {
                string vFirstSupportMan_Temp = eFirstSupportMan_Temp.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vFirstSupportMan_Temp.Trim() + "' ";
                string vFirstSupportManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vFirstSupportManName_Temp == "")
                {
                    vFirstSupportManName_Temp = vFirstSupportMan_Temp.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vFirstSupportManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                    vFirstSupportMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eFirstSupportMan_Temp.Text = vFirstSupportMan_Temp.Trim();
                eFirstSupportManName_Temp.Text = vFirstSupportManName_Temp.Trim();
            }
        }

        protected void eSecondSupportMan_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eSecondSupportMan_Temp = (TextBox)fvRescueCarAssignDetail.FindControl("eSecondSupportMan_INS");
            Label eSecondSupportManName_Temp = (Label)fvRescueCarAssignDetail.FindControl("eSecondSupportManName_INS");
            if (eSecondSupportMan_Temp != null)
            {
                string vSecondSupportMan_Temp = eSecondSupportMan_Temp.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vSecondSupportMan_Temp.Trim() + "' ";
                string vSecondSupportManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vSecondSupportManName_Temp == "")
                {
                    vSecondSupportManName_Temp = vSecondSupportMan_Temp.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vSecondSupportManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                    vSecondSupportMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eSecondSupportMan_Temp.Text = vSecondSupportMan_Temp.Trim();
                eSecondSupportManName_Temp.Text = vSecondSupportManName_Temp.Trim();
            }
        }

        protected void rbAssignReason_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            RadioButtonList rbAssignReason_INS = (RadioButtonList)fvRescueCarAssignDetail.FindControl("rbAssignReason_INS");
            TextBox eAssignReason_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eAssignReason_INS");
            eAssignReason_INS.Text = rbAssignReason_INS.SelectedValue;
        }

        /// <summary>
        /// 刪除資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eListNo_List = (Label)fvRescueCarAssignDetail.FindControl("eListNo_List");
            string vListNo = (eListNo_List != null) ? eListNo_List.Text.Trim() : "";
            if (vListNo != "")
            {
                string vSQLStr_Temp = "select max(Items) MaxItems from RescueCarAssignList_History where ListNo = '" + vListNo.Trim() + "' ";
                string vMaxItemsStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems");
                int vMaxItems = (vMaxItemsStr != "") ? Int32.Parse(vMaxItemsStr.Trim()) + 1 : 1;
                string vListNoItem = vListNo.Trim() + vMaxItems.ToString("D4");
                string vExecSQL_Temp = "insert into RescueCarAssignList_History " + Environment.NewLine +
                                       "       (ListNoItem, Items, ModifyMode, ListNo, CaseDate, Car_ID, StartPosition, StartDate, StartTime, StartMeterMiles, " + Environment.NewLine +
                                       "        FinishPosition, FinishDate, FinishTime, FinishMeterMiles, AssignMan, DepNo, FirstSupportMan, SecondSupportMan, " + Environment.NewLine +
                                       "        AssignReason, ReasonNote, UsedMiles, UsedMin, Remark, BuMan, BuDate, ModifyMan, ModifyDate, TargetCarID) " + Environment.NewLine +
                                       "select '" + vListNoItem + "', '" + vMaxItems.ToString("D4") + "', 'DEL',ListNo, CaseDate, Car_ID, " + Environment.NewLine +
                                       "       StartPosition, StartDate, StartTime, StartMeterMiles, FinishPosition, FinishDate, FinishTime, FinishMeterMiles, " + Environment.NewLine +
                                       "       AssignMan, DepNo, FirstSupportMan, SecondSupportMan, AssignReason, ReasonNote, UsedMiles, UsedMin, " + Environment.NewLine +
                                       "       Remark, BuMan, BuDate, ModifyMan, ModifyDate, TargetCarID " + Environment.NewLine +
                                       "  from RescueCarAssignList " + Environment.NewLine +
                                       " where ListNo = '" + vListNo + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vExecSQL_Temp);
                    sdsRescueCarAssignDetail.DeleteParameters.Clear();
                    sdsRescueCarAssignDetail.DeleteParameters.Add(new Parameter("ListNo", DbType.String, vListNo.Trim()));
                    sdsRescueCarAssignDetail.Delete();
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
                Response.Write("alert('請先選擇要刪除的派車單！')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
        }

        protected void sdsRescueCarAssignDetail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridRescueCarAssignList.DataBind();
            }
        }

        protected void sdsRescueCarAssignDetail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridRescueCarAssignList.DataBind();
            }
        }

        protected void sdsRescueCarAssignDetail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eCaseDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eCaseDate_INS");
            if (eCaseDate_INS != null)
            {
                DateTime vCaseDate_Temp = (eCaseDate_INS.Text.Trim() != "") ? DateTime.Parse(eCaseDate_INS.Text.Trim()) : DateTime.Today;
                TextBox eStartDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eStartDate_INS");
                DateTime vStartDate_Temp = (eStartDate_INS.Text.Trim() != "") ? DateTime.Parse(eStartDate_INS.Text.Trim()) : DateTime.Today;
                TextBox eFinishDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishDate_INS");
                DateTime vFinishDate_Temp = (eFinishDate_INS.Text.Trim() != "") ? DateTime.Parse(eFinishDate_INS.Text.Trim()) : DateTime.Today;
                string vCaseYM = vCaseDate_Temp.Year.ToString() + vCaseDate_Temp.Month.ToString("D2");
                string vSQLStr_Temp = "select max(ListNo) MaxNo from RescueCarAssignList where ListNo like '" + vCaseYM + "%' ";
                string vMaxNoStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                int vNewIndex = (vMaxNoStr != "") ? Int32.Parse(vMaxNoStr.Replace(vCaseYM, "")) + 1 : 1;
                string vListNo = vCaseYM + vNewIndex.ToString("D4");
                //Session["RescueNo"] = vListNo;
                Label eBuDate_Temp = (Label)fvRescueCarAssignDetail.FindControl("eBuDate_INS");
                DateTime vBuDate_Temp = (eBuDate_Temp.Text.Trim() != "") ? DateTime.Parse(eBuDate_Temp.Text.Trim()) : DateTime.Today;
                e.Command.Parameters["@CaseDate"].Value = vCaseDate_Temp;
                e.Command.Parameters["@StartDate"].Value = vStartDate_Temp;
                e.Command.Parameters["@FinishDate"].Value = vFinishDate_Temp;
                e.Command.Parameters["@BuDate"].Value = vBuDate_Temp;
                e.Command.Parameters["@ListNo"].Value = vListNo;
            }
        }

        protected void eCaseDate_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eCaseDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eCaseDate_Edit");
            TextBox eStartDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eStartDate_Edit");
            TextBox eFinishDate_Edit = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishDate_Edit");
            eStartDate_Edit.Text = (eCaseDate_Edit.Text.Trim() != "") ? eCaseDate_Edit.Text.Trim() : eStartDate_Edit.Text.Trim();
            eFinishDate_Edit.Text = (eCaseDate_Edit.Text.Trim() != "") ? eCaseDate_Edit.Text.Trim() : eFinishDate_Edit.Text.Trim();
        }

        protected void eCaseDate_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eCaseDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eCaseDate_INS");
            TextBox eStartDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eStartDate_INS");
            TextBox eFinishDate_INS = (TextBox)fvRescueCarAssignDetail.FindControl("eFinishDate_INS");
            eStartDate_INS.Text = (eCaseDate_INS.Text.Trim() != "") ? eCaseDate_INS.Text.Trim() : eStartDate_INS.Text.Trim();
            eFinishDate_INS.Text = (eCaseDate_INS.Text.Trim() != "") ? eCaseDate_INS.Text.Trim() : eFinishDate_INS.Text.Trim();
        }
    }
}