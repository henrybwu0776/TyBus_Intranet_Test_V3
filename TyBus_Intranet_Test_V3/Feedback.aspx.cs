using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Net;
using System.Net.Mail;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class Feedback : System.Web.UI.Page
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
                    //動態掛載日期輸入視窗
                    string vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eDay_B_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eDay_B_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eDay_E_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eDay_E_Search.Attributes["onClick"] = vDateScript;

                    if (!IsPostBack)
                    {
                        plPrint.Visible = false;
                        eDay_B_Search.Text = "";
                        eDay_E_Search.Text = "";
                    }
                    BindDataList();
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

        private void BindDataList()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "SELECT sheetno, day, " + Environment.NewLine +
                                "       (SELECT [NAME] FROM EMPLOYEE WHERE(EMPNO = s.buman)) AS BuMan, " + Environment.NewLine +
                                "       CASE WHEN isnull(FixType, '01') = '01' THEN '軟體' " + Environment.NewLine +
                                "            WHEN isnull(FixType, '01') = '02' THEN '硬體' " + Environment.NewLine +
                                "            ELSE '資訊評估' END AS FixType, " + Environment.NewLine +
                                "       CASE WHEN isnull(aboutReport, 'X') = 'V' THEN '列印報表' ELSE '' END + " + Environment.NewLine +
                                "       CASE WHEN isnull(aboutDataModify, 'X') = 'V' THEN '資料修改' ELSE '' END + " + Environment.NewLine +
                                "       CASE WHEN isnull(aboutDesign, 'X') = 'V' THEN '設計資料或表單' ELSE '' END + " + Environment.NewLine +
                                "       CASE WHEN isnull(aboutHardDriver, 'X') = 'V' THEN '電腦故障' ELSE '' END + " + Environment.NewLine +
                                "       CASE WHEN isnull(aboutSetting, 'X') = 'V' THEN '安裝或設定' ELSE '' END + " + Environment.NewLine +
                                "       CASE WHEN isnull(aboutSurver, 'X') = 'V' THEN '偵測或補發' ELSE '' END + " + Environment.NewLine +
                                "       CASE WHEN isnull(aboutPurchase, 'X') = 'V' THEN '購置' ELSE '' END + " + Environment.NewLine +
                                "       CASE WHEN isnull(aboutOthers, 'X') = 'V' THEN '其他' ELSE '' END AS FixNote, " + Environment.NewLine +
                                "       FixRemark, (SELECT [NAME] FROM EMPLOYEE WHERE(EMPNO = s.Handler)) AS Handler, AssignDate, " + Environment.NewLine +
                                "       CASE WHEN isnull(Disposal, '01') = '01' THEN '未派工' " + Environment.NewLine +
                                "            WHEN isnull(Disposal, '01') = '02' THEN '施工中' " + Environment.NewLine +
                                "            ELSE '已完工' END AS Disposal, FixFinishDate " + Environment.NewLine +
                                "  FROM SHEETA AS s " + Environment.NewLine +
                                " WHERE (SheetType = 'Z') ";
            //過濾報修申請人
            string vWStr_BuildMan = (eBuildMan_Search.Text.Trim() != "") ? "   AND s.BuMan = '" + eBuildMan_Search.Text.Trim() + "' " + Environment.NewLine : "";

            //過濾申請日期
            DateTime vDay_S;
            DateTime vDay_E;
            string vWStr_Day = "";
            if ((eDay_B_Search.Text.Trim() != "") && (eDay_E_Search.Text.Trim() != ""))
            {
                vDay_S = DateTime.Parse(eDay_B_Search.Text.Trim());
                vDay_E = DateTime.Parse(eDay_E_Search.Text.Trim());
                vWStr_Day = "   AND s.Day between '" + vDay_S.Year.ToString("D4") + "/" + vDay_S.ToString("MM/dd") + "' AND '" + vDay_E.Year.ToString("D4") + "/" + vDay_E.ToString("MM/dd") + "' " + Environment.NewLine;
            }
            else if ((eDay_B_Search.Text.Trim() != "") && (eDay_E_Search.Text.Trim() == ""))
            {
                vDay_S = DateTime.Parse(eDay_B_Search.Text.Trim());
                vWStr_Day = "   AND s.Day = '" + vDay_S.Year.ToString("D4") + "/" + vDay_S.ToString("MM/dd") + "' " + Environment.NewLine;
            }
            else if ((eDay_B_Search.Text.Trim() == "") && (eDay_E_Search.Text.Trim() != ""))
            {
                vDay_E = DateTime.Parse(eDay_E_Search.Text.Trim());
                vWStr_Day = "   AND s.Day = '" + vDay_E.Year.ToString("D4") + "/" + vDay_E.ToString("MM/dd") + "' " + Environment.NewLine;
            }
            else
            {
                vWStr_Day = "";
            }

            //過濾處理進度
            string vWStr_Disposal = (eDisposal_Search.SelectedValue != "00") ? "   AND isnull(s.Disposal, '01') = '" + eDisposal_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";

            //過濾關鍵字
            string vWStr_KeyWord = (eKeyword_Search.Text.Trim() != "") ? "   AND ((s.FixRemark like '%" + eKeyword_Search.Text.Trim() + "%') or (s.WorkReport like '%" + eKeyword_Search.Text.Trim() + "%'))" + Environment.NewLine : "";
            sdsFixRequestList.SelectCommand = "";
            sdsFixRequestList.SelectCommand = vSelectStr + vWStr_BuildMan + vWStr_Day + vWStr_Disposal + vWStr_KeyWord + Environment.NewLine + " ORDER BY s.Day DESC";
            gridFixRequestList.DataBind();
        }

        private void FormView_ReadOnlyMode()
        {
            Button bbTemp = (Button)fvFixRequestDetail.FindControl("bbModify");
            Button bbDel = (Button)fvFixRequestDetail.FindControl("bbDel");
            Label lbDisposal = (Label)fvFixRequestDetail.FindControl("eDisposal_List");
            if (lbDisposal != null)
            {
                switch (lbDisposal.Text.Trim())
                {
                    case "01": //未派工
                        bbTemp.Text = "派工作業";
                        bbTemp.Visible = true;
                        bbDel.Visible = true;
                        break;
                    case "02": //施工中
                        bbTemp.Text = "完工回報";
                        bbTemp.Visible = true;
                        bbDel.Visible = false;
                        break;
                    case "03": //已完工
                        bbTemp.Visible = false;
                        bbDel.Visible = false;
                        break;
                }
            }
        }

        private void FormView_EditMode()
        {
            Button bbOK = (Button)fvFixRequestDetail.FindControl("bbOK_Edit");
            Label lbDisposal = (Label)fvFixRequestDetail.FindControl("eDisposal_Edit");
            Label eHandler = (Label)fvFixRequestDetail.FindControl("eHandler_Edit");
            Label eHandler_C = (Label)fvFixRequestDetail.FindControl("eHandler_C_Edit");
            TextBox eAssignDate = (TextBox)fvFixRequestDetail.FindControl("eAssignDate_Edit");
            TextBox eFinishDate = (TextBox)fvFixRequestDetail.FindControl("eFinishDate_Edit");
            TextBox eChiefNote = (TextBox)fvFixRequestDetail.FindControl("eChiefNote_Edit");
            TextBox eWorkReport = (TextBox)fvFixRequestDetail.FindControl("eWorkReport_Edit");
            DropDownList ddlHandler = (DropDownList)fvFixRequestDetail.FindControl("ddlHandler_Edit");

            if (lbDisposal != null)
            {
                string vDateURL = "";
                string vDateScript = "";

                bbOK.Enabled = false;
                eAssignDate.Enabled = false;
                eFinishDate.Enabled = false;
                eChiefNote.Enabled = false;
                eWorkReport.Enabled = false;
                ddlHandler.Visible = false;
                eHandler_C.Visible = false;
                DateTime vToday = DateTime.Today;

                switch (lbDisposal.Text.Trim())
                {
                    case "01": //未派工
                        bbOK.Text = "派工";
                        bbOK.Enabled = true;
                        vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eAssignDate.ClientID;
                        vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                        eAssignDate.Attributes["onClick"] = vDateScript;
                        eAssignDate.Text = (vToday.Year - 1911).ToString("D3") + "/" + vToday.ToString("MM/dd");
                        eAssignDate.Enabled = true;
                        ddlHandler.Visible = true;
                        eChiefNote.Enabled = true;
                        eHandler.Text = "";
                        break;
                    case "02": //施工中
                        bbOK.Text = "完工";
                        bbOK.Enabled = true;
                        vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eFinishDate.ClientID;
                        vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                        eFinishDate.Attributes["onClick"] = vDateScript;
                        eFinishDate.Text = (vToday.Year - 1911).ToString("D3") + "/" + vToday.ToString("MM/dd");
                        eFinishDate.Enabled = true;
                        eWorkReport.Enabled = true;
                        eHandler_C.Visible = true;
                        break;
                    case "03": //已完工

                        break;
                }
            }
        }

        protected void eBuildMan_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vBuildMan = eBuildMan_Search.Text.Trim();
            string vBuildManName = "";
            string vSQLStr = "select [Name] from Employee where EMpNo = '" + vBuildMan + "' and LeaveDay is null";
            vBuildManName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vBuildManName == "")
            {
                vBuildManName = vBuildMan;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vBuildManName + "' and LeaveDay is null";
                vBuildMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eBuildMan_Search.Text = vBuildMan;
            eBuildManName_Search.Text = vBuildManName;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            BindDataList();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvFixRequestDetail_DataBound(object sender, EventArgs e)
        {
            switch (fvFixRequestDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    FormView_ReadOnlyMode();
                    break;
                case FormViewMode.Edit:
                    FormView_EditMode();
                    break;
                case FormViewMode.Insert:
                    break;
                default:
                    break;
            }
        }

        protected void ddlHandler_Edit_TextChanged(object sender, EventArgs e)
        {
            DropDownList ddlHandler = (DropDownList)fvFixRequestDetail.FindControl("ddlHandler_Edit");
            Label eHandler = (Label)fvFixRequestDetail.FindControl("eHandler_Edit");
            eHandler.Text = ddlHandler.SelectedValue.Trim();
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {
            Label eSheetNo = (Label)fvFixRequestDetail.FindControl("eSheetNo_List");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                string vSheetNo = eSheetNo.Text.Trim();
                string vSelectStr = "select SheetNo , convert(varchar, Day, 111) [Day], " + Environment.NewLine +
                                    "       DepNo, (select [Name] from Department where Department.DepNo = s.DepNo) DepName, " + Environment.NewLine +
                                    "       BuMan, (select [Name] from Employee where Employee.EmpNo = s.BuMan) BuManName, " + Environment.NewLine +
                                    "       case when isnull(FixType,'01') = '01' then '軟體' " + Environment.NewLine +
                                    "            when isnull(FixType,'01') = '02' then '硬體' " + Environment.NewLine +
                                    "            else '資訊評估' end FixType_C, FixType, " + Environment.NewLine +
                                    "       case when isnull(aboutReport,'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                    "       case when isnull(aboutDataModify,'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                    "       case when isnull(aboutDesign,'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                    "       case when isnull(aboutHardDriver,'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                    "       case when isnull(aboutSetting,'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                    "       case when isnull(aboutSurver,'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                    "       case when isnull(aboutPurchase,'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                    "       case when isnull(aboutOthers,'X') = 'V' then '其他' else '' end FixNote, " + Environment.NewLine +
                                    "       FixRemark , Handler, (select [Name] from Employee where Employee.EmpNo = s.Handler) HandlerName, " + Environment.NewLine +
                                    "       convert(varchar, AssignDate, 111) AssignDate, " + Environment.NewLine +
                                    "       case when isnull(Disposal,'01') = '01' then '未派工' " + Environment.NewLine +
                                    "            when isnull(Disposal,'01') = '02' then '施工中' " + Environment.NewLine +
                                    "            else '已完工' end Disposal_C, Disposal, convert(varchar, FixFinishDate, 111) FixFinishDate, WorkReport " + Environment.NewLine +
                                    "  from SheetA s " + Environment.NewLine +
                                    " where SheetType = 'Z' and SheetNo = '" + vSheetNo + "' ";
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable("FixSheetP");
                daPrint.Fill(dtPrint);
                ReportDataSource rdsPrint = new ReportDataSource("FixSheetP", dtPrint);
                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = @"Report\FixSheetP.rdlc";
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plShowData.Visible = false;
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plShowData.Visible = true;
        }

        protected void sdsFixRequestDetail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label lbSheetNo = (Label)fvFixRequestDetail.FindControl("eSheetNo_Edit");
            Label lbDisposal = (Label)fvFixRequestDetail.FindControl("eDisposal_Edit");
            Label lbHandler = (Label)fvFixRequestDetail.FindControl("eHandler_Edit");
            Label lbHandler_C = (Label)fvFixRequestDetail.FindControl("eHandler_C_Edit");
            Label lbBuMan = (Label)fvFixRequestDetail.FindControl("eBuMan_Edit");
            Label lbBuMan_C = (Label)fvFixRequestDetail.FindControl("eBuMan_C_Edit");
            Label lbFixType_Edit = (Label)fvFixRequestDetail.FindControl("eFixType_Edit");
            Label lbFixNote_Edit = (Label)fvFixRequestDetail.FindControl("eFixNote_Edit");

            TextBox eAssignDate = (TextBox)fvFixRequestDetail.FindControl("eAssignDate_Edit");
            TextBox eFinishDate = (TextBox)fvFixRequestDetail.FindControl("eFinishDate_Edit");
            TextBox eFixRemark = (TextBox)fvFixRequestDetail.FindControl("eFixRemark_Edit");
            TextBox eChiefNote = (TextBox)fvFixRequestDetail.FindControl("eChiefNote_Edit");
            TextBox eWorkReport = (TextBox)fvFixRequestDetail.FindControl("eWorkReport_Edit");

            DateTime vTempDate_1;
            DateTime vTempDate_2;
            string vSheetNo_Edit = lbSheetNo.Text.Trim();
            string vRemark_Edit = eFixRemark.Text.Trim();
            string vTargetEmail = "";
            string vTargetEmail_CC = "";

            if (lbDisposal != null)
            {
                switch (lbDisposal.Text.Trim())
                {
                    case "01": //未派工
                        if (eAssignDate.Text.Trim() != "")
                        {
                            vTempDate_1 = DateTime.Parse(eAssignDate.Text.Trim());
                            e.Command.Parameters["@AssignDate"].Value = vTempDate_1;
                            e.Command.Parameters["@Disposal"].Value = "02";
                            vTargetEmail = PF.GetValue(vConnStr, "select [eMail] from Employee where Empno = '" + lbHandler.Text.Trim() + "' ", "eMail");
                            vTargetEmail_CC = PF.GetValue(vConnStr, "select [eMail] from Employee where Empno = '" + lbBuMan.Text.Trim() + "' ", "eMail");
                            if (vTargetEmail.Trim() != "")
                            {
                                using (MailMessage vMailMessage = new MailMessage())
                                {
                                    vMailMessage.From = new MailAddress("webmaster@tybus.com.tw");
                                    vMailMessage.To.Add(new MailAddress(vTargetEmail));
                                    if (eChiefNote.Text.Trim() != "")
                                    {
                                        vMailMessage.CC.Add(new MailAddress(vTargetEmail_CC));
                                    }
                                    vMailMessage.Subject = "線上系統報修派工通知";
                                    vMailMessage.Body = "報修單號：" + vSheetNo_Edit + Environment.NewLine +
                                                        "報修人員：" + lbBuMan_C.Text.Trim() + Environment.NewLine +
                                                        "報修大類：" + lbFixType_Edit.Text.Trim() + Environment.NewLine +
                                                        "報修項目：" + lbFixNote_Edit.Text.Trim() + Environment.NewLine +
                                                        "說明事項：" + eFixRemark.Text.Trim() + Environment.NewLine +
                                                        "主管交辦：" + eChiefNote.Text.Trim() + Environment.NewLine +
                                                        "派工日期：" + eAssignDate.Text.Trim();
                                    //SmtpClient vMailClient = new SmtpClient();
                                    /* 測試改用 GMail
                                    vMailClient.Host = "Mail.Tybus.com.tw";
                                    vMailClient.Port = 25;
                                    vMailClient.Credentials = new NetworkCredential("webmaster", "TyBusMis9999");
                                    */
                                    //2022.06.09 因應 GMail 政策改變，試看看改用 "應用程式密碼登入" 的方式操作 GMail
                                    //vMailClient.Credentials = new NetworkCredential("tybusmis@gmail.com", "TyBusMis9999");
                                    // 2023.11.24 改回用公司主機
                                    /*NetworkCredential cred = new NetworkCredential("tybusmis@gmail.com", "icyemoozikjooazh");
                                    CredentialCache credCache = new CredentialCache();
                                    vMailClient.Host = "smtp.gmail.com";
                                    vMailClient.Port = 587;
                                    credCache.Add(vMailClient.Host, vMailClient.Port, "login", cred);
                                    vMailClient.EnableSsl = true;
                                    vMailClient.Credentials = credCache; //*/
                                    string smtpAddress = "172.20.3.142";
                                    bool enableSSL = false;
                                    string eMailFrom = "webmaster@tybus.com.tw";
                                    string eMailPassword = "TyBusMis9999";
                                    try
                                    {
                                        SmtpClient vMailClient = new SmtpClient(smtpAddress, 587);
                                        vMailClient.Credentials = new NetworkCredential(eMailFrom, eMailPassword);
                                        vMailClient.EnableSsl = enableSSL;
                                        vMailClient.Send(vMailMessage);
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
                        break;
                    case "02": //施工中
                        if (eFinishDate.Text.Trim() != "")
                        {
                            vTempDate_1 = DateTime.Parse(eAssignDate.Text.Trim());
                            vTempDate_2 = DateTime.Parse(eFinishDate.Text.Trim());
                            e.Command.Parameters["@AssignDate"].Value = vTempDate_1;
                            e.Command.Parameters["@FixFinishDate"].Value = vTempDate_2;
                            e.Command.Parameters["@Disposal"].Value = "03";
                            vTargetEmail = PF.GetValue(vConnStr, "select [eMail] from Employee where Empno = '" + lbBuMan.Text.Trim() + "' ", "eMail");
                            if (vTargetEmail.Trim() != "")
                            {
                                using (MailMessage vMailMessage = new MailMessage())
                                {
                                    vMailMessage.From = new MailAddress("webmaster@tybus.com.tw");
                                    vMailMessage.To.Add(new MailAddress(vTargetEmail));
                                    vMailMessage.Subject = "線上系統報修完工通知";
                                    vMailMessage.Body = (eWorkReport.Text.Trim() != "") ?
                                                        // 有註明處置情況時的郵件內文
                                                        "報修單號：" + vSheetNo_Edit + Environment.NewLine +
                                                        "處理人員：" + lbHandler_C.Text.Trim() + Environment.NewLine +
                                                        "報修大類：" + lbFixType_Edit.Text.Trim() + Environment.NewLine +
                                                        "報修項目：" + lbFixNote_Edit.Text.Trim() + Environment.NewLine +
                                                        "說明事項：" + eFixRemark.Text.Trim() + Environment.NewLine +
                                                        "完工日期：" + eFinishDate.Text.Trim() + Environment.NewLine +
                                                        "處置情況：" + eWorkReport.Text.Trim() :
                                                        //沒有說明處置情況時的郵件內文
                                                        "報修單號：" + vSheetNo_Edit + Environment.NewLine +
                                                        "處理人員：" + lbHandler_C.Text.Trim() + Environment.NewLine +
                                                        "報修大類：" + lbFixType_Edit.Text.Trim() + Environment.NewLine +
                                                        "報修項目：" + lbFixNote_Edit.Text.Trim() + Environment.NewLine +
                                                        "說明事項：" + eFixRemark.Text.Trim() + Environment.NewLine +
                                                        "完工日期：" + eFinishDate.Text.Trim();
                                    //SmtpClient vMailClient = new SmtpClient();
                                    /* 測試改用 GMail
                                    vMailClient.Host = "Mail.Tybus.com.tw";
                                    vMailClient.Port = 25;
                                    vMailClient.Credentials = new NetworkCredential("webmaster", "TyBusMis9999");
                                    */
                                    //2022.06.09 因應 GMail 政策改變，試看看改用 "應用程式密碼登入" 的方式操作 GMail
                                    //vMailClient.Credentials = new NetworkCredential("tybusmis@gmail.com", "TyBusMis9999");
                                    // 2023.11.24 改回用公司主機
                                    /*NetworkCredential cred = new NetworkCredential("tybusmis@gmail.com", "icyemoozikjooazh");
                                    CredentialCache credCache = new CredentialCache();
                                    vMailClient.Host = "smtp.gmail.com";
                                    vMailClient.Port = 587;
                                    credCache.Add(vMailClient.Host, vMailClient.Port, "login", cred);
                                    vMailClient.EnableSsl = true;
                                    vMailClient.Credentials = credCache; //*/
                                    string smtpAddress = "172.20.3.142";
                                    bool enableSSL = false;
                                    string eMailFrom = "webmaster@tybus.com.tw";
                                    string eMailPassword = "TyBusMis9999";
                                    try
                                    {
                                        SmtpClient vMailClient = new SmtpClient(smtpAddress, 587);
                                        vMailClient.Credentials = new NetworkCredential(eMailFrom, eMailPassword);
                                        vMailClient.EnableSsl = enableSSL;
                                        vMailClient.Send(vMailMessage);
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
                        break;
                    case "03": //已完工
                        break;
                }
            }
        }

        protected void bbDel_Click(object sender, EventArgs e)
        {
            string vDelCommandStr = "Delete SheetA where SheetNo = @SheetNo";
            Label eSheetNo_List = (Label)fvFixRequestDetail.FindControl("eSheetNo_List");
            if ((eSheetNo_List != null) && (eSheetNo_List.Text.Trim() != ""))
            {
                try
                {
                    sdsFixRequestDetail.DeleteParameters.Clear();
                    sdsFixRequestDetail.DeleteCommand = vDelCommandStr;
                    sdsFixRequestDetail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, eSheetNo_List.Text.Trim()));
                    sdsFixRequestDetail.Delete();
                    gridFixRequestList.DataBind();
                    fvFixRequestDetail.DataBind();
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