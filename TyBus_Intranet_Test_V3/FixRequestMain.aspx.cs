using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Net;
using System.Net.Mail;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class FixRequestMain : Page
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
                    string vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBudate_S_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBudate_S_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBudate_E_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBudate_E_Search.Attributes["onClick"] = vDateScript;

                    if (!IsPostBack)
                    {
                        eBudate_S_Search.Text = "";
                        eBudate_E_Search.Text = "";
                    }
                    GetDataList();
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

        private void GetDataList()
        {
            string vSelectEmpStr = "select s.SheetNo, s.Day, " + Environment.NewLine +
                                   "       s.BuMan, (select [Name] from Employee where Employee.EmpNo = s.BuMan) BuManName, " + Environment.NewLine +
                                   "       case when isnull(s.FixType, '01') = '01' then '軟體'" + Environment.NewLine +
                                   "            when isnull(s.FixType, '01') = '02' then '硬體'" + Environment.NewLine +
                                   "            else '資訊評估' end as FixType, " + Environment.NewLine +
                                   "       case when isnull(aboutReport,'X') = 'V' then '列印報表' else '' end + " + Environment.NewLine +
                                   "       case when isnull(aboutDataModify,'X') = 'V' then '資料修改' else '' end + " + Environment.NewLine +
                                   "       case when isnull(aboutDesign,'X') = 'V' then '設計資料或表單' else '' end + " + Environment.NewLine +
                                   "       case when isnull(aboutHardDriver,'X') = 'V' then '電腦故障' else '' end + " + Environment.NewLine +
                                   "       case when isnull(aboutSetting,'X') = 'V' then '安裝或設定' else '' end + " + Environment.NewLine +
                                   "       case when isnull(aboutSurver,'X') = 'V' then '偵測或補發' else '' end + " + Environment.NewLine +
                                   "       case when isnull(aboutPurchase,'X') = 'V' then '購置' else '' end + " + Environment.NewLine +
                                   "       case when isnull(aboutOthers,'X') = 'V' then '其他' else '' end as FixNote," + Environment.NewLine +
                                   "       FixRemark, " + Environment.NewLine +
                                   "       (select Name from Employee where Employee.EmpNo = s.Handler) as Handler, " + Environment.NewLine +
                                   "       AssignDate, " + Environment.NewLine +
                                   "       case when isnull(Disposal,'01') = '01' then '未派工'" + Environment.NewLine +
                                   "            when isnull(Disposal,'01') = '02' then '施工中'" + Environment.NewLine +
                                   "            else '已完工' end as Disposal, FixFinishDate " + Environment.NewLine +
                                   "  from SheetA s " + Environment.NewLine +
                                   " where SheetType = 'Z' ";
            string vWStr_DepNo = "";
            string vWStr_EmpNo = "";
            string vWStr_BuDate = "";
            DateTime vStartDate;
            DateTime vEndDate;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vWStr_DepNo = (vLoginDepNo != "09") ? "   and s.DepNo = '" + vLoginDepNo.Trim() + "' " + Environment.NewLine : "";
            vWStr_EmpNo = (eBuMan_Search.Text.Trim() != "") ? "   and s.BuMan = '" + eBuMan_Search.Text.Trim() + "' " + Environment.NewLine : "";

            if ((eBudate_S_Search.Text.Trim() != "") && (eBudate_E_Search.Text.Trim() != ""))
            {
                vStartDate = DateTime.Parse(eBudate_S_Search.Text.Trim());
                vEndDate = DateTime.Parse(eBudate_E_Search.Text.Trim());
                vWStr_BuDate = "   and s.Day between '" + vStartDate.Year.ToString("D4") + "/" + vStartDate.ToString("MM/dd") + "' and '" + vEndDate.Year.ToString("D4") + "/" + vEndDate.ToString("MM/dd") + "' " + Environment.NewLine;
            }
            else if ((eBudate_S_Search.Text.Trim() != "") && (eBudate_E_Search.Text.Trim() == ""))
            {
                vStartDate = DateTime.Parse(eBudate_S_Search.Text.Trim());
                vWStr_BuDate = "   and s.Day = '" + vStartDate.Year.ToString("D4") + "/" + vStartDate.ToString("MM/dd") + "' " + Environment.NewLine;
            }
            else if ((eBudate_S_Search.Text.Trim() == "") && (eBudate_E_Search.Text.Trim() != ""))
            {
                vEndDate = DateTime.Parse(eBudate_E_Search.Text.Trim());
                vWStr_BuDate = "   and s.Day = '" + vEndDate.Year.ToString("D4") + "/" + vEndDate.ToString("MM/dd") + "' " + Environment.NewLine;
            }
            vSelectEmpStr = vSelectEmpStr + vWStr_DepNo + vWStr_BuDate + vWStr_EmpNo + Environment.NewLine + " order by s.SheetNo DESC";
            sdsFixRequestList.SelectCommand = "";
            sdsFixRequestList.SelectCommand = vSelectEmpStr;
            gridShowData.DataBind();
        }

        private void SendMail(string fRequestSheetNo)
        {
            if (fRequestSheetNo != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSQLStr_Check = "select isnull(IsCreate, 0) IsCreate from SheetA where SheetNo = '" + fRequestSheetNo + "' ";
                string vIsCreate = PF.GetValue(vConnStr, vSQLStr_Check, "IsCreate");
                if (vIsCreate == "False")
                {
                    string vSQLStr_GetRequest = "SELECT sheetno AS SheetNo, CONVERT(varchar, day, 111) AS Day, " + Environment.NewLine +
                                                "       (SELECT [NAME] FROM DEPARTMENT WHERE(DEPNO = s.depno)) AS DepName, " + Environment.NewLine +
                                                "       (SELECT [NAME] FROM EMPLOYEE WHERE(EMPNO = s.buman)) AS BuManName, " + Environment.NewLine +
                                                "       CASE WHEN isnull(FixType, '01') = '01' THEN '軟體' " + Environment.NewLine +
                                                "            WHEN isnull(FixType, '01') = '02' THEN '硬體' " + Environment.NewLine +
                                                "            ELSE '資訊評估' END AS FixType_C, " + Environment.NewLine +
                                                "       CASE WHEN isnull(aboutReport, 'X') = 'V' THEN '列印報表 ' ELSE '' END + " + Environment.NewLine +
                                                "       CASE WHEN isnull(aboutDataModify, 'X') = 'V' THEN '資料修改 ' ELSE '' END + " + Environment.NewLine +
                                                "       CASE WHEN isnull(aboutDesign, 'X') = 'V' THEN '設計資料或表單 ' ELSE '' END + " + Environment.NewLine +
                                                "       CASE WHEN isnull(aboutHardDriver, 'X') = 'V' THEN '電腦故障 ' ELSE '' END + " + Environment.NewLine +
                                                "       CASE WHEN isnull(aboutSetting, 'X') = 'V' THEN '安裝或設定 ' ELSE '' END + " + Environment.NewLine +
                                                "       CASE WHEN isnull(aboutSurver, 'X') = 'V' THEN '偵測或補發 ' ELSE '' END + " + Environment.NewLine +
                                                "       CASE WHEN isnull(aboutPurchase, 'X') = 'V' THEN '購置 ' ELSE '' END + " + Environment.NewLine +
                                                "       CASE WHEN isnull(aboutOthers, 'X') = 'V' THEN '其他 ' ELSE '' END AS FixItems, " + Environment.NewLine +
                                                "       FixRemark, (SELECT [NAME] FROM EMPLOYEE WHERE(EMPNO = s.Handler)) AS HandlerName, " + Environment.NewLine +
                                                "       CONVERT(varchar, AssignDate, 111) AS AssignDate, " + Environment.NewLine +
                                                "       CASE WHEN isnull(Disposal, '01') = '01' THEN '未派工' " + Environment.NewLine +
                                                "            WHEN isnull(Disposal, '01') = '02' THEN '施工中' ELSE '已完工' END AS Disposal_C, " + Environment.NewLine +
                                                "       CONVERT(varchar, FixFinishDate, 111) AS FixFinishDate " + Environment.NewLine +
                                                "  FROM SHEETA AS s " + Environment.NewLine +
                                                " WHERE sheetno = '" + fRequestSheetNo + "' ";
                    using (SqlConnection connMail = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdMail = new SqlCommand(vSQLStr_GetRequest, connMail);
                        connMail.Open();
                        SqlDataReader drMail = cmdMail.ExecuteReader();
                        if (drMail.HasRows)
                        {
                            while (drMail.Read())
                            {
                                MailMessage vMailMessage = new MailMessage();
                                vMailMessage.From = new MailAddress("webmaster@tybus.com.tw");
                                vMailMessage.To.Add(new MailAddress("computer@tybus.com.tw"));
                                vMailMessage.Subject = "線上系統報修需求";
                                vMailMessage.Body = "報修單號：" + drMail["SheetNo"] + Environment.NewLine +
                                                    "報修單位：" + drMail["DepName"] + Environment.NewLine +
                                                    "報修人員：" + drMail["BuManNAme"] + Environment.NewLine +
                                                    "報修大類：" + drMail["FixType_C"] + Environment.NewLine +
                                                    "報修項目：" + drMail["FixItems"] + Environment.NewLine +
                                                    "說明事項：" + drMail["FixRemark"];
                                //SmtpClient vMailClient = new SmtpClient();
                                //改用 GMail
                                //vMailClient.Host = "Mail.Tybus.com.tw";
                                //vMailClient.Port = 25;
                                //vMailClient.Credentials = new NetworkCredential("webmaster", "TyBusMis9999");
                                //2022.06.09 因應 GMail 政策改變，試看看改用 "應用程式密碼登入" 的方式操作 GMail
                                //vMailClient.Credentials = new NetworkCredential("tybusmis@gmail.com", "TyBusMis9999");
                                //2023.11.24 改回用公司主機寄信
                                /*NetworkCredential cred = new NetworkCredential("tybusmis@gmail.com", "icyemoozikjooazh");
                                CredentialCache credCache = new CredentialCache();
                                vMailClient.Host = "smtp.gmail.com";
                                vMailClient.Port = 587;
                                credCache.Add(vMailClient.Host, vMailClient.Port, "login", cred); 
                                vMailClient.EnableSsl = true;
                                vMailClient.Credentials = credCache;//*/
                                //===================================================================================================================
                                string smtpAddress = "172.20.3.142";
                                bool enableSSL = true;
                                string eMailFrom = "webmaster@tybus.com.tw";
                                string eMailPassword = "TyBusMis9999";
                                try
                                {
                                    SmtpClient vMailClient = new SmtpClient(smtpAddress, 587);
                                    vMailClient.Credentials = new NetworkCredential(eMailFrom, eMailPassword);
                                    vMailClient.EnableSsl = enableSSL;
                                    vMailClient.Send(vMailMessage);
                                    Session["FixRequestNo"] = "";
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
        }

        protected void eBuMan_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo = eBuMan_Search.Text.Trim();
            string vEmpName = "";
            string vSelectEmp = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' and LeaveDay is null";
            vEmpName = PF.GetValue(vConnStr, vSelectEmp, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSelectEmp = "select EmpNo from Employee where [Name] = '" + vEmpName + "' and LeaveDay is null";
                vEmpNo = PF.GetValue(vConnStr, vSelectEmp, "EmpNo");
            }
            eBuMan_Search.Text = vEmpNo;
            eBuManName_Search.Text = vEmpName;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            GetDataList();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvFixRequest_DataBound(object sender, EventArgs e)
        {
            switch (fvFixRequest.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    //顯示模式
                    Label lbSeetNo_Temp = (Label)fvFixRequest.FindControl("eSheetNo_List");
                    Button bbPrintFixQuery_Temp = (Button)fvFixRequest.FindControl("bbPrintFixQuery_List");
                    if ((bbPrintFixQuery_Temp != null) && (lbSeetNo_Temp.Text.Trim() != ""))
                    {
                        bbPrintFixQuery_Temp.Visible = true;
                    }
                    break;
                case FormViewMode.Edit:
                    //編輯模式，因為報修單不允許編輯所以沒有做任何功能
                    break;
                case FormViewMode.Insert:
                    //新增模式
                    Session["FixRequestNo"] = "";
                    Label lbSheetType_Temp = (Label)fvFixRequest.FindControl("eSheetType_Insert");
                    lbSheetType_Temp.Text = "Z";
                    Label lbSheetClass_Temp = (Label)fvFixRequest.FindControl("eSheetClass_Insert");
                    lbSheetClass_Temp.Text = "1";
                    Label lbMode_Temp = (Label)fvFixRequest.FindControl("eMode_Insert");
                    lbMode_Temp.Text = "1";
                    Label lbUnitChief_Temp = (Label)fvFixRequest.FindControl("eUnitChief_Insert");
                    lbUnitChief_Temp.Text = "V";
                    TextBox eDay_Temp = (TextBox)fvFixRequest.FindControl("eDay_Insert");
                    eDay_Temp.Text = (DateTime.Today.Year - 1911).ToString("D3") + "/" + DateTime.Today.ToString("MM/dd");
                    TextBox eBuMan_Temp = (TextBox)fvFixRequest.FindControl("eBuMan_Insert");
                    eBuMan_Temp.Text = vLoginID;
                    Label eBuManName_Temp = (Label)fvFixRequest.FindControl("eBuManName_Insert");
                    eBuManName_Temp.Text = Session["LoginName"].ToString().Trim();
                    TextBox eDepNo_Temp = (TextBox)fvFixRequest.FindControl("eDepNo_Insert");
                    eDepNo_Temp.Text = Session["LoginDepNo"].ToString().Trim();
                    Label eDepName_Temp = (Label)fvFixRequest.FindControl("eDepName_Insert");
                    eDepName_Temp.Text = Session["LoginDepName"].ToString().Trim();
                    break;
            }
        }

        protected void ddlFixType_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label eFixType_Temp = (Label)fvFixRequest.FindControl("eFixType_Insert");
            DropDownList ddlFixType = (DropDownList)fvFixRequest.FindControl("ddlFixType_Insert");
            eFixType_Temp.Text = ddlFixType.SelectedValue;
        }

        protected void bbPrintFixQuery_List_Click(object sender, EventArgs e)
        {
            Label lbSheetNo_Print = (Label)fvFixRequest.FindControl("eSheetNo_List");
            string vSheetNo = lbSheetNo_Print.Text.Trim();
            if (vSheetNo != "")
            {
                string vSQLStr = "SELECT sheetno AS SheetNo, CONVERT(varchar, day, 111) AS Day, " + Environment.NewLine +
                                 "       (SELECT [NAME] FROM DEPARTMENT WHERE(DEPNO = s.depno)) AS DepName, " + Environment.NewLine +
                                 "       (SELECT [NAME] FROM EMPLOYEE WHERE(EMPNO = s.buman)) AS BuManName, " + Environment.NewLine +
                                 "       CASE WHEN isnull(FixType, '01') = '01' THEN '軟體' " + Environment.NewLine +
                                 "            WHEN isnull(FixType, '01') = '02' THEN '硬體' " + Environment.NewLine +
                                 "            ELSE '資訊評估' END AS FixType_C, " + Environment.NewLine +
                                 "       CASE WHEN isnull(aboutReport, 'X') = 'V' THEN '列印報表 ' ELSE '' END + " + Environment.NewLine +
                                 "       CASE WHEN isnull(aboutDataModify, 'X') = 'V' THEN '資料修改 ' ELSE '' END + " + Environment.NewLine +
                                 "       CASE WHEN isnull(aboutDesign, 'X') = 'V' THEN '設計資料或表單 ' ELSE '' END + " + Environment.NewLine +
                                 "       CASE WHEN isnull(aboutHardDriver, 'X') = 'V' THEN '電腦故障 ' ELSE '' END + " + Environment.NewLine +
                                 "       CASE WHEN isnull(aboutSetting, 'X') = 'V' THEN '安裝或設定 ' ELSE '' END + " + Environment.NewLine +
                                 "       CASE WHEN isnull(aboutSurver, 'X') = 'V' THEN '偵測或補發 ' ELSE '' END + " + Environment.NewLine +
                                 "       CASE WHEN isnull(aboutPurchase, 'X') = 'V' THEN '購置 ' ELSE '' END + " + Environment.NewLine +
                                 "       CASE WHEN isnull(aboutOthers, 'X') = 'V' THEN '其他 ' ELSE '' END AS FixItems, " + Environment.NewLine +
                                 "       FixRemark, (SELECT [NAME] FROM EMPLOYEE WHERE(EMPNO = s.Handler)) AS HandlerName, " + Environment.NewLine +
                                 "       CONVERT(varchar, AssignDate, 111) AS AssignDate, " + Environment.NewLine +
                                 "       CASE WHEN isnull(Disposal, '01') = '01' THEN '未派工' " + Environment.NewLine +
                                 "            WHEN isnull(Disposal, '01') = '02' THEN '施工中' ELSE '已完工' END AS Disposal_C, " + Environment.NewLine +
                                 "       CONVERT(varchar, FixFinishDate, 111) AS FixFinishDate, " + Environment.NewLine +
                                 "       ChiefNote, WorkReport " + Environment.NewLine +
                                 "  FROM SHEETA AS s " + Environment.NewLine +
                                 " WHERE sheetno = '" + vSheetNo + "' ";
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    using (SqlConnection connPrint = new SqlConnection(vConnStr))
                    {
                        SqlDataAdapter daPrint = new SqlDataAdapter(vSQLStr, connPrint);
                        connPrint.Open();
                        DataTable dtPrint = new DataTable("FixRequestP");
                        daPrint.Fill(dtPrint);
                        ReportDataSource rdsPrint = new ReportDataSource("FixRequestP", dtPrint);
                        rvPrint.LocalReport.DataSources.Clear();
                        rvPrint.LocalReport.ReportPath = @"Report\FixRequestP.rdlc";
                        rvPrint.LocalReport.DataSources.Add(rdsPrint);
                        rvPrint.LocalReport.Refresh();
                        plDetailData.Visible = false;
                        plPrint.Visible = true;
                    }
                }
                catch
                {

                }
            }
        }

        protected void HidePrint_Click(object sender, EventArgs e)
        {
            plDetailData.Visible = true;
            plPrint.Visible = false;
        }

        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            try
            {
                Label eSheetType_Temp = (Label)fvFixRequest.FindControl("eSheetType_Insert");
                DateTime vToday = DateTime.Today;
                string vSheetType = eSheetType_Temp.Text.Trim().ToUpper();
                string vSheetNo = vToday.Year.ToString("D4") + vToday.ToString("MMdd") + vSheetType;
                string vSelectStr = "select top 1 SheetNo from SheetA where SheetNo like '" + vSheetNo + "%' order by SheetNo DESC";
                string vLastSheetNo = PF.GetValue(vConnStr, vSelectStr, "SheetNo") ?? "";
                int vIndex = 0;
                if (vLastSheetNo != "")
                {
                    Int32.TryParse(vLastSheetNo.Replace(vSheetNo, ""), out vIndex);
                }
                vIndex++;
                vSheetNo = vSheetNo + vIndex.ToString("D3");
                TextBox eDay_Temp = (TextBox)fvFixRequest.FindControl("eDay_Insert");
                DateTime vDay = (eDay_Temp.Text.Trim() != "") ? DateTime.Parse(eDay_Temp.Text.Trim()) : DateTime.Today;
                TextBox eBuMan_Temp = (TextBox)fvFixRequest.FindControl("eBuMan_Insert");
                string vBuMan = eBuMan_Temp.Text.Trim();
                Label eSheetClass_Temp = (Label)fvFixRequest.FindControl("eSheetClass_Insert");
                string vSheetClass = eSheetClass_Temp.Text.Trim();
                Label eMode_Temp = (Label)fvFixRequest.FindControl("eMode_Insert");
                string vMode = eMode_Temp.Text.Trim();
                Label eUnitChief_Temp = (Label)fvFixRequest.FindControl("eUnitChief_Insert");
                string vUnitChief = eUnitChief_Temp.Text.Trim().ToUpper();
                Label eFixType_Temp = (Label)fvFixRequest.FindControl("eFixType_Insert");
                string vFixType = eFixType_Temp.Text.Trim();
                CheckBox eAboutReport_Temp = (CheckBox)fvFixRequest.FindControl("eAboutReport_Insert");
                string vAboutReport = (eAboutReport_Temp.Checked) ? "V" : "X";
                CheckBox eAboutDataModify_Temp = (CheckBox)fvFixRequest.FindControl("eAboutDataModify_Insert");
                string vAboutDataModify = (eAboutDataModify_Temp.Checked) ? "V" : "X";
                CheckBox eAboutDesign_Temp = (CheckBox)fvFixRequest.FindControl("eAboutDesign_Insert");
                string vAboutDesign = (eAboutDesign_Temp.Checked) ? "V" : "X";
                CheckBox eAboutHardDriver_Temp = (CheckBox)fvFixRequest.FindControl("eAboutHardDriver_Insert");
                string vAboutHardDriver = (eAboutHardDriver_Temp.Checked) ? "V" : "X";
                CheckBox eAboutSetting_Temp = (CheckBox)fvFixRequest.FindControl("eAboutSetting_Insert");
                string vAboutSetting = (eAboutSetting_Temp.Checked) ? "V" : "X";
                CheckBox eAboutSurver_Temp = (CheckBox)fvFixRequest.FindControl("eAboutSurver_Insert");
                string vAboutSurver = (eAboutSurver_Temp.Checked) ? "V" : "X";
                CheckBox eAboutPurchase_Temp = (CheckBox)fvFixRequest.FindControl("eAboutPurchase_Insert");
                string vAboutPurchase = (eAboutPurchase_Temp.Checked) ? "V" : "X";
                CheckBox eAboutOther_Temp = (CheckBox)fvFixRequest.FindControl("eAboutOther_Insert");
                string vAboutOthers = (eAboutOther_Temp.Checked) ? "V" : "X";
                TextBox eFixRemark_Temp = (TextBox)fvFixRequest.FindControl("eFixRemark_Insert");
                string vFixRemark = eFixRemark_Temp.Text.Trim();
                string vDisposal = "01";
                TextBox eDepNo_Temp = (TextBox)fvFixRequest.FindControl("eDepNo_Insert");
                string vDepNo_Temp = eDepNo_Temp.Text.Trim();


                sdsFixRequestMain.InsertCommand = "INSERT INTO SHEETA " + Environment.NewLine +
                                                  "       (sheetno, sheettype, day, buman, sheetclass, mode, UnitChief, FixType, aboutReport, aboutDataModify, aboutDesign, " + Environment.NewLine +
                                                  "        aboutHardDriver, aboutSetting, aboutSurver, aboutPurchase, aboutOthers, FixRemark, Disposal, depno) " + Environment.NewLine +
                                                  "VALUES (@SheetNo, @SheetType, @Day, @BuMan, @SheetClass, @Mode, @UnitChief, @FixType, @aboutReport, @aboutDataModify, " + Environment.NewLine +
                                                  "        @aboutDesign, @aboutHardDriver, @aboutSetting, @aboutSurver, @aboutPurchase, @aboutOthers, @FixRemark, @Disposal, @DepNo)";
                sdsFixRequestMain.InsertParameters.Clear();
                sdsFixRequestMain.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo.Trim()));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("SheetType", DbType.String, vSheetType.Trim()));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("Day", DbType.Date, (eDay_Temp.Text.Trim() != "") ? vDay.ToShortDateString() : String.Empty));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("BuMan", DbType.String, (vBuMan.Trim() != "") ? vBuMan.Trim() : String.Empty));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("SheetClass", DbType.String, vSheetClass.Trim()));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("Mode", DbType.String, (vMode.Trim() != "") ? vMode.Trim() : String.Empty));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("UnitChief", DbType.String, (vUnitChief.Trim() != "") ? vUnitChief.Trim() : String.Empty));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("FixType", DbType.String, (vFixType.Trim() != "") ? vFixType.Trim() : String.Empty));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("AboutReport", DbType.String, vAboutReport));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("AboutDataModify", DbType.String, vAboutDataModify));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("AboutDesign", DbType.String, vAboutDesign));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("AboutHardDriver", DbType.String, vAboutHardDriver));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("AboutSetting", DbType.String, vAboutSetting));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("AboutSurver", DbType.String, vAboutSurver));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("AboutPurchase", DbType.String, vAboutPurchase));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("AboutOthers", DbType.String, vAboutOthers));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("FixRemark", DbType.String, (vFixRemark.Trim() != "") ? vFixRemark.Trim() : String.Empty));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("Disposal", DbType.String, vDisposal));
                sdsFixRequestMain.InsertParameters.Add(new Parameter("DepNo", DbType.String, (vDepNo_Temp.Trim() != "") ? vDepNo_Temp.Trim() : String.Empty));

                sdsFixRequestMain.Insert();
                SendMail(vSheetNo);
                fvFixRequest.ChangeMode(FormViewMode.ReadOnly);
                gridShowData.DataBind();
                fvFixRequest.DataBind();
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