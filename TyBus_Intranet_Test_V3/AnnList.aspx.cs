using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class AnnList : System.Web.UI.Page
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

                    //事件日期
                    string vCaseDateURL = "InputDate.aspx?TextboxID=" + ePostDate_S_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    ePostDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + ePostDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    ePostDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eOpenDate_S_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eOpenDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eOpenDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eOpenDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        ePostDate_S_Search.Text = PF.GetMonthFirstDay(DateTime.Today.AddMonths(-1), "B");
                        ePostDate_E_Search.Text = PF.GetMonthLastDay(DateTime.Today, "B");
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

        private string GetSelStr()
        {
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ? "   and a.EmpNo = '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_PostDate = ((ePostDate_S_Search.Text.Trim() != "") && (ePostDate_E_Search.Text.Trim() != "")) ? "   and a.PostDate between '" + ePostDate_S_Search.Text.Trim() + "' and '" + ePostDate_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((ePostDate_S_Search.Text.Trim() != "") && (ePostDate_E_Search.Text.Trim() == "")) ? "   and a.PostDate = '" + ePostDate_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((ePostDate_S_Search.Text.Trim() == "") && (ePostDate_E_Search.Text.Trim() != "")) ? "   and a.PostDate = '" + ePostDate_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_OpenDate = ((eOpenDate_S_Search.Text.Trim() != "") && (eOpenDate_E_Search.Text.Trim() != "")) ? "   and a.EndDate >= '" + eOpenDate_S_Search.Text.Trim() + "' and a.StartDate <= '" + eOpenDate_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eOpenDate_S_Search.Text.Trim() != "") && (eOpenDate_E_Search.Text.Trim() == "")) ? "   and a.EndDate >= '" + eOpenDate_S_Search.Text.Trim() + "' and a.StartDate <= '" + eOpenDate_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eOpenDate_S_Search.Text.Trim() == "") && (eOpenDate_E_Search.Text.Trim() != "")) ? "   and a.EndDate >= '" + eOpenDate_E_Search.Text.Trim() + "' and a.StartDate <= '" + eOpenDate_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_PostTitle = (ePostTitle_Search.Text.Trim() != "") ? "   and a.PostTitle like '%" + ePostTitle_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Remark = (eRemark_Search.Text.Trim() != "") ? "   and a.Remark like '%" + eRemark_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vSelectStr = "select AnnNo, DepNo, (select [Name] from Department where DepNo = a.DepNo) DepName, " + Environment.NewLine +
                                "       EmpNo, (select [Name] from Employee where EmpNo = a.EmpNo) EmpName, " + Environment.NewLine +
                                "       PostDate, StartDate, EndDate, PostTitle " + Environment.NewLine +
                                "  from AnnList a " + Environment.NewLine +
                                " where (1 = 1) " + Environment.NewLine +
                                vWStr_DepNo + vWStr_EmpNo + vWStr_OpenDate + vWStr_PostDate + vWStr_PostTitle + vWStr_Remark +
                                " order by PostDate DESC, DepNo, EmpNo";
            return vSelectStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSelStr();
            sdsAnnList_List.SelectCommand = vSelectStr;
            gridAnnList.DataBind();
        }

        /// <summary>
        /// 發送公文上架通知訊息郵件
        /// </summary>
        /// <param name="fAnnDepName">發文單位</param>
        /// <param name="fAnnBuMan">發文人員</param>
        /// <param name="fAnnTitle">發文主旨</param>
        /// <param name="fAnnNote">發文內文</param>
        /// <param name="fFileName">附檔檔名</param>
        /// <param name="fFileSize">附檔大小</param>
        private void SendMail(string fAnnDepName, string fAnnBuMan, string fAnnTitle, string fAnnNote, string fFileName, int fFileSize)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DataTable dtToAdress = new DataTable();
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                string vSQLStr_Temp = "select Content from SysFlag where FormName = 'unSysflag' and ControlItem = 'a_WebCompanyName'";
                string vWebTitle = PF.GetValue(vConnStr, vSQLStr_Temp, "Content");

                vSQLStr_Temp = "select Email from Employee where LeaveDay is null and Title in ('010','020','030','080','090','100','110','120','130','140','142','170','180','190','200','230')";
                SqlDataAdapter daTemp = new SqlDataAdapter(vSQLStr_Temp, connTemp);
                connTemp.Open();
                daTemp.Fill(dtToAdress);
                if (dtToAdress.Rows.Count > 0)
                {
                    //有取回EMAIL地址才開始
                    using (MailMessage vMailMessage = new MailMessage())
                    {
                        vMailMessage.From = new MailAddress("webmaster@tybus.com.tw");
                        foreach (DataRow drTemp in dtToAdress.Rows)
                        {
                            if (drTemp[0].ToString().Trim() != "")
                            {
                                vMailMessage.To.Add(new MailAddress(drTemp[0].ToString()));
                            }
                        }
                        vMailMessage.Subject = "[" + vWebTitle + "] 公告張貼通知";
                        vMailMessage.Body = "發文單位：" + fAnnDepName.Trim() + Environment.NewLine +
                                            "發文人員：" + fAnnBuMan.Trim() + Environment.NewLine +
                                            "公告主旨：" + fAnnTitle.Trim() + Environment.NewLine +
                                            "內文：" + fAnnNote.Trim()
                                            ;
                        //2021.05.23 先拿掉檔案大小檢查的部份
                        //if ((fFileName.Trim() != "") && (fFileSize < 4100000))
                        if (fFileName.Trim() != "")
                        {
                            vMailMessage.Attachments.Add(new Attachment(fFileName));
                        }
                        SmtpClient vMailClient = new SmtpClient();
                        /* 測試改用 GMail
                        vMailClient.Host = "Mail.Tybus.com.tw";
                        vMailClient.Port = 25;
                        vMailClient.Credentials = new NetworkCredential("webmaster", "TyBusMis9999");
                        */
                        //2022.06.09 因應 GMail 政策改變，試看看改用 "應用程式密碼登入" 的方式操作 GMail
                        //vMailClient.Credentials = new NetworkCredential("tybusmis@gmail.com", "TyBusMis9999");
                        NetworkCredential cred = new NetworkCredential("tybusmis@gmail.com", "icyemoozikjooazh");
                        CredentialCache credCache = new CredentialCache();
                        vMailClient.Host = "smtp.gmail.com";
                        vMailClient.Port = 587;
                        credCache.Add(vMailClient.Host, vMailClient.Port, "login", cred);
                        vMailClient.EnableSsl = true;
                        vMailClient.Credentials = credCache;
                        try
                        {
                            vMailClient.Send(vMailMessage);
                        }
                        catch (Exception eMessage)
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('" + eMessage.Message + eMessage.ToString() + "')");
                            Response.Write("</" + "Script>");
                        }
                    }
                }
            }
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_Search.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            if ((vDepNo_Temp.Trim() != "") && (vDepName_Temp.Trim() != ""))
            {
                eDepNo_Search.Text = vDepNo_Temp.Trim();
                eDepName_Search.Text = vDepName_Temp.Trim();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到您輸入的單位，請重新輸入！')");
                Response.Write("</" + "Script>");
                eDepNo_Search.Focus();
            }
        }

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo_Temp = eEmpNo_Search.Text.Trim();
            string vEmpName_Temp = "";
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
            vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp;
                vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vEmpName_Temp + "' ";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            if ((vEmpNo_Temp.Trim() != "") && (vEmpName_Temp.Trim() != ""))
            {
                eEmpNo_Search.Text = vEmpNo_Temp.Trim();
                eEmpName_Search.Text = vEmpName_Temp.Trim();
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('找不到您輸入的員工，請重新輸入！')");
                Response.Write("</" + "Script>");
                eEmpNo_Search.Focus();
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvAnnList_Detail_DataBound(object sender, EventArgs e)
        {
            string vCaseDateURL = "";
            string vCaseDateScript = "";

            switch (fvAnnList_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Button bbDownLoad_List = (Button)fvAnnList_Detail.FindControl("bbDownLoad_List");
                    Label ePostFiles_List = (Label)fvAnnList_Detail.FindControl("ePostFiles_List");
                    if (bbDownLoad_List != null)
                    {
                        bbDownLoad_List.Visible = (ePostFiles_List.Text.Trim() != "");
                    }
                    Label ePostOpen_List = (Label)fvAnnList_Detail.FindControl("ePostOpen_List");
                    CheckBox cbPostOpen_List = (CheckBox)fvAnnList_Detail.FindControl("cbPostOpen_List");
                    if (ePostOpen_List != null)
                    {
                        cbPostOpen_List.Checked = (ePostOpen_List.Text.Trim().ToUpper() == "V");
                    }
                    Label eSendMail_List = (Label)fvAnnList_Detail.FindControl("eSendMail_List");
                    CheckBox cbSendMail_List = (CheckBox)fvAnnList_Detail.FindControl("cbSendMail_List");
                    if (eSendMail_List != null)
                    {
                        cbSendMail_List.Checked = (eSendMail_List.Text.Trim().ToUpper() == "V");
                    }
                    break;
                case FormViewMode.Edit:
                    TextBox ePostDate_Edit = (TextBox)fvAnnList_Detail.FindControl("ePostDate_Edit");
                    if (ePostDate_Edit != null)
                    {
                        vCaseDateURL = "InputDate.aspx?TextboxID=" + ePostDate_Edit.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        ePostDate_Edit.Attributes["onClick"] = vCaseDateScript;
                    }
                    TextBox eStartDate_Edit = (TextBox)fvAnnList_Detail.FindControl("eStartDate_Edit");
                    if (eStartDate_Edit != null)
                    {
                        vCaseDateURL = "InputDate.aspx?TextboxID=" + eStartDate_Edit.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eStartDate_Edit.Attributes["onClick"] = vCaseDateScript;
                    }
                    TextBox eEndDate_Edit = (TextBox)fvAnnList_Detail.FindControl("eEndDate_Edit");
                    if (eEndDate_Edit != null)
                    {
                        vCaseDateURL = "InputDate.aspx?TextboxID=" + eEndDate_Edit.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eEndDate_Edit.Attributes["onClick"] = vCaseDateScript;
                    }
                    Label eModifyMan_Edit = (Label)fvAnnList_Detail.FindControl("eModifyMan_Edit");
                    Label eModifyManName_Edit = (Label)fvAnnList_Detail.FindControl("eModifyManName_Edit");
                    if (eModifyMan_Edit != null)
                    {
                        eModifyMan_Edit.Text = vLoginID.Trim();
                        eModifyManName_Edit.Text = Session["LoginName"].ToString().Trim();
                    }
                    Label eModifyDate_Edit = (Label)fvAnnList_Detail.FindControl("eModifyDate_Edit");
                    if (eModifyDate_Edit != null)
                    {
                        eModifyDate_Edit.Text = DateTime.Today.ToString("yyyy/MM/dd");
                    }
                    TextBox ePostOpen_Edit = (TextBox)fvAnnList_Detail.FindControl("ePostOpen_Edit");
                    CheckBox cbPostOpen_Edit = (CheckBox)fvAnnList_Detail.FindControl("cbPostOpen_Edit");
                    if (ePostOpen_Edit != null)
                    {
                        cbPostOpen_Edit.Checked = (ePostOpen_Edit.Text.Trim().ToUpper() == "V");
                    }
                    TextBox eSendMail_Edit = (TextBox)fvAnnList_Detail.FindControl("eSendMail_Edit");
                    CheckBox cbSendMail_Edit = (CheckBox)fvAnnList_Detail.FindControl("cbSendMail_Edit");
                    if (eSendMail_Edit != null)
                    {
                        cbSendMail_Edit.Checked = (eSendMail_Edit.Text.Trim().ToUpper() == "V");
                    }
                    break;
                case FormViewMode.Insert:
                    TextBox ePostDate_INS = (TextBox)fvAnnList_Detail.FindControl("ePostDate_INS");
                    if (ePostDate_INS != null)
                    {
                        vCaseDateURL = "InputDate.aspx?TextboxID=" + ePostDate_INS.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        ePostDate_INS.Attributes["onClick"] = vCaseDateScript;
                        ePostDate_INS.Text = DateTime.Today.ToString("yyy/MM/dd");
                    }
                    TextBox eStartDate_INS = (TextBox)fvAnnList_Detail.FindControl("eStartDate_INS");
                    if (eStartDate_INS != null)
                    {
                        vCaseDateURL = "InputDate.aspx?TextboxID=" + eStartDate_INS.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eStartDate_INS.Attributes["onClick"] = vCaseDateScript;
                    }
                    TextBox eEndDate_INS = (TextBox)fvAnnList_Detail.FindControl("eEndDate_INS");
                    if (eEndDate_INS != null)
                    {
                        vCaseDateURL = "InputDate.aspx?TextboxID=" + eEndDate_INS.ClientID;
                        vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                        eEndDate_INS.Attributes["onClick"] = vCaseDateScript;
                    }
                    Label eBuMan_INS = (Label)fvAnnList_Detail.FindControl("eBuMan_INS");
                    Label eBuManName_INS = (Label)fvAnnList_Detail.FindControl("eBuManName_INS");
                    if (eBuMan_INS != null)
                    {
                        eBuMan_INS.Text = vLoginID.Trim();
                        eBuManName_INS.Text = Session["LoginName"].ToString().Trim();
                    }
                    Label eBuDate_INS = (Label)fvAnnList_Detail.FindControl("eBuDate_INS");
                    if (eBuDate_INS != null)
                    {
                        eBuDate_INS.Text = DateTime.Today.ToString("yyyy/MM/dd");
                    }
                    TextBox eDepNo_INS = (TextBox)fvAnnList_Detail.FindControl("eDepNo_INS");
                    Label eDepName_Temp = (Label)fvAnnList_Detail.FindControl("eDepName_INS");
                    if (eDepNo_INS != null)
                    {
                        eDepNo_INS.Text = Session["LoginDepNo"].ToString().Trim();
                        eDepName_Temp.Text = Session["LoginDepName"].ToString().Trim();
                    }
                    TextBox eEmpNo_INS = (TextBox)fvAnnList_Detail.FindControl("eEmpNo_INS");
                    Label eEmpName_INS = (Label)fvAnnList_Detail.FindControl("eEmpName_INS");
                    if (eEmpNo_INS != null)
                    {
                        eEmpNo_INS.Text = vLoginID;
                        eEmpName_INS.Text = Session["LoginName"].ToString().Trim();
                    }
                    break;
            }
        }

        protected void cbPostOpen_Edit_CheckedChanged(object sender, EventArgs e)
        {
            TextBox ePostOpen_Edit = (TextBox)fvAnnList_Detail.FindControl("ePostOpen_Edit");
            CheckBox cbPostOpen_Edit = (CheckBox)fvAnnList_Detail.FindControl("cbPostOPen_Edit");
            ePostOpen_Edit.Text = (cbPostOpen_Edit.Checked == true) ? "V" : "X";
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo_INS = (TextBox)fvAnnList_Detail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvAnnList_Detail.FindControl("eDepName_INS");
            string vDepNo_Temp = eDepNo_INS.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name").Trim();
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo").Trim();
            }
            eDepNo_INS.Text = vDepNo_Temp.Trim();
            eDepName_INS.Text = vDepName_Temp.Trim();
        }

        protected void eEmpNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eEmpNo_INS = (TextBox)fvAnnList_Detail.FindControl("eEmpNo_INS");
            Label eEmpName_INS = (Label)fvAnnList_Detail.FindControl("eEmpName_INS");
            string vEmpNo_Temp = eEmpNo_INS.Text.Trim();
            string vEmpName_Temp = "";
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp.Trim() + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp.Trim();
                vSQLStr_Temp = "select EmpNo from Employee where [Name] = '" + vEmpName_Temp.Trim() + "' ";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eEmpNo_INS.Text = vEmpNo_Temp.Trim();
            eEmpName_INS.Text = vEmpName_Temp.Trim();
        }

        protected void cbPostOpen_INS_CheckedChanged(object sender, EventArgs e)
        {
            TextBox ePostOpen_INS = (TextBox)fvAnnList_Detail.FindControl("ePostOpen_INS");
            CheckBox cbPostOpen_INS = (CheckBox)fvAnnList_Detail.FindControl("cbPostOPen_INS");
            ePostOpen_INS.Text = (cbPostOpen_INS.Checked == true) ? "V" : "X";
        }

        /// <summary>
        /// 刪除公告
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eAnnNo_List = (Label)fvAnnList_Detail.FindControl("eAnnNo_List");
            string vAnnNo_Temp = eAnnNo_List.Text.Trim();
            string vSQLStr_Del = "";
            if (vAnnNo_Temp != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    vSQLStr_Del = "select top 1 Items from AnnList_History where AnnNo = '" + vAnnNo_Temp.Trim() + "' order by Items DESC";
                    string vItems_Str = PF.GetValue(vConnStr, vSQLStr_Del, "Items");
                    int vNewItems = (vItems_Str != "") ? Int32.Parse(vItems_Str) + 1 : 1;
                    string vNewItems_Str = vNewItems.ToString("D4");
                    string vAnnNoItems = vAnnNo_Temp.Trim() + vNewItems_Str.Trim();
                    //寫入一筆異動資料
                    vSQLStr_Del = "INSERT INTO [dbo].[AnnList_History] " + Environment.NewLine +
                                  "       ([AnnNoItem], [Items], [AnnNo], [DepNo], [EmpNo], [PostDate], [StartDate], [EndDate], [PostTitle], [Remark], " + Environment.NewLine +
                                  "        [PostFiles], [BuMan], [BuDate], [ModifyMan], [ModifyDate], [PostOpen], [ModifyMode]) " + Environment.NewLine +
                                  "select '" + vAnnNoItems + "', '" + vNewItems_Str + "', [AnnNo], [DepNo], [EmpNo], [PostDate], [StartDate], [EndDate], [PostTitle], [Remark], " + Environment.NewLine +
                                  "       [PostFiles], [BuMan], [BuDate], [ModifyMan], [ModifyDate], [PostOpen], 'DEL' " + Environment.NewLine +
                                  "  from AnnList " + Environment.NewLine +
                                  " where AnnNo = '" + vAnnNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Del);
                    //刪除資料
                    sdsAnnList_Detail.DeleteParameters.Clear();
                    sdsAnnList_Detail.DeleteParameters.Add(new Parameter("AnnNo", DbType.String, vAnnNo_Temp));
                    sdsAnnList_Detail.Delete();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇要刪除的公告')");
                Response.Write("</" + "Script>");
            }
        }

        protected void sdsAnnList_Detail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridAnnList.DataBind();
                fvAnnList_Detail.DataBind();
            }
        }

        protected void sdsAnnList_Detail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridAnnList.DataBind();
                fvAnnList_Detail.DataBind();
            }
        }

        protected void sdsAnnList_Detail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridAnnList.DataBind();
                fvAnnList_Detail.DataBind();
            }
        }

        /// <summary>
        /// 按下新增
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            string vSQLStr_INS = "";
            string vUploadPath = "";
            string vUploadFileName = "";
            int vFileSize = 0;

            TextBox ePostDate_INS = (TextBox)fvAnnList_Detail.FindControl("ePostDate_INS");
            if (ePostDate_INS != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    DateTime vPostDate_INS = (ePostDate_INS.Text.Trim() != "") ? DateTime.Parse(ePostDate_INS.Text.Trim()) : DateTime.Today;
                    TextBox eStartDate_INS = (TextBox)fvAnnList_Detail.FindControl("eStartDate_INS");
                    TextBox eEndDate_INS = (TextBox)fvAnnList_Detail.FindControl("eEndDate_INS");
                    Label eBuDate_INS = (Label)fvAnnList_Detail.FindControl("eBuDate_INS");
                    string vYearMonth_Str = vPostDate_INS.ToString("yyyyMM");
                    vSQLStr_INS = "select top 1 AnnNo from AnnList where AnnNo like '" + vYearMonth_Str + "%' order by AnnNo DESC";
                    string vLastAnnNo = PF.GetValue(vConnStr, vSQLStr_INS, "AnnNo");
                    string vIndex_Str = vLastAnnNo.Replace(vYearMonth_Str, "");
                    int vNewIndex = (vIndex_Str != "") ? Int32.Parse(vIndex_Str) + 1 : 1;
                    string vNewAnnNo = vYearMonth_Str.Trim() + vNewIndex.ToString("D4");
                    TextBox eDepNo_INS = (TextBox)fvAnnList_Detail.FindControl("eDepNo_INS");
                    string vDepNo_Temp = (eDepNo_INS.Text.Trim() != "") ? eDepNo_INS.Text.Trim() : vLoginDepNo;
                    TextBox eEmpNo_INS = (TextBox)fvAnnList_Detail.FindControl("eEmpNo_INS");
                    string vEmpNo_Temp = (eEmpNo_INS.Text.Trim() != "") ? eEmpNo_INS.Text.Trim() : vLoginID;
                    TextBox ePostTitle_INS = (TextBox)fvAnnList_Detail.FindControl("ePostTitle_INS");
                    string vPostTitle_Temp = ePostTitle_INS.Text.Trim();
                    TextBox eRemark_INS = (TextBox)fvAnnList_Detail.FindControl("eRemark_INS");
                    string vRemark_Temp = eRemark_INS.Text.Trim();
                    TextBox ePostFiles_INS = (TextBox)fvAnnList_Detail.FindControl("ePostFiles_INS");
                    FileUpload fuPostFiles_INS = (FileUpload)fvAnnList_Detail.FindControl("fuPostFiles_INS");
                    string vPostFiles_Temp = "";
                    if ((fuPostFiles_INS != null) && (fuPostFiles_INS.FileName.Trim() != ""))
                    {
                        vFileSize = fuPostFiles_INS.PostedFile.ContentLength;
                        vPostFiles_Temp = vNewAnnNo + fuPostFiles_INS.FileName;
                        vUploadPath = @"\\172.20.3.17\公用暫存區\26.各項公告資料\";
                        vUploadFileName = vUploadPath + vNewAnnNo + fuPostFiles_INS.FileName;
                        fuPostFiles_INS.SaveAs(vUploadFileName);
                    }
                    TextBox ePostOpen_INS = (TextBox)fvAnnList_Detail.FindControl("ePostOpen_INS");
                    string vPostOpen_Temp = ePostOpen_INS.Text.Trim();
                    TextBox eSendMail_INS = (TextBox)fvAnnList_Detail.FindControl("eSendMail_INS");
                    string vSendMail_Temp = eSendMail_INS.Text.Trim();

                    vSQLStr_INS = "INSERT INTO AnnList " + Environment.NewLine +
                                  "       (AnnNo, DepNo, EmpNo, PostDate, StartDate, EndDate, PostTitle, Remark, PostFiles, BuMan, BuDate, PostOpen, SendMail) " + Environment.NewLine +
                                  "VALUES (@AnnNo, @DepNo, @EmpNo, @PostDate, @StartDate, @EndDate, @PostTitle, @Remark, @PostFiles, @BuMan, @BuDate, @PostOpen, @SendMail)";
                    sdsAnnList_Detail.InsertCommand = vSQLStr_INS;
                    sdsAnnList_Detail.InsertParameters.Clear();
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("AnnNo", DbType.String, vNewAnnNo));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, vDepNo_Temp));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("EmpNo", DbType.String, vEmpNo_Temp));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("PostDate", DbType.Date, vPostDate_INS.ToShortDateString()));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("StartDate", DbType.Date, (eStartDate_INS.Text.Trim() != "") ? eStartDate_INS.Text.Trim() : String.Empty));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("EndDate", DbType.Date, (eEndDate_INS.Text.Trim() != "") ? eEndDate_INS.Text.Trim() : String.Empty));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("PostTitle", DbType.String, (vPostTitle_Temp.Trim() != "") ? vPostTitle_Temp.Trim() : String.Empty));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark_Temp.Trim() != "") ? vRemark_Temp.Trim() : String.Empty));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("PostFiles", DbType.String, (vPostFiles_Temp.Trim() != "") ? vPostFiles_Temp.Trim() : String.Empty));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.Today.ToShortDateString()));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("PostOpen", DbType.String, (vPostOpen_Temp.Trim() != "") ? vPostOpen_Temp.Trim() : String.Empty));
                    sdsAnnList_Detail.InsertParameters.Add(new Parameter("SendMail", DbType.String, (vSendMail_Temp.Trim() != "") ? vSendMail_Temp.Trim() : String.Empty));

                    sdsAnnList_Detail.Insert();
                    Label eDepName_INS = (Label)fvAnnList_Detail.FindControl("eDepName_INS");
                    string vDepName_Temp = eDepName_INS.Text.Trim();
                    Label eEmpName_INS = (Label)fvAnnList_Detail.FindControl("eEmpName_INS");
                    string vEmpName_Temp = eEmpName_INS.Text.Trim();
                    CheckBox cbSendMail_INS = (CheckBox)fvAnnList_Detail.FindControl("cbSendMail_INS");
                    if (vSendMail_Temp.Trim().ToUpper() == "V")
                    {
                        SendMail(vDepName_Temp.Trim(), vEmpName_Temp.Trim(), vPostTitle_Temp.Trim(), vRemark_Temp.Trim(), vUploadFileName, vFileSize);
                    }

                    fvAnnList_Detail.ChangeMode(FormViewMode.ReadOnly);
                    gridAnnList.DataBind();
                    fvAnnList_Detail.DataBind();
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

        /// <summary>
        /// 資料修改存檔前
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void sdsAnnList_Detail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            Label eAnnNo_Edit = (Label)fvAnnList_Detail.FindControl("eAnnNo_Edit");
            if (eAnnNo_Edit.Text.Trim() != "")
            {
                string vAnnNo_Edit = eAnnNo_Edit.Text.Trim();
                TextBox ePostDate_Edit = (TextBox)fvAnnList_Detail.FindControl("ePostDate_Edit");
                TextBox eStartDate_Edit = (TextBox)fvAnnList_Detail.FindControl("eStartDate_Edit");
                TextBox eEndDate_Edit = (TextBox)fvAnnList_Detail.FindControl("eEndDate_Edit");
                Label eModifyDate_Edit = (Label)fvAnnList_Detail.FindControl("eModifyDate_Edit");
                e.Command.Parameters["@PostDate"].Value = (ePostDate_Edit.Text.Trim() != "") ? DateTime.Parse(ePostDate_Edit.Text.Trim()) : DateTime.Today;
                if (eStartDate_Edit.Text.Trim() != "")
                {
                    e.Command.Parameters["@StartDate"].Value = DateTime.Parse(eStartDate_Edit.Text.Trim());
                }
                if (eEndDate_Edit.Text.Trim() != "")
                {
                    e.Command.Parameters["@EndDate"].Value = DateTime.Parse(eEndDate_Edit.Text.Trim());
                }
                e.Command.Parameters["@ModifyMan"].Value = (eModifyDate_Edit.Text.Trim() != "") ? DateTime.Parse(eModifyDate_Edit.Text.Trim()) : DateTime.Today;
                TextBox ePostFiles_Edit = (TextBox)fvAnnList_Detail.FindControl("ePostFiles_Edit");
                FileUpload fuPostFiles_Edit = (FileUpload)fvAnnList_Detail.FindControl("fuPostFiles_Edit");
                if ((fuPostFiles_Edit != null) && (fuPostFiles_Edit.FileName.Trim() != ""))
                {
                    e.Command.Parameters["@PostFiles"].Value = vAnnNo_Edit + fuPostFiles_Edit.FileName;
                    string vUploadPath = @"\\172.20.3.17\公用暫存區\26.各項公告資料\";
                    string vUploadFileName = vUploadPath + vAnnNo_Edit + fuPostFiles_Edit.FileName;
                    fuPostFiles_Edit.SaveAs(vUploadFileName);
                }
            }
        }

        protected void bbDownLoad_List_Click(object sender, EventArgs e)
        {
            Label ePostFiles_List = (Label)fvAnnList_Detail.FindControl("ePostFiles_List");
            if (ePostFiles_List.Text.Trim() != "")
            {
                string vDownLoadFiles = @"\\172.20.3.17\公用暫存區\26.各項公告資料\" + ePostFiles_List.Text.Trim();
                string vFileName = ePostFiles_List.Text.Trim();
                //============================================================================================================================
                string vRecordNote = "下載檔案：公告管理" + Environment.NewLine +
                                     "AnnList.aspx" + Environment.NewLine +
                                     "檔案名稱：" + vFileName;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                //============================================================================================================================
                FileStream fsTemp = new FileStream(vDownLoadFiles, FileMode.Open);
                byte[] bytes_Temp = new byte[(int)fsTemp.Length];
                fsTemp.Read(bytes_Temp, 0, bytes_Temp.Length);
                fsTemp.Close();
                Response.ContentType = "application/octet-stream";
                HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                string TourVerision = brObject.Type;
                if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                {
                    HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName, System.Text.Encoding.UTF8));
                }
                else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                {
                    // 設定強制下載標頭
                    Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName));
                }
                // 輸出檔案
                Response.BinaryWrite(bytes_Temp);
                Response.Flush();
                Response.End();
            }
        }

        protected void cbSendMail_Edit_CheckedChanged(object sender, EventArgs e)
        {
            TextBox eSendMail_Edit = (TextBox)fvAnnList_Detail.FindControl("eSendMail_Edit");
            CheckBox cbSendMail_Edit = (CheckBox)fvAnnList_Detail.FindControl("cbSendMail_Edit");
            eSendMail_Edit.Text = (cbSendMail_Edit.Checked == true) ? "V" : "X";
        }

        protected void cbSendMail_INS_CheckedChanged(object sender, EventArgs e)
        {
            TextBox eSendMail_INS = (TextBox)fvAnnList_Detail.FindControl("eSendMail_INS");
            CheckBox cbSendMail_INS = (CheckBox)fvAnnList_Detail.FindControl("cbSendMail_INS");
            eSendMail_INS.Text = (cbSendMail_INS.Checked == true) ? "V" : "X";
        }

        protected void gridAnnList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridAnnList.PageIndex = e.NewPageIndex;
            gridAnnList.DataBind();
        }
    }
}