using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class OfficialDocument : System.Web.UI.Page
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
        private DateTime vToday = DateTime.Today;

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

                string vSQLStr = "";

                if (vLoginID != "")
                {
                    string vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_Start_Search.ClientID;
                    string vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eDocDate_Start_Search.Attributes["onClick"] = vDocDateScript;
                    vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_End_Search.ClientID;
                    vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eDocDate_End_Search.Attributes["onClick"] = vDocDateScript;
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }

                    if (vLoginDepNo == "")
                    {
                        vSQLStr = "select DepNo from Employee where EmpNo = '" + vLoginID + "' and LeaveDay is null ";
                        vLoginDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
                    }
                    if (!IsPostBack)
                    {
                        vSQLStr = "select count(ControlName) RCount from WebPermissionB where EmpNo = '" + vLoginID + "' and AllowPermission = 1 and ControlName = 'EditOfficialDocument' ";
                        int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));

                        eDocDep_Start_Search.Text = vLoginDepNo;
                        eDocDep_End_Search.Text = "";
                        eDocDep_Start_Search.Enabled = (vRCount > 0);
                        eDocDep_End_Search.Enabled = (vRCount > 0);
                        eDocDate_Start_Search.Text = PF.GetMonthFirstDay(vToday, "C");
                        eDocDate_End_Search.Text = PF.GetMonthLastDay(vToday, "C");
                        eUndertaker_Start_Search.Text = vLoginID;
                    }
                    DocumentDataBind();
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

        private void DocumentDataBind()
        {
            string vSQLStr = "";
            if (vLoginID != "")
            {
                string vWStr_DocDepNo = ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() != "")) ? " and DocDep between '" + eDocDep_Start_Search.Text.Trim() + "' and '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                        ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() == "")) ? " and DocDep = '" + eDocDep_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                        ((eDocDep_Start_Search.Text.Trim() == "") && (eDocDep_End_Search.Text.Trim() != "")) ? " and DocDep = '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine : "";

                string vWStr_DocDate = ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() != "")) ?
                                       " and DocDate between '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() == "")) ?
                                       " and DocDate = '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eDocDate_Start_Search.Text.Trim() == "") && (eDocDate_End_Search.Text.Trim() != "")) ?
                                       " and DocDate = '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";

                string vWStr_DocType = (eDocType_Search.Text.Trim() != "") ? " and DocType = '" + eDocType_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_Undertaker = ((eUndertaker_Start_Search.Text.Trim() != "") && (eUndertaker_End_Search.Text.Trim() != "")) ? " and Undertaker between '" + eUndertaker_Start_Search.Text.Trim() + "' and '" + eUndertaker_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                          ((eUndertaker_Start_Search.Text.Trim() != "") && (eUndertaker_End_Search.Text.Trim() == "")) ? " and Undertaker = '" + eUndertaker_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                          ((eUndertaker_Start_Search.Text.Trim() == "") && (eUndertaker_End_Search.Text.Trim() != "")) ? " and Undertaker = '" + eUndertaker_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_DocTitle = (eDocTitle_Search.Text.Trim() != "") ? " and DocTitle like '%" + eDocTitle_Search.Text.Trim() + "%' " : "";
                string vWStr_DocNo = ((eDocNo_S_Search.Text.Trim() != "") && (eDocNo_E_Search.Text.Trim() != "")) ? "   and DocNo between '" + Int32.Parse(eDocNo_S_Search.Text.Trim()).ToString("D4") + "' and '" + Int32.Parse(eDocNo_E_Search.Text.Trim()).ToString("D4") + "' " + Environment.NewLine :
                                     ((eDocNo_S_Search.Text.Trim() != "") && (eDocNo_E_Search.Text.Trim() == "")) ? "   and DocNo = '" + Int32.Parse(eDocNo_S_Search.Text.Trim()).ToString("D4") + "' " + Environment.NewLine :
                                     ((eDocNo_S_Search.Text.Trim() == "") && (eDocNo_E_Search.Text.Trim() != "")) ? "   and DocNo = '" + Int32.Parse(eDocNo_E_Search.Text.Trim()).ToString("D4") + "' " + Environment.NewLine : "";
                string vWStr_StoreDate = "";
                string vWStr_DocYears = (eDocYears_Search.Text.Trim() != "") ? "   and DocYears = '" + Int32.Parse(eDocYears_Search.Text.Trim()).ToString("D3") + "' " + Environment.NewLine : "";
                /*
                string vWStr_FirstWord = (vLoginID == "731520") ? " and DocFirstWord = 'FW0204' " + Environment.NewLine : //只能看到人字的公文
                                         (vLoginID == "020219") ? " and DocFirstWord in ('FW0201', 'FW0203', 'FW0205', 'FW0206')" + Environment.NewLine : //只能看到股，總，福，財四種公文
                                         ((vLoginID == "740112") || (vLoginID == "050745")) ? " and DocFirstWord = 'FW0202' " + Environment.NewLine : //只能看到勞字的公文，2020.04.01 新增一個 050745
                                         (vLoginID == "080895") ? " and DocFirstWord = 'FW0207' " + Environment.NewLine ://只能看到職字的公文
                                         ((vLoginID == "998968") || (vLoginID == "057560") || (vLoginID == "948173") || (vLoginID == "080893")) ? " and DocFirstWord = 'FW0201' " + Environment.NewLine : //只能看到總字的公文
                                         ""; //*/
                //2021.05.17 林玉蘭 (731520) 調回人事股，江筱書 (938122) 調三峽站
                /* 2021.06.16 修改公文系統權限取得方式
                string vWStr_FirstWord = //(vLoginID == "938122") ? " and DocFirstWord = 'FW0204' " + Environment.NewLine : //只能看到人字的公文
                                         (vLoginID == "731520") ? " and DocFirstWord = 'FW0204' " + Environment.NewLine : //只能看到人字的公文
                                         (vLoginID == "020219") ? " and DocFirstWord in ('FW0201', 'FW0203', 'FW0205', 'FW0206')" + Environment.NewLine : //只能看到股，總，福，財四種公文
                                         ((vLoginID == "740112") || (vLoginID == "050745")) ? " and DocFirstWord = 'FW0202' " + Environment.NewLine : //只能看到勞字的公文，2020.04.01 新增一個 050745
                                         //(vLoginID == "080895") ? " and DocFirstWord = 'FW0207' " + Environment.NewLine ://只能看到職字的公文
                                         (vLoginID == "100012") ? " and DocFirstWord = 'FW0207' " + Environment.NewLine ://只能看到職字的公文
                                         //2021.05.07 修正總務課人員...原工號 948173 王俊傑置換成 756320 李碧華
                                         //((vLoginID == "998968") || (vLoginID == "057560") || (vLoginID == "948173") || (vLoginID == "080893")) ? " and DocFirstWord = 'FW0201' " + Environment.NewLine : //只能看到總字的公文
                                         ((vLoginID == "998968") || (vLoginID == "057560") || (vLoginID == "756320") || (vLoginID == "080893")) ? " and DocFirstWord = 'FW0201' " + Environment.NewLine : //只能看到總字的公文
                                         ""; //*/
                string vWStr_FirstWord = "";
                string vTempStr = "";
                using (SqlConnection connGetFirstWord = new SqlConnection(vConnStr))
                {
                    vSQLStr = "select FirstCodeNo from FirstWordPermission where EmpNo = '" + vLoginID + "' and IsUsed = 1";
                    SqlCommand cmdGetFirstWord = new SqlCommand(vSQLStr, connGetFirstWord);
                    connGetFirstWord.Open();
                    SqlDataReader drGetFirstWord = cmdGetFirstWord.ExecuteReader();
                    if (drGetFirstWord.HasRows)
                    {
                        while (drGetFirstWord.Read())
                        {
                            vTempStr = (vTempStr == "") ? "'" + drGetFirstWord[0].ToString().Trim() + "'" : vTempStr + ", '" + drGetFirstWord[0].ToString().Trim() + "'";
                        }
                        vWStr_FirstWord = " and DocFirstWord in (" + vTempStr + ") " + Environment.NewLine;
                    }
                    else
                    {
                        vWStr_FirstWord = "";
                    }
                }
                //=================================================================================================================================================================
                switch (rbIsStored.SelectedValue)
                {
                    default:
                        vWStr_StoreDate = "";
                        break;
                    case "1":
                        vWStr_StoreDate = " and StoreDate is not null " + Environment.NewLine;
                        break;
                    case "2":
                        vWStr_StoreDate = " and StoreDate is null " + Environment.NewLine;
                        break;
                }

                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSelectStr = "";
                vSQLStr = "select count(Items) RCount from WebPermissionB where ControlName = 'EditOfficialDocument' and AllowPermission = 1 and EmpNo = '" + vLoginID + "' ";
                int vHasEditPermission = Int32.Parse(PF.GetValue(vConnStr, vSQLStr, "RCount"));
                if (vHasEditPermission > 0)
                {
                    //有被授與修改公文登錄權限的 "個人" (全單位給予權限者不包括在內)
                    vSelectStr = "select DocIndex, DocDate, DocYears, " + Environment.NewLine +
                                 "       DocDep, (select [Name] from Department where DepNo = a.DocDep) DocDep_C, " + Environment.NewLine +
                                 "       DocFirstWord, (SELECT DocFirstCWord FROM DOCFirstWord WHERE (FWNo = a.DocFirstWord)) AS DocFirstWord_C, " + Environment.NewLine +
                                 "       DocNo, DocSourceUnit, " + Environment.NewLine +
                                 "       DocType, (select ClassTxt from DBDICB where FKey = (cast('公文收發登記' as char(16)) + cast('OfficialDocument' as char(16)) + 'DocType') and ClassNo = a.DocType) DocType_C, " +
                                 "       DocTitle, " + Environment.NewLine +
                                 "       Undertaker, (select [Name] from Employee where EmpNo = a.Undertaker) Undertaker_C, " + Environment.NewLine +
                                 "       OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, " + Environment.NewLine +
                                 "       BuildMan, (select [Name] from Employee where EmpNo = a.BuildMan) BuildMan_C, " + Environment.NewLine +
                                 "       BuildDate, Remark, StoreDate, " +
                                 "       StoreMan, (select [Name] from Employee where EmpNo = a.StoreMan) StoreMan_C, " +
                                 "       Remark_Store, IsHide " + Environment.NewLine +
                                 "  from OfficialDocument a " + Environment.NewLine +
                                 " where (1 = 1) " + Environment.NewLine +
                                 vWStr_DocDepNo + vWStr_DocDate + vWStr_DocType + vWStr_Undertaker + vWStr_DocTitle + vWStr_StoreDate + vWStr_FirstWord + vWStr_DocNo + vWStr_DocYears +
                                 " order by a.DocDate DESC, a.DocDep, a.DocNo DESC";
                }
                else
                {
                    vSelectStr = "select DocIndex, DocDate, DocYears, " + Environment.NewLine +
                                 "       DocDep, (select [Name] from Department where DepNo = a.DocDep) DocDep_C, " + Environment.NewLine +
                                 "       DocFirstWord, (SELECT DocFirstCWord FROM DOCFirstWord WHERE (FWNo = a.DocFirstWord)) AS DocFirstWord_C, " + Environment.NewLine +
                                 "       DocNo, DocSourceUnit, " + Environment.NewLine +
                                 "       DocType, (select ClassTxt from DBDICB where FKey = (cast('公文收發登記' as char(16)) + cast('OfficialDocument' as char(16)) + 'DocType') and ClassNo = a.DocType) DocType_C, " +
                                 "       DocTitle, " + Environment.NewLine +
                                 "       Undertaker, (select [Name] from Employee where EmpNo = a.Undertaker) Undertaker_C, " + Environment.NewLine +
                                 "       OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, " + Environment.NewLine +
                                 "       BuildMan, (select [Name] from Employee where EmpNo = a.BuildMan) BuildMan_C, " + Environment.NewLine +
                                 "       BuildDate, Remark, StoreDate, " +
                                 "       StoreMan, (select [Name] from Employee where EmpNo = a.StoreMan) StoreMan_C, " +
                                 "       Remark_Store, IsHide " + Environment.NewLine +
                                 "  from ( " + Environment.NewLine +
                                 "        select DocIndex, DocDate, DocYears, DocDep, DocFirstWord, DocNo, DocSourceUnit, DocType, DocTitle, Undertaker, " + Environment.NewLine +
                                 "               OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, BuildMan, BuildDate, Remark, StoreDate, " + Environment.NewLine +
                                 "               StoreMan, Remark_Store, IsHide " + Environment.NewLine +
                                 "          from OfficialDocument " + Environment.NewLine +
                                 "         where IsHide = 0 " + Environment.NewLine +
                                 "         union all " + Environment.NewLine +
                                 "        select DocIndex, DocDate, DocYears, DocDep, DocFirstWord, DocNo, DocSourceUnit, DocType, DocTitle, Undertaker, " + Environment.NewLine +
                                 "               OutsideDocFirstWord, OutsideDocNo, Attachement, Implementation, BuildMan, BuildDate, Remark, StoreDate, " + Environment.NewLine +
                                 "               StoreMan, Remark_Store, IsHide " + Environment.NewLine +
                                 "          from OfficialDocument " + Environment.NewLine +
                                 "         where IsHide = 1 and Undertaker = '" + vLoginID + "' " + Environment.NewLine +
                                 "        ) a " + Environment.NewLine +
                                 " where (1 = 1) " + Environment.NewLine +
                                 vWStr_DocDepNo + vWStr_DocDate + vWStr_DocType + vWStr_Undertaker + vWStr_DocTitle + vWStr_StoreDate + vWStr_FirstWord + vWStr_DocNo + vWStr_DocYears +
                                 " order by a.DocDate DESC, a.DocDep, a.DocNo DESC";
                }
                sdsOfficialDocument_List.SelectCommand = "";
                sdsOfficialDocument_List.SelectCommand = vSelectStr;
                gridDataList.DataBind();
            }
        }

        private void BeginReadOnlyMode()
        {
            string vSQLStr = "";
            Button bbTemp = (Button)fvOfficialDocumentDetail.FindControl("bbStored_List");
            Label lbStoreDate_List = (Label)fvOfficialDocumentDetail.FindControl("eStoreDate_List");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (vLoginDepNo == "")
            {
                vSQLStr = "select DepNo from Employee where EmpNo = '" + vLoginID + "' ";
                vLoginDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            if (bbTemp != null)
            {
                bbTemp.Visible = ((vLoginDepNo == "02") && (lbStoreDate_List.Text.Trim() == ""));
            }
        }

        private void BeginUpdateMode()
        {
            string vSQLStr = "";
            TextBox eDocDate_Edit = (TextBox)fvOfficialDocumentDetail.FindControl("eDocDate_Edit");
            TextBox eStoreDate_Edit = (TextBox)fvOfficialDocumentDetail.FindControl("eStoreDate_Edit");
            Label eStoreMan_Edit = (Label)fvOfficialDocumentDetail.FindControl("eStoreMan_Edit");
            Label eStoreMan_C_Edit = (Label)fvOfficialDocumentDetail.FindControl("eStoreMan_C_Edit");
            TextBox eRemark_Store_Edit = (TextBox)fvOfficialDocumentDetail.FindControl("eRemark_Store_Edit");

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (vLoginDepNo == "")
            {
                vSQLStr = "select DepNo from Employee where EmpNo = '" + vLoginID + "' ";
                vLoginDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            string vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_Edit.ClientID;
            string vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
            eDocDate_Edit.Attributes["onClick"] = vDocDateScript;
        }

        private void BeginInsertMode()
        {
            string vSQLStr = "";
            TextBox eDocDate_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eDocDate_Insert");
            Label eDocDep_INS = (Label)fvOfficialDocumentDetail.FindControl("eDocDep_Insert");
            Label eDocDep_C_INS = (Label)fvOfficialDocumentDetail.FindControl("eDocDep_C_Insert");
            TextBox eUndertaker_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eUndertaker_Insert");
            Label eUndertaker_C_INS = (Label)fvOfficialDocumentDetail.FindControl("eUndertaker_C_Insert");
            DropDownList ddlDocFirstWord_INS = (DropDownList)fvOfficialDocumentDetail.FindControl("ddlDocFirstWord_Insert");
            Label eDocFirstWord_INS = (Label)fvOfficialDocumentDetail.FindControl("eDocFirstWord_Insert");
            Label eBuildMan_INS = (Label)fvOfficialDocumentDetail.FindControl("eBuildMan_Insert");
            Label eBuildMan_C_INS = (Label)fvOfficialDocumentDetail.FindControl("eBuildMan_C_Insert");
            Label eBuildDate_INS = (Label)fvOfficialDocumentDetail.FindControl("eBuildDate_Insert");

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (vLoginDepNo == "")
            {
                vSQLStr = "select DepNo from Employee where EmpNo = '" + vLoginID + "' ";
                vLoginDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            string vSelectStr = "select FWNo, DocFirstCWord from DocFirstWord where DepNo = '" + vLoginDepNo + "' order by FWNo ";
            //2021.06.16 修改取得公文字號方式
            string vFW_PermissionStr = "select FirstCodeNo as FWNo, (select DocFirstCWord from DocFirstWord where FWNo = FirstWordPermission.FirstCodeNo) as DocFirstCWord " + Environment.NewLine +
                                       "  from FirstWordPermission " + Environment.NewLine +
                                       " where EmpNo = '" + vLoginID + "' and IsUsed = 1";
            using (SqlConnection connGetFW = new SqlConnection(vConnStr))
            {
                SqlCommand cmdGetFW = new SqlCommand(vFW_PermissionStr, connGetFW);
                connGetFW.Open();
                SqlDataReader drGetFW = cmdGetFW.ExecuteReader();
                if (drGetFW.HasRows)
                {
                    while (drGetFW.Read())
                    {
                        ListItem liTemp = new ListItem(drGetFW["DocFirstCWord"].ToString().Trim(), drGetFW["FWNo"].ToString().Trim());
                        ddlDocFirstWord_INS.Items.Add(liTemp);
                    }
                }
                else
                {
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSelectStr, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        while (drTemp.Read())
                        {
                            ListItem liTemp = new ListItem(drTemp["DocFirstCWord"].ToString().Trim(), drTemp["FWNo"].ToString().Trim());
                            ddlDocFirstWord_INS.Items.Add(liTemp);
                        }
                    }
                }
                /*
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["DocFirstCWord"].ToString().Trim(), drTemp["FWNo"].ToString().Trim());
                        if ((vLoginDepNo != "02") ||
                            ((vLoginDepNo == "02") && ((vLoginID == "938236"))) ||
                            //2021.05.17 林玉蘭 (731520) 調回人事股，江筱書 (938122) 調三峽站
                            ((vLoginID == "731520") && (drTemp["FWNo"].ToString().Trim() == "FW0204")) ||
                            //((vLoginID == "938122") && (drTemp["FWNo"].ToString().Trim() == "FW0204")) ||
                            (vLoginID == "020219") && ((drTemp["FWNo"].ToString().Trim() == "FW0201") || (drTemp["FWNo"].ToString().Trim() == "FW0203") || (drTemp["FWNo"].ToString().Trim() == "FW0205") || (drTemp["FWNo"].ToString().Trim() == "FW0206")) ||
                            (((vLoginID == "740112") || (vLoginID == "050745")) && (drTemp["FWNo"].ToString().Trim() == "FW0202")) || //2020.04.01 新增一名 050745
                            ((vLoginID == "998968") && (drTemp["FWNo"].ToString().Trim() == "FW0201")) ||
                            ((vLoginID == "080893") && (drTemp["FWNo"].ToString().Trim() == "FW0201")) ||
                            ((vLoginID == "057560") && (drTemp["FWNo"].ToString().Trim() == "FW0201")) ||
                            //((vLoginID == "080895") && (drTemp["FWNo"].ToString().Trim() == "FW0207")) ||
                            ((vLoginID == "100012") && (drTemp["FWNo"].ToString().Trim() == "FW0207")) ||
                            //2021.05.07 修正總務課人員...原工號 948173 王俊傑置換成 756320 李碧華
                            //((vLoginID == "948173") && (drTemp["FWNo"].ToString().Trim() == "FW0201")))
                            ((vLoginID == "756320") && (drTemp["FWNo"].ToString().Trim() == "FW0201")))
                        {
                            ddlDocFirstWord_INS.Items.Add(liTemp);
                        }
                    }
                } //*/
            }
            ddlDocFirstWord_INS.SelectedIndex = 0;
            eDocFirstWord_INS.Text = ddlDocFirstWord_INS.Items[0].Value.ToString().Trim();
            eDocDate_INS.Text = DateTime.Today.ToString("yyyy/MM/dd");
            eDocDep_INS.Text = vLoginDepNo;
            vSQLStr = "select [Name] from Department where DepNo = '" + vLoginDepNo + "' ";
            eDocDep_C_INS.Text = PF.GetValue(vConnStr, vSQLStr, "Name");
            eUndertaker_INS.Text = vLoginID;
            vSQLStr = "select [Name] from Employee where EmpNo = '" + vLoginID + "' ";
            eUndertaker_C_INS.Text = PF.GetValue(vConnStr, vSQLStr, "Name");
            eBuildMan_INS.Text = vLoginID;
            eBuildMan_C_INS.Text = PF.GetValue(vConnStr, vSQLStr, "Name");
            eBuildDate_INS.Text = PF.TransDateString(DateTime.Today, "C");
            string vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_INS.ClientID;
            string vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
            eDocDate_INS.Attributes["onClick"] = vDocDateScript;
        }

        protected void eDocDep_Start_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDocDep_Start_Search.Text.Trim();
            string vDepName_Temp = "";

            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDocDep_Start_Search.Text = vDepNo_Temp.Trim();
            eDocDepName_Start_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDocDep_End_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDocDep_End_Search.Text.Trim();
            string vDepName_Temp = "";

            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDocDep_End_Search.Text = vDepNo_Temp.Trim();
            eDocDepName_End_Search.Text = vDepName_Temp.Trim();
        }

        protected void ddlDocType_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDocType_Search.Text = ddlDocType_Search.SelectedValue.Trim();
        }

        protected void eUndertaker_Start_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo = eUndertaker_Start_Search.Text.Trim();
            string vEmpName = "";

            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' and LeaveDay is null ";
            vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName + "' and LeaveDay is null ";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eUndertaker_Start_Search.Text = vEmpNo.Trim();
            eUndertakerName_Start_Search.Text = vEmpName.Trim();
        }

        protected void eUndertaker_End_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo = eUndertaker_End_Search.Text.Trim();
            string vEmpName = "";

            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' and LeaveDay is null ";
            vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName + "' and LeaveDay is null ";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eUndertaker_End_Search.Text = vEmpNo.Trim();
            eUndertakerName_End_Search.Text = vEmpName.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            DocumentDataBind();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvOfficialDocumentDetail_DataBound(object sender, EventArgs e)
        {
            switch (fvOfficialDocumentDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    BeginReadOnlyMode();
                    break;
                case FormViewMode.Edit:
                    BeginUpdateMode();
                    break;
                case FormViewMode.Insert:
                    BeginInsertMode();
                    break;
            }
        }

        protected void ddlDocType_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTemp = (DropDownList)fvOfficialDocumentDetail.FindControl("ddlDocType_Edit");
            Label lbTemp = (Label)fvOfficialDocumentDetail.FindControl("eDocType_Edit");
            lbTemp.Text = ddlTemp.SelectedValue.Trim();
        }

        protected void eUndertaker_Edit_TextChanged(object sender, EventArgs e)
        {
            string vSQLStr = "";
            TextBox eTemp = (TextBox)fvOfficialDocumentDetail.FindControl("eUndertaker_Edit");
            Label lbTemp = (Label)fvOfficialDocumentDetail.FindControl("eUndertaker_C_Edit");
            string vUndertaker = eTemp.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select [Name] from Employee where EmpNo = '" + vUndertaker + "' and LeaveDay is null ";
            string vUndertakerName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vUndertakerName == "")
            {
                vUndertakerName = vUndertaker;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vUndertakerName + "' and LeaveDay is null ";
                vUndertaker = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eTemp.Text = vUndertaker.Trim();
            lbTemp.Text = vUndertakerName.Trim();
        }

        protected void ddlDocFirstWord_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            Label lbDocFirstWord = (Label)fvOfficialDocumentDetail.FindControl("eDocFirstWord_Insert");
            DropDownList ddlDocFirstWord = (DropDownList)fvOfficialDocumentDetail.FindControl("ddlDocFirstWord_Insert");
            lbDocFirstWord.Text = ddlDocFirstWord.SelectedValue.ToString().Trim();
        }

        protected void ddlDocType_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTemp = (DropDownList)fvOfficialDocumentDetail.FindControl("ddlDocType_Insert");
            Label lbTemp = (Label)fvOfficialDocumentDetail.FindControl("eDocType_Insert");
            lbTemp.Text = ddlTemp.SelectedValue.Trim();
        }

        protected void eUndertaker_Insert_TextChanged(object sender, EventArgs e)
        {
            string vSQLStr = "";
            TextBox eTemp = (TextBox)fvOfficialDocumentDetail.FindControl("eUndertaker_Insert");
            Label lbTemp = (Label)fvOfficialDocumentDetail.FindControl("eUndertaker_C_Insert");
            Label lbDepNo_Temp = (Label)fvOfficialDocumentDetail.FindControl("eDocDep_Insert");
            Label lbDepName_Temp = (Label)fvOfficialDocumentDetail.FindControl("eDocDep_C_Insert");
            Label lbUndertakerError_Temp = (Label)fvOfficialDocumentDetail.FindControl("lbUndertakerError_Insert");

            string vUndertaker = eTemp.Text.Trim();
            string vBelongDepNo = lbDepNo_Temp.Text.Trim();
            string vDepName = lbDepName_Temp.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select [Name] from Employee where EmpNo = '" + vUndertaker + "' and LeaveDay is null ";
            string vUndertakerName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vUndertakerName == "")
            {
                vUndertakerName = vUndertaker;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vUndertakerName + "' and LeaveDay is null ";
                vUndertaker = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }

            vSQLStr = "select count(EmpNo) ECount from Employee where EmpNo = '" + vUndertaker + "' and DepNo = '" + vBelongDepNo + "' ";
            int vEmpCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr, "ECount"));
            if (vEmpCount == 0)
            {
                lbUndertakerError_Temp.Text = "員工 [" + vUndertakerName + "] 並非單位 [" + vDepName + "] 人員";
                lbUndertakerError_Temp.Visible = true;
                eTemp.Text = "";
                lbTemp.Text = "";
                eTemp.Focus();
            }
            else
            {
                lbUndertakerError_Temp.Text = "";
                lbUndertakerError_Temp.Visible = false;
                eTemp.Text = vUndertaker.Trim();
                lbTemp.Text = vUndertakerName.Trim();
            }
        }

        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eDocIndex_Del = (Label)fvOfficialDocumentDetail.FindControl("eDocIndex_List");
            string vDocIndex_Del = eDocIndex_Del.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDelStr = "INSERT INTO [dbo].[OfficialDocumentHistory]" + Environment.NewLine +
                             "            ([DocIndex], [DocDate], [DocYears], [DocDep], [DocFirstWord], [DocNo], [DocSourceUnit], [DocType], " + Environment.NewLine +
                             "             [DocTitle], [Undertaker], [OutsideDocFirstWord], [OutsideDocNo], [Attachement], [Implementation]," + Environment.NewLine +
                             "             [BuildMan], [BuildDate], [Remark], [StoreDate], [StoreMan], [Remark_Store], [ModifyDate], [ModifyMode], [ModifyMan], [IsHide]) " + Environment.NewLine +
                             " select [DocIndex], [DocDate], [DocYears], [DocDep], [DocFirstWord], [DocNo], [DocSourceUnit], [DocType], " + Environment.NewLine +
                             "        [DocTitle], [Undertaker], [OutsideDocFirstWord], [OutsideDocNo], [Attachement], [Implementation]," + Environment.NewLine +
                             "        [BuildMan], [BuildDate], [Remark], [StoreDate], [StoreMan], [Remark_Store], GetDate(), 'DEL', '" + vLoginID + "', [IsHide] " + Environment.NewLine +
                             "   from [dbo].[OfficialDocument] " + Environment.NewLine +
                             "  where DocIndex = '" + vDocIndex_Del + "' ";
            try
            {
                PF.ExecSQL(vConnStr, vDelStr);
                sdsOfficialDocument_Detail.DeleteParameters.Clear();
                sdsOfficialDocument_Detail.DeleteParameters.Add(new Parameter("DocIndex", DbType.String, vDocIndex_Del));
                sdsOfficialDocument_Detail.Delete();
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        protected void bbStored_List_Click(object sender, EventArgs e)
        {
            string vSQLStr = "";
            Label lbDocIndex = (Label)fvOfficialDocumentDetail.FindControl("eDocIndex_List");
            string vDocIndex = lbDocIndex.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "update OfficialDocument" + Environment.NewLine +
                      "   set StoreDate = GetDate(), StoreMan = '" + vLoginID + "' " + Environment.NewLine +
                      " where DocIndex = '" + vDocIndex + "' ";
            PF.ExecSQL(vConnStr, vSQLStr);
            DocumentDataBind();
        }

        protected void sdsOfficialDocument_Detail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            /* 2023.06.30 改另一種寫法，加入判斷公文主旨不可空白
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eDocFirstWord_INS = (Label)fvOfficialDocumentDetail.FindControl("eDocFirstWord_Insert");
            TextBox eDocDate_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eDocDate_Insert");
            Label eBuildDate_INS = (Label)fvOfficialDocumentDetail.FindControl("eBuildDate_Insert");

            string vDocYear_Temp = DateTime.Parse(eDocDate_INS.Text.Trim()).Year.ToString("D4");
            string vDocIndex_Temp = vDocYear_Temp + eDocFirstWord_INS.Text.Trim();
            DateTime vDocDate_Temp = DateTime.Parse(eDocDate_INS.Text.Trim());
            DateTime vBuildDate_Temp = DateTime.Parse(eBuildDate_INS.Text.Trim());

            string vSQLStr = "select MAX(DocNo) DocNo_Max from OfficialDocument where DocIndex like '" + vDocIndex_Temp + "%' ";
            string vDocNo_INS = PF.GetValue(vConnStr, vSQLStr, "DocNo_Max");
            int vDocNo_I = (vDocNo_INS != "") ? Int32.Parse(vDocNo_INS) + 1 : 1;
            string vDocYears = ((vDocDate_Temp.Year) - 1911).ToString();

            e.Command.Parameters["@DocNo"].Value = vDocNo_I.ToString("D4");
            e.Command.Parameters["@DocIndex"].Value = vDocIndex_Temp + vDocNo_I.ToString("D4");
            e.Command.Parameters["@DocDate"].Value = vDocDate_Temp;
            e.Command.Parameters["@BuildDate"].Value = vBuildDate_Temp;
            e.Command.Parameters["@DocYears"].Value = vDocYears;
            //*/
        }

        protected void sdsOfficialDocument_Detail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                DocumentDataBind();
            }
        }

        protected void sdsOfficialDocument_Detail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                DocumentDataBind();
            }
        }

        protected void sdsOfficialDocument_Detail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                DocumentDataBind();
            }
        }

        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            //Label eDocYears_INS = (Label)fvOfficialDocumentDetail.FindControl("eDocYears_Insert");
            Label eDocFirstWord_INS = (Label)fvOfficialDocumentDetail.FindControl("eDocFirstWord_Insert");
            //Label eDocNo_INS = (Label)fvOfficialDocumentDetail.FindControl("eDocNo_Insert");
            TextBox eDocDate_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eDocDate_Insert");
            Label eDocDep_INS = (Label)fvOfficialDocumentDetail.FindControl("eDocDep_Insert");
            Label eDocType_INS = (Label)fvOfficialDocumentDetail.FindControl("eDocType_Insert");
            TextBox eUndertaker_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eUndertaker_Insert");
            TextBox eAttachement_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eAttachement_Insert");
            TextBox eDocSourceUnit_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eDocSourceUnit_Insert");
            TextBox eOutsideDocFirstWord_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eOutsideDocFirstWord_Insert");
            TextBox eOutsideDocNo_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eOutsideDocNo_Insert");
            TextBox eDocTitle_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eDocTitle_Insert");
            TextBox eImplementation_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eImplementation_Insert");
            TextBox eRemark_INS = (TextBox)fvOfficialDocumentDetail.FindControl("eRemark_Insert");
            //Label eBuildDate_INS = (Label)fvOfficialDocumentDetail.FindControl("eBuildDate_Insert");

            if ((eDocTitle_INS is null) || (eDocTitle_INS.Text.Trim() == ""))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('公文事由不可空白')");
                Response.Write("</" + "Script>");                
            }
            else
            {
                string vDocDep_Temp = eDocDep_INS.Text.Trim();
                string vDocFirstWord_Temp = eDocFirstWord_INS.Text.Trim();
                string vDocYear_Temp = DateTime.Parse(eDocDate_INS.Text.Trim()).Year.ToString("D4");
                string vDocIndex_Temp = vDocYear_Temp + eDocFirstWord_INS.Text.Trim();
                DateTime vDocDate_Temp = DateTime.Parse(eDocDate_INS.Text.Trim());
                string vDocDateStr_Temp = vDocDate_Temp.ToShortDateString();
                string vDocSourceUnit_Temp = eDocSourceUnit_INS.Text.Trim();
                string vDocType_Temp = eDocType_INS.Text.Trim();
                string vDocTitle_Temp = eDocTitle_INS.Text.Trim();
                string vUndertaker_Temp = eUndertaker_INS.Text.Trim();
                string vOutsideDocFirstWord_Temp = eOutsideDocFirstWord_INS.Text.Trim();
                string vOutsideDocNo_Temp = eOutsideDocNo_INS.Text.Trim();
                string vAttachement_Temp = eAttachement_INS.Text.Trim();
                string vImplementation_Temp = eImplementation_INS.Text.Trim();
                string vBuildDateStr_Temp = DateTime.Today.ToShortTimeString();
                string vRemark_Temp = eRemark_INS.Text.Trim();

                string vSQLStr = "select MAX(DocNo) DocNo_Max from OfficialDocument where DocIndex like '" + vDocIndex_Temp + "%' ";
                string vDocNo_INS = PF.GetValue(vConnStr, vSQLStr, "DocNo_Max");
                int vDocNo_I = (vDocNo_INS != "") ? Int32.Parse(vDocNo_INS) + 1 : 1;
                string vDocYears = ((vDocDate_Temp.Year) - 1911).ToString();

                string vSQLStr_Temp = "INSERT INTO OfficialDocument " + Environment.NewLine +
                                      "       (DocIndex, DocDate, DocDep, DocFirstWord, DocNo, DocSourceUnit, DocType, DocTitle, Undertaker, OutsideDocFirstWord, OutsideDocNo, " + Environment.NewLine +
                                      "        Attachement, Implementation, BuildMan, BuildDate, Remark, DocYears, IsHide) " + Environment.NewLine +
                                      "VALUES (@DocIndex, @DocDate, @DocDep, @DocFirstWord, @DocNo, @DocSourceUnit, @DocType, @DocTitle, @Undertaker, @OutsideDocFirstWord, @OutsideDocNo, " + Environment.NewLine +
                                      "        @Attachement, @Implementation, @BuildMan, @BuildDate, @Remark, @DocYears, @IsHide)";
                sdsOfficialDocument_Detail.InsertCommand = vSQLStr_Temp;
                sdsOfficialDocument_Detail.InsertParameters.Clear();
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("DocIndex", DbType.String, vDocIndex_Temp + vDocNo_I.ToString("D4")));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("DocDate", DbType.Date, vDocDateStr_Temp));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("DocDep", DbType.String, (vDocDep_Temp != "") ? vDocDep_Temp : String.Empty));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("DocFirstWord", DbType.String, vDocFirstWord_Temp));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("DocNo", DbType.String, vDocNo_I.ToString("D4")));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("DocSourceUnit", DbType.String, (vDocSourceUnit_Temp != "") ? vDocSourceUnit_Temp : String.Empty));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("DocType", DbType.String, (vDocType_Temp != "") ? vDocType_Temp : String.Empty));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("DocTitle", DbType.String, vDocTitle_Temp));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("Undertaker", DbType.String, (vUndertaker_Temp != "") ? vUndertaker_Temp : String.Empty));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("OutsideDocFirstWord", DbType.String, (vOutsideDocFirstWord_Temp != "") ? vOutsideDocFirstWord_Temp : String.Empty));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("OutsideDocNo", DbType.String, (vOutsideDocNo_Temp != "") ? vOutsideDocNo_Temp : String.Empty));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("Attachement", DbType.String, (vAttachement_Temp != "") ? vAttachement_Temp : String.Empty));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("Implementation", DbType.String, (vImplementation_Temp != "") ? vImplementation_Temp : String.Empty));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("BuildMan", DbType.String, vLoginID));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("BuildDate", DbType.Date, vBuildDateStr_Temp));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("Remark", DbType.String, (vRemark_Temp != "") ? vRemark_Temp : String.Empty));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("DocYears", DbType.String, vDocYears));
                sdsOfficialDocument_Detail.InsertParameters.Add(new Parameter("IsHide", DbType.Boolean, "false"));
                sdsOfficialDocument_Detail.Insert();
                gridDataList.DataBind();
                fvOfficialDocumentDetail.DataBind();
                fvOfficialDocumentDetail.ChangeMode(FormViewMode.ReadOnly);
            }
        }
    }
}