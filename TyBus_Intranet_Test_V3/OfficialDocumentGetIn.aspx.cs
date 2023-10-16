using Amaterasu_Function;
using System;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class OfficialDocumentGetIn : System.Web.UI.Page
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

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    string vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_Start_Search.ClientID;
                    string vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eDocDate_Start_Search.Attributes["onClick"] = vDocDateScript;
                    vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate_End_Search.ClientID;
                    vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eDocDate_End_Search.Attributes["onClick"] = vDocDateScript;
                    if (!IsPostBack)
                    {
                        eDocDate_Start_Search.Text = PF.GetMonthFirstDay(vToday.AddMonths(-1), "C");
                        eDocDate_End_Search.Text = PF.GetMonthLastDay(vToday, "C");
                        Session["DocGetIn"] = "-1";
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
            if (vLoginID != "")
            {
                string vWStr_DocDepNo = ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() != "")) ? "   and DocDep between '" + eDocDep_Start_Search.Text.Trim() + "' and '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                        ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() == "")) ? "   and DocDep = '" + eDocDep_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                        ((eDocDep_Start_Search.Text.Trim() == "") && (eDocDep_End_Search.Text.Trim() != "")) ? "   and DocDep = '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_DocDate = ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() != "")) ? "   and DocDate between '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() == "")) ? "   and DocDate = '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eDocDate_Start_Search.Text.Trim() == "") && (eDocDate_End_Search.Text.Trim() != "")) ? "   and DocDate = '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";
                string vWStr_DocType = (eDocType_Search.Text.Trim() != "") ? "   and DocType = '" + eDocType_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_Undertaker = (eUndertaker_Search.Text.Trim() != "") ? "   and Undertaker = '" + eUndertaker_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_DocTitle = (eDocTitle_Search.Text.Trim() != "") ? "   and DocTitle like '%" + eDocTitle_Search.Text.Trim() + "%' " + Environment.NewLine : "";
                string vWStr_FirstWord = (eFirstWord_Search.Text.Trim() != "") ? "   and DocFirstWord = '" + eFirstWord_Search.Text.Trim() + "' " + Environment.NewLine : "";
                string vWStr_IsStored = (cbIsGetIn_Search.Checked) ? "   and isnull(StoreDate, '') <> '' " + Environment.NewLine : "   and StoreDate is null " + Environment.NewLine;
                string vSelectStr = "select DocIndex, DocDate, DocYears, " + Environment.NewLine +
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
                                    "       Remark_Store " + Environment.NewLine +
                                    "  from OfficialDocument a " + Environment.NewLine +
                                    " where isnull(DocIndex, '') <> '' " + Environment.NewLine +
                                    vWStr_DocDepNo +
                                    vWStr_DocDate +
                                    vWStr_DocType +
                                    vWStr_Undertaker +
                                    vWStr_DocTitle +
                                    vWStr_FirstWord +
                                    vWStr_IsStored +
                                    " order by DocDate DESC, DocDep, DocNo DESC";
                sdsOfficialDocument_List.SelectCommand = "";
                sdsOfficialDocument_List.SelectCommand = vSelectStr;
                sdsOfficialDocument_List.DataBind();
            }
        }

        private void GetFirstCWordList()
        {
            string vWStr_DepNo = ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() != "")) ? "   and DepNo between '" + eDocDep_Start_Search.Text.Trim() + "' and '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() == "")) ? "   and DepNo = '" + eDocDep_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDocDep_Start_Search.Text.Trim() == "") && (eDocDep_End_Search.Text.Trim() != "")) ? "   and DepNo = '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine : "";
            //string vWStr_EmpNo = (eUndertaker_Search.Text.Trim() == "731520") ? "   and FWNo = 'FW0204'" + Environment.NewLine :
            string vWStr_EmpNo = (eUndertaker_Search.Text.Trim() == "938122") ? "   and FWNo = 'FW0204'" + Environment.NewLine :
                                 (eUndertaker_Search.Text.Trim() == "040712") ? "   and FWNo in ('FW0201', 'FW0203', 'FW0205', 'FW0206')" + Environment.NewLine :
                                 (eUndertaker_Search.Text.Trim() == "740112") ? "   and FWNo = 'FW0202'" + Environment.NewLine :
                                 (eUndertaker_Search.Text.Trim() == "080895") ? "   and FWNo = 'FW0207'" + Environment.NewLine :
                                 ((eUndertaker_Search.Text.Trim() == "020219") || (eUndertaker_Search.Text.Trim() == "057560") || (eUndertaker_Search.Text.Trim() == "948173")) ? "   and FWNo = 'FW0201'" + Environment.NewLine :
                                 (eUndertaker_Search.Text.Trim() != "") ? "   and DepNo = (select DepNo from EMployee where EmpNo = '" + eUndertaker_Search.Text.Trim() + "')" : "";
            string vSelectStr = "select Cast('' as varchar) FWNo, cast('' as varchar) DocFirstCWord " + Environment.NewLine +
                                "union all " + Environment.NewLine +
                                "select FWNo, DocFirstCWord from DOCFirstWord " + Environment.NewLine +
                                " where 1 = 1 " + vWStr_DepNo + vWStr_EmpNo;
            sdsFirstWord_Search.SelectCommand = "";
            sdsFirstWord_Search.SelectCommand = vSelectStr;
            ddlFirstWord_Search.DataBind();
        }

        protected void eDocDep_Start_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDocDep_Start_Search.Text.Trim();
            string vDepName_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSelectStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSelectStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSelectStr, "DepNo");
            }
            eDocDep_Start_Search.Text = vDepNo_Temp;
            eDocDepName_Start_Search.Text = vDepName_Temp;
        }

        protected void eDocDep_End_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDocDep_End_Search.Text.Trim();
            string vDepName_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSelectStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSelectStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSelectStr, "DepNo");
            }
            eDocDep_End_Search.Text = vDepNo_Temp;
            eDocDepName_End_Search.Text = vDepName_Temp;
        }

        protected void eUndertaker_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo = eUndertaker_Search.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpName = "";
            string vSelectStr = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' ";
            vEmpName = PF.GetValue(vConnStr, vSelectStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSelectStr = "select EmpNo from Employee where [Name] = '" + vEmpName + "' ";
                vEmpNo = PF.GetValue(vConnStr, vSelectStr, "EmpNo");
            }
            eUndertaker_Search.Text = vEmpNo;
            eUndertakerName_Search.Text = vEmpName;
            GetFirstCWordList();
        }

        protected void ddlFirstWord_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eFirstWord_Search.Text = ddlFirstWord_Search.SelectedValue.Trim();
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
                    Label eStoreDate_List = (Label)fvOfficialDocumentDetail.FindControl("eStoreDate_List");
                    Button bbEdit_List = (Button)fvOfficialDocumentDetail.FindControl("bbEdit_List");
                    Button bbDelete_List = (Button)fvOfficialDocumentDetail.FindControl("bbDelete_List");
                    if ((eStoreDate_List != null) && (bbEdit_List != null) && (bbDelete_List != null))
                    {
                        if (eStoreDate_List.Text.Trim() != "")
                        {
                            bbEdit_List.Visible = false;
                            bbDelete_List.Visible = true;
                        }
                        else
                        {
                            bbEdit_List.Visible = true;
                            bbDelete_List.Visible = false;
                        }
                    }
                    break;
                case FormViewMode.Edit:
                    TextBox eStoreDate = (TextBox)fvOfficialDocumentDetail.FindControl("eStoreDate_Edit");
                    Label eStoreMan = (Label)fvOfficialDocumentDetail.FindControl("eStoreMan_Edit");
                    Label eStoreMan_C = (Label)fvOfficialDocumentDetail.FindControl("eStoreMan_C_Edit");

                    string vStoreDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eStoreDate.ClientID;
                    string vStoreDateScript = "window.open('" + vStoreDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eStoreDate.Attributes["onClick"] = vStoreDateScript;

                    eStoreDate.Text = DateTime.Today.ToString("yyyy/MM/dd");
                    eStoreMan.Text = vLoginID;
                    eStoreMan_C.Text = (string)Session["LoginName"];
                    break;
                case FormViewMode.Insert:
                    break;
                default:
                    break;
            }
        }

        protected void cbSelectIndex_CheckedChanged(object sender, EventArgs e)
        {
            string vDocIndex = "";
            int vFieldIndex = 0;
            for (int i = 0; i < gridDataList.Rows.Count; i++)
            {
                for (int j = 0; j < gridDataList.Rows[i].Cells.Count; j++)
                {
                    if (gridDataList.Columns[j].HeaderText == "序號")
                    {
                        vFieldIndex = j;
                    }
                }
                CheckBox cbTemp_Print = (CheckBox)gridDataList.Rows[i].FindControl("cbSelectIndex");
                vDocIndex = gridDataList.Rows[i].Cells[vFieldIndex].Text.Trim();

                if (cbTemp_Print.Checked)
                {
                    if (Session["DocGetIn"].ToString().IndexOf(vDocIndex, 0) == -1)
                    {
                        if ((Session["DocGetIn"].ToString().Trim() == "-1") || (Session["DocGetIn"].ToString().Trim() == ""))
                        {
                            Session["DocGetIn"] = "'" + vDocIndex + "'";
                        }
                        else
                        {
                            Session["DocGetIn"] = Session["DocGetIn"].ToString().Trim() + "," + "'" + vDocIndex + "'";
                        }
                    }
                }
                else
                {
                    if (Session["DocGetIn"].ToString().Trim() != "-1")
                    {
                        Session["DocGetIn"] = Session["DocGetIn"].ToString().Trim().Replace("'" + vDocIndex + "'", "");
                        Session["DocGetIn"] = Session["DocGetIn"].ToString().Trim().Replace(",,", ",");
                        if ((Session["DocGetIn"].ToString().Length > 0) && (Session["DocGetIn"].ToString().Substring(0, 1) == ","))
                        {
                            Session["DocGetIn"] = Session["DocGetIn"].ToString().Substring(1, Session["DocGetIn"].ToString().Length - 1);
                        }
                        else if ((Session["DocGetIn"].ToString().Length > 0) && (Session["DocGetIn"].ToString().Substring(Session["DocGetIn"].ToString().Length - 2, 1) == ","))
                        {
                            Session["DocGetIn"] = Session["DocGetIn"].ToString().Substring(0, Session["DocGetIn"].ToString().Length - 1);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 批次歸檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void BatchGetIn_Click(object sender, EventArgs e)
        {
            if (fvOfficialDocumentDetail.CurrentMode == FormViewMode.Edit)
            {
                fvOfficialDocumentDetail.ChangeMode(FormViewMode.ReadOnly);
            }
            string vSelectIndex = Session["DocGetIn"].ToString().Trim();
            if ((vSelectIndex.Length == 0) || (vSelectIndex == "-1"))
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇歸檔公文！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSQLStr = "update OfficialDocument set StoreMan = '" + vLoginID + "', StoreDate = GetDate() where DocIndex in (" + (string)Session["DocGetIn"] + ") ";
                PF.ExecSQL(vConnStr, vSQLStr);
                DocumentDataBind();
            }
        }

        protected void sdsOfficialDocument_Detail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            TextBox eStoreDate = (TextBox)fvOfficialDocumentDetail.FindControl("eStoreDate_Edit");
            DateTime vStoreDate = DateTime.Parse(eStoreDate.Text.Trim());
            e.Command.Parameters["@StoreDate"].Value = vStoreDate;
        }

        protected void sdsOfficialDocument_Detail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridDataList.DataBind();
            }
        }

        /// <summary>
        /// 解除公文歸檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eDocIndex_List = (Label)fvOfficialDocumentDetail.FindControl("eDocIndex_List");
            if ((eDocIndex_List != null) && (eDocIndex_List.Text.Trim() != ""))
            {
                string vDocIndexTemp = eDocIndex_List.Text.Trim();
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    string vRecordNote = "解除公文歸檔_" +vDocIndexTemp + Environment.NewLine +
                                         "OfficialDocumentGetIn.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    string vSQLStrTemp = "update OfficialDocument set StoreDate = null, StoreMan = null where DocIndex = '" + vDocIndexTemp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStrTemp);
                    gridDataList.DataBind();
                    fvOfficialDocumentDetail.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert(" + eMessage.Message + ")");
                    Response.Write("</" + "Script>");
                }
            }
        }
    }
}