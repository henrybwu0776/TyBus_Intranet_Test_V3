using Amaterasu_Function;
using System;
using System.Data;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class EditOfficialDocument : System.Web.UI.Page
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

        private string GetSelStr()
        {
            string vWStr_DocNo = (eDocNo_Search.Text.Trim() != "") ? " and DocNo = '" + eDocNo_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vWStr_DocDepNo = ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() != "")) ? " and DocDep between '" + eDocDep_Start_Search.Text.Trim() + "' and '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDocDep_Start_Search.Text.Trim() != "") && (eDocDep_End_Search.Text.Trim() == "")) ? " and DocDep = '" + eDocDep_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDocDep_Start_Search.Text.Trim() == "") && (eDocDep_End_Search.Text.Trim() != "")) ? " and DocDep = '" + eDocDep_End_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vWStr_DocDate = ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() != "")) ? " and DocDate between '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                   ((eDocDate_Start_Search.Text.Trim() != "") && (eDocDate_End_Search.Text.Trim() == "")) ? " and DocDate = '" + PF.TransDateString(eDocDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                   ((eDocDate_Start_Search.Text.Trim() == "") && (eDocDate_End_Search.Text.Trim() != "")) ? " and DocDate = '" + PF.TransDateString(eDocDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine : "";

            string vWStr_DocType = (eDocType_Search.Text.Trim() != "") ? " and DocType = '" + eDocType_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vWStr_Undertaker = ((eUndertaker_Start_Search.Text.Trim() != "") && (eUndertaker_End_Search.Text.Trim() != "")) ? " and Undertaker between '" + eUndertaker_Start_Search.Text.Trim() + "' and '" + eUndertaker_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                      ((eUndertaker_Start_Search.Text.Trim() != "") && (eUndertaker_End_Search.Text.Trim() == "")) ? " and Undertaker = '" + eUndertaker_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                      ((eUndertaker_Start_Search.Text.Trim() == "") && (eUndertaker_End_Search.Text.Trim() != "")) ? " and Undertaker = '" + eUndertaker_End_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vWStr_DocTitle = (eDocTitle_Search.Text.Trim() != "") ? " and DocTitle like '%" + eDocTitle_Search.Text.Trim() + "%' " + Environment.NewLine : "";

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
                                " where StoreDate is null " + Environment.NewLine +
                                vWStr_DocDepNo + vWStr_DocDate + vWStr_DocType + vWStr_Undertaker + vWStr_DocTitle + vWStr_DocNo +
                                " order by DocDate DESC, DocDep, DocNo DESC";
            return vSelectStr;
        }

        private void DocumentDataBind()
        {
            sdsOfficialDocument_List.SelectCommand = "";
            sdsOfficialDocument_List.SelectCommand = GetSelStr();
            gridDataList.DataBind();
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

        protected void ddlDocType_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eDocType_Search.Text = ddlDocType_Search.SelectedValue;
        }

        protected void eUndertaker_Start_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo = eUndertaker_Start_Search.Text.Trim();
            string vEmpName = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "select [Name] from Employee where LeaveDay is null and EmpNo = '" + vEmpNo + "' ";
            vEmpName = PF.GetValue(vConnStr, vSelectStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSelectStr = "select EmpNo from Employee where LeaveDay is null and [Name] = '" + vEmpName + "' ";
                vEmpNo = PF.GetValue(vConnStr, vSelectStr, "EmpNo");
            }
            eUndertaker_Start_Search.Text = vEmpNo;
            eUndertakerName_Start_Search.Text = vEmpName;
        }

        protected void eUndertaker_End_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo = eUndertaker_End_Search.Text.Trim();
            string vEmpName = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = "select [Name] from Employee where LeaveDay is null and EmpNo = '" + vEmpNo + "' ";
            vEmpName = PF.GetValue(vConnStr, vSelectStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSelectStr = "select EmpNo from Employee where LeaveDay is null and [Name] = '" + vEmpName + "' ";
                vEmpNo = PF.GetValue(vConnStr, vSelectStr, "EmpNo");
            }
            eUndertaker_End_Search.Text = vEmpNo;
            eUndertakerName_End_Search.Text = vEmpName;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            DocumentDataBind();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvEditOfficialDocument_DataBound(object sender, EventArgs e)
        {
            if (fvEditOfficialDocument.CurrentMode == FormViewMode.Edit)
            {
                TextBox eDocDate = (TextBox)fvEditOfficialDocument.FindControl("eDocDate_Edit");
                string vDocDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDocDate.ClientID;
                string vDocDateScript = "window.open('" + vDocDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                eDocDate.Attributes["onClick"] = vDocDateScript;
            }
        }

        protected void eUndertaker_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eUndertaker = (TextBox)fvEditOfficialDocument.FindControl("eUndertaker_Edit");
            Label eUndertaker_C = (Label)fvEditOfficialDocument.FindControl("eUndertaker_C_Edit");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo = eUndertaker.Text.Trim();
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' and LeaveDay is null ";
            string vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName + "' and LeaveDay is null ";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eUndertaker.Text = vEmpNo;
            eUndertaker_C.Text = vEmpName;
        }

        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eDocIndex_DEL = (Label)fvEditOfficialDocument.FindControl("eDocIndex_List");
            string vDocIndex_Del = eDocIndex_DEL.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select MAX(Items) Items from OfficialDocumentHistory where DocIndex = '" + vDocIndex_Del + "' and ModifyMode = 'DEL' ";
            string vMaxIndex = PF.GetValue(vConnStr, vSQLStr, "Items");
            int vNewIndex = (vMaxIndex == "") ? 1 : Int32.Parse(vMaxIndex) + 1;
            string vDelStr = "INSERT INTO [dbo].[OfficialDocumentHistory]" + Environment.NewLine +
                             "            ([DocIndex], [Items], [DocDate], [DocDep], [DocYears], [DocFirstWord], [DocNo], [DocSourceUnit], [DocType], " + Environment.NewLine +
                             "             [DocTitle], [Undertaker], [OutsideDocFirstWord], [OutsideDocNo], [Attachement], [Implementation]," + Environment.NewLine +
                             "             [BuildMan], [BuildDate], [Remark], [StoreDate], [StoreMan], [Remark_Store], [ModifyDate], [ModifyMode], [ModifyMan], [IsHide]) " + Environment.NewLine +
                             " select [DocIndex], '" + vNewIndex.ToString("D4") + "', [DocDate], [DocDep], [DocYears], [DocFirstWord], [DocNo], [DocSourceUnit], [DocType], " + Environment.NewLine +
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

        protected void sdsOfficialDocument_Detail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
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

        protected void sdsOfficialDocument_Detail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            TextBox eDocDate_Temp = (TextBox)fvEditOfficialDocument.FindControl("eDocDate_Edit");
            Label eDocIndex_Temp = (Label)fvEditOfficialDocument.FindControl("eDocIndex_Edit");
            if (eDocDate_Temp != null)
            {
                string vDocIndex_Edit = eDocIndex_Temp.Text.Trim();

                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSQLStr = "select MAX(Items) Items from OfficialDocumentHistory where DocIndex = '" + vDocIndex_Edit + "' and ModifyMode = 'EDIT' ";
                string vMaxIndex = PF.GetValue(vConnStr, vSQLStr, "Items");
                int vNewIndex = (vMaxIndex == "") ? 1 : Int32.Parse(vMaxIndex) + 1;
                vSQLStr = "INSERT INTO [dbo].[OfficialDocumentHistory]" + Environment.NewLine +
                             "            ([DocIndex], [Items], [DocDate], [DocDep], [DocYears], [DocFirstWord], [DocNo], [DocSourceUnit], [DocType], " + Environment.NewLine +
                             "             [DocTitle], [Undertaker], [OutsideDocFirstWord], [OutsideDocNo], [Attachement], [Implementation]," + Environment.NewLine +
                             "             [BuildMan], [BuildDate], [Remark], [StoreDate], [StoreMan], [Remark_Store], [ModifyDate], [ModifyMode], [ModifyMan], [IsHide]) " + Environment.NewLine +
                             " select [DocIndex], '" + vNewIndex.ToString("D4") + "', [DocDate], [DocDep], [DocYears], [DocFirstWord], [DocNo], [DocSourceUnit], [DocType], " + Environment.NewLine +
                             "        [DocTitle], [Undertaker], [OutsideDocFirstWord], [OutsideDocNo], [Attachement], [Implementation]," + Environment.NewLine +
                             "        [BuildMan], [BuildDate], [Remark], [StoreDate], [StoreMan], [Remark_Store], GetDate(), 'EDIT', '" + vLoginID + "', [IsHide] " + Environment.NewLine +
                             "   from [dbo].[OfficialDocument] " + Environment.NewLine +
                             "  where DocIndex = '" + vDocIndex_Edit + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStr);
                    if ((eDocDate_Temp != null) && (eDocDate_Temp.Text.Trim() != ""))
                    {
                        DateTime vDocDate = DateTime.Parse(eDocDate_Temp.Text.Trim());
                        e.Command.Parameters["@DocDate"].Value = vDocDate;
                    }
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