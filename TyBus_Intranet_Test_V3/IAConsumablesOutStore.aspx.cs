using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class IAConsumablesOutStore : Page
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
                DateTime vToday = DateTime.Today;
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

                    //開單日期
                    string vCaseDateURL = "InputDate.aspx?TextboxID=" + eBuDateS_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDateS_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eBuDateE_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDateE_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        plShowData.Visible = true;
                        plPrint.Visible = false;
                    }
                    else
                    {
                        //把 OpenData 寫在這裡，確保第一次打開頁面時不會自動查詢資料，但是每次查詢資料或是切換頁面後會更新內容
                        OpenData();
                    }
                    plShowData_Detail.Visible = (gridIASheetA_List.Rows.Count > 0);
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

        /// <summary>
        /// 取回資料查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelStr()
        {
            string vResultStr = "";

            String vBuDate_S = (eBuDateS_Search.Text.Trim() != "") ? PF.TransDateString(eBuDateS_Search.Text.Trim(), "B") : DateTime.Now.ToShortDateString();
            string vBuDate_E = (eBuDateE_Search.Text.Trim() != "") ? PF.TransDateString(eBuDateE_Search.Text.Trim(), "B") : DateTime.Now.ToShortDateString();

            string vWStr_SheetNo = ((eSheetNoS_Search.Text.Trim() != "") && (eSheetNoE_Search.Text.Trim() != "")) ? "   and a.SheetNo between '" + eSheetNoS_Search.Text.Trim() + "' and '" + eSheetNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                   ((eSheetNoS_Search.Text.Trim() != "") && (eSheetNoE_Search.Text.Trim() == "")) ? "   and a.SheetNo = '" + eSheetNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                   ((eSheetNoS_Search.Text.Trim() == "") && (eSheetNoE_Search.Text.Trim() != "")) ? "   and a.SheetNo = '" + eSheetNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo = ((eDepNameS_Search.Text.Trim() != "") && (eDepNameE_Search.Text.Trim() != "")) ?
                                 "   and a.DepNo between '" + eDepNoS_Search.Text.Trim() + "' and '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNameS_Search.Text.Trim() != "") && (eDepNameE_Search.Text.Trim() == "")) ?
                                 "   and a.DepNo = '" + eDepNoS_Search.Text.Trim() + "' " + Environment.NewLine :
                                 ((eDepNameS_Search.Text.Trim() == "") && (eDepNameE_Search.Text.Trim() != "")) ?
                                 "   and a.DepNo = '" + eDepNoE_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_SheetStatus = (ddlSheetStatus_Search.SelectedValue.Trim() != "") ? "   and a.SheetStatus = '" + ddlSheetStatus_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuMan = (eBuMan_Search.Text.Trim() != "") ? "   and a.BuMan = '" + eBuMan_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuDate = ((eBuDateS_Search.Text.Trim() != "") && (eBuDateE_Search.Text.Trim() != "")) ? "   and a.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine :
                                  ((eBuDateS_Search.Text.Trim() != "") && (eBuDateE_Search.Text.Trim() == "")) ? "   and a.BuDate = '" + vBuDate_S + "' " + Environment.NewLine :
                                  ((eBuDateS_Search.Text.Trim() == "") && (eBuDateE_Search.Text.Trim() != "")) ? "   and a.BuDate = '" + vBuDate_E + "' " + Environment.NewLine : "";
            string vWStr_SheetNote = (eSheetNote_Search.Text.Trim() != "") ? "   and a.SheetNote like '%" + eSheetNote_Search.Text.Trim() + "%' " + Environment.NewLine : "";

            string vWStr_ConsNo = (eConsNo_Search.Text.Trim() != "") ? "           and ConsNo = '" + eConsNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "           and ConsNo in (select ConsNo from IAConsumables where ConsName lke '%" + eConsName_Search.Text.Trim() + "%') " + Environment.NewLine : "";
            string vWStr_Brand = (eBrand_Search.Text.Trim() != "") ? "           and ConsNo in (select ConsNo from IAConsumables where Brand = '" + eBrand_Search.Text.Trim() + "') " + Environment.NewLine : "";
            string vWStr_CorrespondModel = (eCorrespondModel_Search.Text.Trim() != "") ? "           and ConsNo in (select ConsNo from IAConsumables where CorrespondModel like '%" + eCorrespondModel_Search.Text.Trim() + "%') " + Environment.NewLine : "";
            string vWStr_Substr = ((eConsNo_Search.Text.Trim() == "") && (eConsName_Search.Text.Trim() == "") && (eBrand_Search.Text.Trim() == "") && (eCorrespondModel_Search.Text.Trim() == "")) ? "" :
                                  "   and a.SheetNo in " + Environment.NewLine +
                                  "       (select distinct SheetNo " + Environment.NewLine +
                                  "          from IASheetB " + Environment.NewLine +
                                  "         where isnull(SheetNo, '') <> '' " + Environment.NewLine +
                                  vWStr_ConsNo +
                                  vWStr_ConsName +
                                  vWStr_Brand +
                                  vWStr_CorrespondModel +
                                  "       )" + Environment.NewLine;


            vResultStr = "SELECT SheetNo, DepNo, (SELECT name FROM Department WHERE DepNo = a.DepNo) AS DepName, BuDate, " + Environment.NewLine +
                         "       AssignMan, (select [Name] from Employee where EmpNo = a.AssignMan) AssignManName,TotalAmount, SheetNote, SheetStatus, " + Environment.NewLine +
                         "       (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      SheetStatus') AND (CLASSNO = a.SheetStatus)) AS SheetStatus_C, RemarkA " + Environment.NewLine +
                         "  FROM IASheetA AS a " + Environment.NewLine +
                         " WHERE ISNULL(a.SheetNo, '') <> '' " + Environment.NewLine +
                         "   AND SheetMode = 'SO' " + Environment.NewLine +
                         vWStr_SheetNo +
                         vWStr_DepNo +
                         vWStr_SheetStatus +
                         vWStr_BuMan +
                         vWStr_BuDate +
                         vWStr_SheetNote +
                         vWStr_Substr +
                         " ORDER BY a.SheetNo DESC ";
            return vResultStr;
        }

        /// <summary>
        /// 取得資料
        /// </summary>
        private void OpenData()
        {
            string vSelectStr = GetSelStr();
            dsIASheetA_List.SelectCommand = vSelectStr;
            gridIASheetA_List.DataBind();
        }

        protected void eDepNoS_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNameS_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName == "")
            {
                vDepName = vDepNo.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] = '" + vDepName + "' ";
                vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNameS_Search.Text = vDepNo.Trim();
            eDepNameS_Search.Text = vDepName.Trim();
        }

        protected void eDepNoE_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNameE_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName == "")
            {
                vDepName = vDepNo.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] = '" + vDepName + "' ";
                vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNameE_Search.Text = vDepNo.Trim();
            eDepNameE_Search.Text = vDepName.Trim();
        }

        protected void eBuMan_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo = eBuMan_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' ";
            string vEmpName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName.Trim() + "' Order by Assumeday DESC";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eBuMan_Search.Text = vEmpNo.Trim();
            eBuManName_Search.Text = vEmpName.Trim();
        }

        protected void bbOK_Search_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// FormView 狀態變更時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void fvIASheetA_Detail_DataBound(object sender, EventArgs e)
        {
            switch (fvIASheetA_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
                    if (eSheetNo_List != null)
                    {
                        Button bbOutStoreA_List = (Button)fvIASheetA_Detail.FindControl("bbOutStoreA_List");
                        Button bbAbortA_List = (Button)fvIASheetA_Detail.FindControl("bbAbortA_List");
                        Label eSheetStatus_List = (Label)fvIASheetA_Detail.FindControl("eSheetStatus_List");
                        switch (eSheetStatus_List.Text.Trim())
                        {
                            default:
                                bbOutStoreA_List.Enabled = false;
                                bbAbortA_List.Enabled = false;
                                break;
                            case "000":
                                bbOutStoreA_List.Enabled = true;
                                bbAbortA_List.Enabled = true;
                                break;
                            case "100":
                                bbOutStoreA_List.Enabled = false;
                                bbAbortA_List.Enabled = true;
                                break;
                        }
                    }
                    break;
                case FormViewMode.Edit:
                    Label eSheetNo_Edit = (Label)fvIASheetA_Detail.FindControl("eSheetNo_Edit");
                    if (eSheetNo_Edit != null)
                    {

                    }
                    break;
                case FormViewMode.Insert:
                    Label eSheetNo_INS = (Label)fvIASheetA_Detail.FindControl("eSheetNo_INS");
                    if (eSheetNo_INS != null)
                    {

                    }
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 主檔完成編輯
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            Label eSheetNo_Edit = (Label)fvIASheetA_Detail.FindControl("eSheetNo_Edit");
            if ((eSheetNo_Edit != null) && (eSheetNo_Edit.Text.Trim() != ""))
            {
                TextBox eDepNo_Edit = (TextBox)fvIASheetA_Detail.FindControl("eDepNo_Edit");
                TextBox eAssignMan_Edit = (TextBox)fvIASheetA_Detail.FindControl("eAssignMan_Edit");
                TextBox eSheetNote_Edit = (TextBox)fvIASheetA_Detail.FindControl("eSheetNote_Edit");
                TextBox eRemarkA_Edit = (TextBox)fvIASheetA_Detail.FindControl("eRemarkA_Edit");
                string vUpdateStr = "UPDATE IASheetA " + Environment.NewLine +
                                    "   SET SheetNote = @SheetNote, RemarkA = @RemarkA, DepNo = @DepNo, ModifyDate = GETDATE(), ModifyMan = @ModifyMan, AssignMan = @AssignMan " + Environment.NewLine +
                                    " WHERE (SheetNo = @SheetNo)";
                try
                {
                    dsIASheetA_Detail.UpdateCommand = vUpdateStr;
                    dsIASheetA_Detail.UpdateParameters.Clear();
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("SheetNote", DbType.String, (eSheetNote_Edit.Text.Trim() != "") ? eSheetNote_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA_Edit.Text.Trim() != "") ? eRemarkA_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_Edit.Text.Trim() != "") ? eDepNo_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan_Edit.Text.Trim() != "") ? eAssignMan_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, eSheetNo_Edit.Text.Trim()));
                    dsIASheetA_Detail.Update();
                    gridIASheetA_List.DataBind();
                    fvIASheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo_Edit = (TextBox)fvIASheetA_Detail.FindControl("eDepNo_Edit");
            Label eDepName_Edit = (Label)fvIASheetA_Detail.FindControl("eDepName_Edit");
            if (eDepNo_Edit != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vDepNo = eDepNo_Edit.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = vDepNo.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '" + vDepName + "%' order by DepNo DESC";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_Edit.Text = vDepNo.Trim();
                eDepName_Edit.Text = vDepName.Trim();
            }
        }

        protected void eAssignMan_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eAssignMan_Edit = (TextBox)fvIASheetA_Detail.FindControl("eAssignMan_Edit");
            Label eAssignManName_Edit = (Label)fvIASheetA_Detail.FindControl("eAssignManName_Edit");
            if (eAssignMan_Edit != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vAssignMan = eAssignMan_Edit.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
                string vAssignManName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vAssignManName == "")
                {
                    vAssignManName = vAssignMan.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vAssignManName + "' order by EmpNo DESC ";
                    vAssignMan = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eAssignMan_Edit.Text = vAssignMan.Trim();
                eAssignManName_Edit.Text = vAssignManName.Trim();
            }
        }

        /// <summary>
        /// 主檔新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            Label eSheetNo_INS = (Label)fvIASheetA_Detail.FindControl("eSheetNo_INS");
            if (eSheetNo_INS != null)
            {
                //產生單號
                string vSheetNo = "";
                string vSQLStr_Temp = "select max(SheetNo) MaxIndex from IASheetA where SheetMode = 'SO' ";
                string vOldSheetNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                string vOldIndex = (vOldSheetNo != "") ? vOldSheetNo.Substring(8) : "0";
                int vNewIndex = 0;
                if (int.TryParse(vOldIndex, out vNewIndex))
                {
                    vNewIndex++;
                }
                else
                {
                    vNewIndex = 1;
                }
                vSheetNo = DateTime.Now.Year.ToString("D4") + DateTime.Now.Month.ToString("D2") + "SO" + vNewIndex.ToString("D4");
                string vInsertStr = "insert into IASheetA (SheetNo, SheetMode, BuDate, BuMan, SheetNote, SheetStatus, StatusDate, RemarkA, DepNo, AssignMan) " + Environment.NewLine +
                                    "              values (@SheetNo, 'SO', GetDate(), @BuMan, @SheetNote, '000', GetDate(), @RemarkA, @DepNo, @AssignMan)";
                TextBox eSheetNote_INS = (TextBox)fvIASheetA_Detail.FindControl("eSheetNote_INS");
                TextBox eRemarkA_INS = (TextBox)fvIASheetA_Detail.FindControl("eRemarkA_INS");
                TextBox eDepNo_INS = (TextBox)fvIASheetA_Detail.FindControl("eDepNo_INS");
                TextBox eAssignMan_INS = (TextBox)fvIASheetA_Detail.FindControl("eAssignMan_INS");
                try
                {
                    dsIASheetA_Detail.InsertCommand = vInsertStr;
                    dsIASheetA_Detail.InsertParameters.Clear();
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("SheetNote", DbType.String, (eSheetNote_INS.Text.Trim() != "") ? eSheetNote_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA_INS.Text.Trim() != "") ? eRemarkA_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo_INS.Text.Trim() != "") ? eDepNo_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan_INS.Text.Trim() != "") ? eAssignMan_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.Insert();
                    fvIASheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                    gridIASheetA_List.DataBind();
                    fvIASheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }

            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo_INS = (TextBox)fvIASheetA_Detail.FindControl("eDepNo_INS");
            Label eDepName_INS = (Label)fvIASheetA_Detail.FindControl("eDepName_INS");
            if (eDepNo_INS != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vDepNo = eDepNo_INS.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = vDepNo.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '" + vDepName + "%' order by DepNo DESC";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_INS.Text = vDepNo.Trim();
                eDepName_INS.Text = vDepName.Trim();
            }
        }

        protected void eAssignMan_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eAssignMan_INS = (TextBox)fvIASheetA_Detail.FindControl("eAssignMan_INS");
            Label eAssignManName_INS = (Label)fvIASheetA_Detail.FindControl("eAssignManName_INS");
            if (eAssignMan_INS != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vAssignMan = eAssignMan_INS.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
                string vAssignManName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vAssignManName == "")
                {
                    vAssignManName = vAssignMan.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vAssignManName + "' order by EmpNo DESC ";
                    vAssignMan = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eAssignMan_INS.Text = vAssignMan.Trim();
                eAssignManName_INS.Text = vAssignManName.Trim();
            }
        }

        /// <summary>
        /// 刪除主檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDeleteA_List_Click(object sender, EventArgs e)
        {
            Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
            if ((eSheetNo_List != null) && (eSheetNo_List.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    string vSheetNo = eSheetNo_List.Text.Trim();
                    //string vConsNo = "";
                    string vSheetNoItems = "";
                    /* 2022.07.14 計算數量交給資料庫的 Trigger 做
                    using (SqlConnection connCalQty=new SqlConnection())
                    {
                        connCalQty.ConnectionString = vConnStr;
                        SqlCommand cmdCalQty = new SqlCommand("select ConsNo from IASheetB where SheetNo = '" + vSheetNo + "' ", connCalQty);
                        connCalQty.Open();
                        SqlDataReader drCalQty = cmdCalQty.ExecuteReader();
                        while (drCalQty.Read())
                        {
                            vConsNo = drCalQty["ConsNo"].ToString().Trim();
                            PF.CalIAStoreQuantity(vConsNo, vConnStr, vSheetNo);
                        }
                    } //*/
                    //因為主檔要刪掉，所以先把明細刪除
                    //dsIASheetB_Detail.DeleteCommand = "delete IASheetB where SheetNo = @SheetNo";
                    //dsIASheetB_Detail.DeleteParameters.Clear();
                    //dsIASheetB_Detail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    //dsIASheetB_Detail.Delete();
                    // 2022.08.30 改成一個明細呼叫一次 Delete()
                    using (SqlConnection connDelete = new SqlConnection())
                    {
                        connDelete.ConnectionString = vConnStr;
                        SqlCommand cmdDelete = new SqlCommand("select SheetNoItems from IASheetB where SheetNo = '" + vSheetNo + "' ", connDelete);
                        connDelete.Open();
                        SqlDataReader drDelete = cmdDelete.ExecuteReader();
                        while (drDelete.Read())
                        {
                            vSheetNoItems = drDelete["SheetNoItems"].ToString().Trim();
                            dsIASheetB_Detail.DeleteCommand = "delete IASheetB where SheetNoItems = @SheetNoItems";
                            dsIASheetB_Detail.DeleteParameters.Clear();
                            dsIASheetB_Detail.DeleteParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                            dsIASheetB_Detail.Delete();
                        }
                    }
                    //明細刪完刪主檔
                    dsIASheetA_Detail.DeleteCommand = "delete IASheetA where SheetNo = @SheetNo";
                    dsIASheetA_Detail.DeleteParameters.Clear();
                    dsIASheetA_Detail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    dsIASheetA_Detail.Delete();
                    gridIASheetA_List.DataBind();
                    fvIASheetA_Detail.DataBind();
                    gridIASheetB_List.DataBind();
                    fvIASheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// 出貨 (列印簽收單)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOutStoreA_List_Click(object sender, EventArgs e)
        {
            string vSQLStrTemp = "";
            string vSheetNo = "";
            //string vConsNo = "";
            Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
            if (eSheetNo_List != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                vSheetNo = eSheetNo_List.Text.Trim();
                DataTable dtIASheetA = new DataTable();
                DataTable dtIASheetB = new DataTable();
                try
                {
                    using (SqlConnection connPrint = new SqlConnection())
                    {
                        connPrint.ConnectionString = vConnStr;

                        vSQLStrTemp = "select SheetNo, DepNo, (select [Name] from Department where DepNo = ia.DepNo) DepName " + Environment.NewLine +
                                      "  from IASheetA ia " + Environment.NewLine +
                                      " where SheetNo = '" + vSheetNo.Trim() + "' ";
                        SqlDataAdapter daIASheetA = new SqlDataAdapter(vSQLStrTemp, connPrint);
                        vSQLStrTemp = "select SheetNoItems, SheetNo, Items, ic.ConsNo, ic.ConsName, ib.Quantity, ic.OriModelNo " + Environment.NewLine +
                                      "  from IASheetB ib left join IAConsumables ic on ic.ConsNo = ib.ConsNo " + Environment.NewLine +
                                      " where SheetNo = '" + vSheetNo.Trim() + "' ";
                        SqlDataAdapter daIASheetB = new SqlDataAdapter(vSQLStrTemp, connPrint);
                        connPrint.Open();
                        daIASheetA.Fill(dtIASheetA);
                        daIASheetB.Fill(dtIASheetB);

                        if ((dtIASheetA.Rows.Count > 0) && (dtIASheetB.Rows.Count > 0))
                        {
                            string vCompanyName = PF.GetValue(vConnStr, "select [Name] from [Custom] where [Code] = 'A000' and Types = 'O' ", "Name");
                            ReportDataSource sdsIASheetA = new ReportDataSource("IASheetA", dtIASheetA);
                            ReportDataSource sdsIASheetB = new ReportDataSource("IASheetB", dtIASheetB);

                            rvPrint.LocalReport.DataSources.Clear();
                            rvPrint.LocalReport.ReportPath = @"Report\IAConsumablesOutStoreP.rdlc";
                            rvPrint.LocalReport.DataSources.Add(sdsIASheetA);
                            rvPrint.LocalReport.DataSources.Add(sdsIASheetB);
                            rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                            rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "電腦耗材簽收單"));
                            rvPrint.LocalReport.Refresh();
                            plPrint.Visible = true;
                            plShowData.Visible = false;

                            //把單據設為 "已撥付"
                            //vSQLStrTemp = "update IASheetB set ItemStatus = '102' where SheetNo = '" + vSheetNo.Trim() + "' and ItemStatus = '000' ";
                            //PF.ExecSQL(vConnStr, vSQLStrTemp);
                            // 2022.08.30 改成明細逐筆變更狀態
                            using (SqlConnection connChangeTypeB=new SqlConnection())
                            {
                                connChangeTypeB.ConnectionString = vConnStr;
                                vSQLStrTemp = "select SheetNoItems from IASheetB where SheetNo = '" + vSheetNo + "' and ItemStatus = '000' ";
                                SqlCommand cmdChangeTypeB = new SqlCommand(vSQLStrTemp, connChangeTypeB);
                                connChangeTypeB.Open();
                                SqlDataReader drChangeTypeB = cmdChangeTypeB.ExecuteReader();
                                while (drChangeTypeB.Read())
                                {
                                    vSQLStrTemp = "update IASheetB set ItemStatus = '102' where SheetNoItems = '" + drChangeTypeB["SheetNoItems"].ToString().Trim() + "' ";
                                    PF.ExecSQL(vConnStr, vSQLStrTemp);
                                }
                            }
                            vSQLStrTemp = "update IASheetA set SheetStatus = '100' where SheetNo = '" + vSheetNo.Trim() + "' ";
                            PF.ExecSQL(vConnStr, vSQLStrTemp);
                            /* 2022.07.14 計算數量交給資料庫的 Trigger 做
                            for (int i = 0; i < dtIASheetB.Rows.Count; i++)
                            {
                                vConsNo = dtIASheetB.Rows[i]["ConsNo"].ToString().Trim();
                                PF.CalIAStoreQuantity(vConsNo, vConnStr, "");
                            } //*/
                            gridIASheetA_List.DataBind();
                            fvIASheetA_Detail.DataBind();
                            gridIASheetB_List.DataBind();
                            fvIASheetB_Detail.DataBind();
                        }
                    }
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// 整張出貨單作廢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbAbortA_List_Click(object sender, EventArgs e)
        {
            string vSQLStrTemp = "";
            string vSheetNo = "";
            //string vConsNo = "";
            Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
            if (eSheetNo_List != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                vSheetNo = eSheetNo_List.Text.Trim();
                vSQLStrTemp = "update IASheetB set ItemStatus = '999' where SheetNo = '" + vSheetNo.Trim() + "' and ItemStatus = '000' ";
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStrTemp);
                    vSQLStrTemp = "update IASheetA set SheetStatus = '999' where SheetNo = '" + vSheetNo.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStrTemp);
                    /* 2022.07.14 計算數量交給資料庫的 Trigger 做
                    using (SqlConnection connCalQty=new SqlConnection())
                    {
                        connCalQty.ConnectionString = vConnStr;
                        vSQLStrTemp = "select ConsNo from IASheetB where SheetNo = '" + vSheetNo.Trim() + "' ";
                        SqlCommand cmdCalQty = new SqlCommand(vSQLStrTemp, connCalQty);
                        connCalQty.Open();
                        SqlDataReader drCalQty = cmdCalQty.ExecuteReader();
                        while (drCalQty.Read())
                        {
                            vConsNo = drCalQty["ConsNo"].ToString().Trim();
                            PF.CalIAStoreQuantity(vConsNo, vConnStr, "");
                        }
                    } //*/
                    gridIASheetA_List.DataBind();
                    fvIASheetA_Detail.DataBind();
                    gridIASheetB_List.DataBind();
                    fvIASheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        protected void fvIASheetB_Detail_DataBound(object sender, EventArgs e)
        {
            string vProductDataURL = "";
            string vProductDataScript = "";
            switch (fvIASheetB_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Label eSheetNoItems_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_List");
                    if (eSheetNoItems_List != null)
                    {
                        Label eItemStatus_List = (Label)fvIASheetB_Detail.FindControl("eItemStatus_List");
                        string vItemStatus = eItemStatus_List.Text.Trim();
                        Button bbOutStoreB_List = (Button)fvIASheetB_Detail.FindControl("bbOutStoreB_List");
                        Button bbArrivedB_List = (Button)fvIASheetB_Detail.FindControl("bbArrivedB_List");
                        Button bbAbortB_List = (Button)fvIASheetB_Detail.FindControl("bbAbortB_List");
                        switch (vItemStatus)
                        {
                            default:
                                bbOutStoreB_List.Enabled = false;
                                bbArrivedB_List.Enabled = false;
                                bbAbortB_List.Enabled = false;
                                break;

                            case "000":
                                bbOutStoreB_List.Enabled = true;
                                bbArrivedB_List.Enabled = false;
                                bbAbortB_List.Enabled = true;
                                break;

                            case "102":
                                bbOutStoreB_List.Enabled = false;
                                bbArrivedB_List.Enabled = true;
                                bbAbortB_List.Enabled = true;
                                break;
                        }
                    }
                    break;
                case FormViewMode.Edit:
                    Label eSheetNoItems_Edit = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_Edit");
                    if (eSheetNoItems_Edit != null)
                    {
                        TextBox eConsNo_Edit = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_Edit");
                        //2023.01.12 修改
                        //vProductDataURL = "SearchProduct.aspx?TextBoxID=" + eConsNo_Edit.ClientID;
                        vProductDataURL = "SearchProduct.aspx?ConsNoID=" + eConsNo_Edit.ClientID;
                        //==================================================================================================================
                        vProductDataScript = "window.open('" + vProductDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        //eConsNo_Edit.Attributes["onClick"] = vProductDataScript;
                        Button bbGetConsNo_Edit = (Button)fvIASheetB_Detail.FindControl("bbGetConsNo_Edit");
                        bbGetConsNo_Edit.Attributes["onClick"] = vProductDataScript;
                    }
                    break;
                case FormViewMode.Insert:
                    Label eSheetNoItems_INS = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_INS");
                    if (eSheetNoItems_INS != null)
                    {
                        TextBox eConsNo_INS = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_INS");
                        //2023.01.12 修改
                        //vProductDataURL = "SearchProduct.aspx?TextBoxID=" + eConsNo_INS.ClientID;
                        vProductDataURL = "SearchProduct.aspx?ConsNoID=" + eConsNo_INS.ClientID;
                        //================================================================================================================
                        vProductDataScript = "window.open('" + vProductDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        Button bbGetConsNo_INS = (Button)fvIASheetB_Detail.FindControl("bbGetConsNo_INS");
                        bbGetConsNo_INS.Attributes["onClick"] = vProductDataScript;
                    }
                    break;
            }
        }

        /// <summary>
        /// 明細修改確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKB_Edit_Click(object sender, EventArgs e)
        {
            Label eSheetNoItems_Edit = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_Edit");
            if ((eSheetNoItems_Edit != null) && (eSheetNoItems_Edit.Text.Trim() != ""))
            {
                TextBox eConsNo_Edit = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_Edit");
                TextBox eQuantity_Edit = (TextBox)fvIASheetB_Detail.FindControl("eQuantity_Edit");
                TextBox eRemarkB_Edit = (TextBox)fvIASheetB_Detail.FindControl("eRemarkB_Edit");
                try
                {
                    string vSelectStr = "UPDATE IASheetB " + Environment.NewLine +
                                        "   SET ConsNo = @ConsNo, Quantity = @Quantity, RemarkB = @RemarkB, QtyMode = -1, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                        " WHERE (SheetNoItems = @SheetNoItems)";
                    dsIASheetB_Detail.UpdateCommand = vSelectStr;
                    dsIASheetB_Detail.UpdateParameters.Clear();
                    dsIASheetB_Detail.UpdateParameters.Add(new Parameter("ConsNo", DbType.String, (eConsNo_Edit.Text.Trim() != "") ? eConsNo_Edit.Text.Trim() : String.Empty));
                    dsIASheetB_Detail.UpdateParameters.Add(new Parameter("Quantity", DbType.Int32, (eQuantity_Edit.Text.Trim() != "") ? eQuantity_Edit.Text.Trim() : "0"));
                    dsIASheetB_Detail.UpdateParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB_Edit.Text.Trim() != "") ? eRemarkB_Edit.Text.Trim() : String.Empty));
                    dsIASheetB_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    dsIASheetB_Detail.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, eSheetNoItems_Edit.Text.Trim()));
                    dsIASheetB_Detail.Update();
                    // 2022.07.14 計算數量交給資料庫的 Trigger 做
                    //PF.CalIAStoreQuantity(eConsNo_Edit.Text.Trim(), vConnStr, "");
                    gridIASheetB_List.DataBind();
                    fvIASheetB_Detail.DataBind();

                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// 明細新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKB_INS_Click(object sender, EventArgs e)
        {
            Label eSheetNoItems_INS = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_INS");
            if (eSheetNoItems_INS != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eSheetNo_INS = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
                string vSheetNo = eSheetNo_INS.Text.Trim();
                string vSQLStr_Temp = "select max(Items) MaxItems from IASheetB where SheetNo = '" + vSheetNo + "' ";
                string vOldItem = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems");
                int vItems = (vOldItem != "") ? Int32.Parse(vOldItem) + 1 : 1;
                string vItemsStr = vItems.ToString("D4");
                string vSheetNoItems = vSheetNo + vItemsStr;
                TextBox eConsNo_INS = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_INS");
                TextBox eQuantity_INS = (TextBox)fvIASheetB_Detail.FindControl("eQuantity_INS");
                TextBox eRemarkB_INS = (TextBox)fvIASheetB_Detail.FindControl("eRemarkB_INS");
                try
                {
                    dsIASheetB_Detail.InsertCommand = "INSERT INTO IASheetB " + Environment.NewLine +
                                                      "       (SheetNoItems, SheetNo, Items, ConsNo, Quantity, RemarkB, QtyMode, ItemStatus, BuMan, BuDate) " + Environment.NewLine +
                                                      "VALUES (@SheetNoItems, @SheetNo, @Items, @ConsNo, @Quantity, @RemarkB, -1, '000', @BuMan, GetDate())";
                    dsIASheetB_Detail.InsertParameters.Clear();
                    dsIASheetB_Detail.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    dsIASheetB_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    dsIASheetB_Detail.InsertParameters.Add(new Parameter("Items", DbType.String, vItemsStr));
                    dsIASheetB_Detail.InsertParameters.Add(new Parameter("ConsNo", DbType.String, (eConsNo_INS.Text.Trim() != "") ? eConsNo_INS.Text.Trim() : String.Empty));
                    dsIASheetB_Detail.InsertParameters.Add(new Parameter("Quantity", DbType.Int32, (eQuantity_INS.Text.Trim() != "") ? eQuantity_INS.Text.Trim() : String.Empty));
                    dsIASheetB_Detail.InsertParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB_INS.Text.Trim() != "") ? eRemarkB_INS.Text.Trim() : String.Empty));
                    dsIASheetB_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    dsIASheetB_Detail.Insert();
                    // 2022.07.14 計算數量交給資料庫的 Trigger 做
                    //PF.CalIAStoreQuantity(eConsNo_INS.Text.Trim(), vConnStr, "");
                    gridIASheetB_List.DataBind();
                    fvIASheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// 明細刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDeleteB_List_Click(object sender, EventArgs e)
        {
            Label eSheetNoItems_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_List");
            if ((eSheetNoItems_List != null) && (eSheetNoItems_List.Text.Trim() != ""))
            {
                Label eConsNo_List = (Label)fvIASheetB_Detail.FindControl("eConsNo_List");
                string vConsNo_List = eConsNo_List.Text.Trim();
                try
                {
                    string vSQLStr_Del = "delete IASheetB where SheetNoItems = @SheetNoItems";
                    dsIASheetB_Detail.DeleteCommand = vSQLStr_Del;
                    dsIASheetB_Detail.DeleteParameters.Clear();
                    dsIASheetB_Detail.DeleteParameters.Add(new Parameter("SheetNoItems", DbType.String, eSheetNoItems_List.Text.Trim()));
                    dsIASheetB_Detail.Delete();
                    // 2022.07.14 計算數量交給資料庫的 Trigger 做
                    //PF.CalIAStoreQuantity(vConsNo_List, vConnStr, "");
                    gridIASheetB_List.DataBind();
                    fvIASheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }

        /// <summary>
        /// 明細撥付
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOutStoreB_List_Click(object sender, EventArgs e)
        {
            Label eSheetNoItems_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems_List != null)
            {
                Label eConsNo_List = (Label)fvIASheetB_Detail.FindControl("eConsNo_List");
                string vSheetNoItems = eSheetNoItems_List.Text.Trim();
                string vConsNo_List = eConsNo_List.Text.Trim();
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                using (SqlDataSource dsTemp = new SqlDataSource())
                {
                    dsTemp.ConnectionString = vConnStr;
                    dsTemp.UpdateCommand = "update IASheetB set ItemStatus = '101' where SheetNoItems = @SheetNoItems";
                    dsTemp.UpdateParameters.Clear();
                    dsTemp.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    dsTemp.Update();
                    // 2022.07.14 計算數量交給資料庫的 Trigger 做
                    //PF.CalIAStoreQuantity(vConsNo_List, vConnStr, "");
                    gridIASheetB_List.DataBind();
                    fvIASheetB_Detail.DataBind();
                }
            }
        }

        /// <summary>
        /// 明細簽收
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbArrivedB_List_Click(object sender, EventArgs e)
        {
            Label eSheetNoItems_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems_List != null)
            {
                string vSheetNoItems = eSheetNoItems_List.Text.Trim();
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                using (SqlDataSource dsTemp = new SqlDataSource())
                {
                    dsTemp.ConnectionString = vConnStr;
                    dsTemp.UpdateCommand = "update IASheetB set ItemStatus = '102' where SheetNoItems = @SheetNoItems";
                    dsTemp.UpdateParameters.Clear();
                    dsTemp.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    dsTemp.Update();
                    gridIASheetB_List.DataBind();
                    fvIASheetB_Detail.DataBind();
                }
            }
        }

        /// <summary>
        /// 明細作廢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbAbortB_List_Click(object sender, EventArgs e)
        {
            Label eSheetNoItems_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems_List != null)
            {
                Label eConsNo_List = (Label)fvIASheetB_Detail.FindControl("eConsNo_List");
                string vSheetNoItems = eSheetNoItems_List.Text.Trim();
                string vConsNo_List = eConsNo_List.Text.Trim();
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                using (SqlDataSource dsTemp = new SqlDataSource())
                {
                    dsTemp.ConnectionString = vConnStr;
                    dsTemp.UpdateCommand = "update IASheetB set ItemStatus = '999' where SheetNoItems = @SheetNoItems";
                    dsTemp.UpdateParameters.Clear();
                    dsTemp.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    dsTemp.Update();
                    // 2022.07.14 計算數量交給資料庫的 Trigger 做
                    //PF.CalIAStoreQuantity(vConsNo_List, vConnStr, "");
                    gridIASheetB_List.DataBind();
                    fvIASheetB_Detail.DataBind();
                }
            }
        }

        /// <summary>
        /// 結束報表預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plPrint.Visible = false;
        }
    }
}