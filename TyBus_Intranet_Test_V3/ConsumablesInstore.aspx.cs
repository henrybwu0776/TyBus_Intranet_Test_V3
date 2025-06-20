using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace TyBus_Intranet_Test_V3
{
    public partial class ConsumablesInstore : Page
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
                    //動態掛載日期輸入視窗
                    string vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDate_S_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBuDate_S_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDate_E_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBuDate_E_Search.Attributes["onClick"] = vDateScript;

                    if (!IsPostBack)
                    {
                        eBuDate_S_Search.Text = "";
                        eBuDate_E_Search.Text = "";
                        eAssignMan_Search.Text = "";
                        eAssignManName_Search.Text = "";
                        eDepNo_Search.Text = "";
                        eDepName_Search.Text = "";
                        eErrorMSG_A.Text = "";
                        eErrorMSG_A.Visible = false;
                        eErrorMSG_B.Text = "";
                        eErrorMSG_B.Visible = false;
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

        private string GetSelectStr()
        {
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_AssignMan = (eAssignMan_Search.Text.Trim() != "") ? "   and a.AssignMan = '" + eAssignMan_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuDate = ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() != "")) ? "   and a.BuDate between '" + PF.GetAD(eBuDate_S_Search.Text.Trim()) + "' and '" + PF.GetAD(eBuDate_E_Search.Text.Trim()) + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() == "")) ? "   and a.BuDate = '" + PF.GetAD(eBuDate_S_Search.Text.Trim()) + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() != "")) ? "   and a.BuDate = '" + PF.GetAD(eBuDate_E_Search.Text.Trim()) + "' " + Environment.NewLine : "";
            string vSelectStr = "select a.SheetNo, a.SupNo, c.[Name] SupName, a.SheetStatus, d.ClassTxt SheetStatus_C, " + Environment.NewLine +
                                "       a.DepNo, b.[Name] DepName, TotalAmount, convert(varchar(10), a.BuDate, 111) BuDate " + Environment.NewLine +
                                "  from ConsSheetA a left join[Custom] c on c.Code = a.SupNo and c.[Types] = 'S' " + Environment.NewLine +
                                "                    left join DBDICB d on d.ClassNo = a.SheetStatus and d.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus' " + Environment.NewLine +
                                "                    left join Department b on b.DepNo = a.DepNo " + Environment.NewLine +
                                " where a.SheetMode = 'SI' " + Environment.NewLine +
                                vWStr_DepNo +
                                vWStr_AssignMan +
                                vWStr_BuDate +
                                " order by a.BuDate DESC ";
            return vSelectStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSelectStr();
            sdsConsSheetA_List.SelectCommand = vSelectStr;
            gridConsSheetA_List.DataBind();
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName == "")
            {
                vDepName = eDepNo_Search.Text.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '" + vDepName + "' order by DepNo DESC";
                vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Search.Text = vDepNo;
            eDepName_Search.Text = vDepName;
        }

        protected void eAssignMan_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vAssignMan = eAssignMan_Search.Text.Trim();
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
            string vAssignManName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vAssignManName == "")
            {
                vAssignManName = eAssignMan_Search.Text.Trim();
                vSQLStr = "select top 1 EmpNo from Employee where [Name] like '%" + vAssignManName + "%' order by Assumeday DESC";
                vAssignMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eAssignMan_Search.Text = vAssignMan.Trim();
            eAssignManName_Search.Text = vAssignManName.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvConsSheetA_Detail_DataBound(object sender, EventArgs e)
        {
            DropDownList ddlTaxType_C;
            DropDownList ddlPayMode_C;
            TextBox eSupNo;
            TextBox eSupName;
            string vSupDataURL;
            string vSupDataScript;
            string vSQLStr_Temp = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvConsSheetA_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    ddlTaxType_C = (DropDownList)fvConsSheetA_Detail.FindControl("eTaxType_C_Edit");
                    if (ddlTaxType_C != null)
                    {
                        ddlTaxType_C.Items.Clear();
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB where FKey = '總務請購單      ConsSheetA      TAXTYPE' " + Environment.NewLine +
                                       " order by ClassNo";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                ddlTaxType_C.Items.Add(liTemp);
                            }
                        }
                        ddlPayMode_C = (DropDownList)fvConsSheetA_Detail.FindControl("ePayMode_C_Edit");
                        ddlPayMode_C.Items.Clear();
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB where FKey = '總務請購單      ConsSheetA      PayMode' " + Environment.NewLine +
                                       " order by ClassNo";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                ddlPayMode_C.Items.Add(liTemp);
                            }
                        }
                        eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_Edit");
                        eSupName = (TextBox)fvConsSheetA_Detail.FindControl("eSupName_Edit");
                        vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo.ClientID + "&SupNameID=" + eSupName.ClientID + "&SupKind=C";
                        vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eSupNo.Attributes["onClick"] = vSupDataScript;
                        Label eModifyDate = (Label)fvConsSheetA_Detail.FindControl("eModifyDateA_Edit");
                        Label eModifyMan = (Label)fvConsSheetA_Detail.FindControl("eModifyManA_Edit");
                        Label eModifyManName = (Label)fvConsSheetA_Detail.FindControl("eModifyManNameA_Edit");
                        eModifyDate.Text = DateTime.Today.ToShortDateString();
                        eModifyMan.Text = vLoginID;
                        eModifyManName.Text = vLoginName;
                    }
                    break;
                case FormViewMode.Insert:
                    ddlTaxType_C = (DropDownList)fvConsSheetA_Detail.FindControl("eTaxType_C_INS");
                    if (ddlTaxType_C != null)
                    {
                        ddlTaxType_C.Items.Clear();
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB where FKey = '總務請購單      ConsSheetA      TAXTYPE' " + Environment.NewLine +
                                       " order by ClassNo";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                ddlTaxType_C.Items.Add(liTemp);
                            }
                        }
                        ddlPayMode_C = (DropDownList)fvConsSheetA_Detail.FindControl("ePayMode_C_INS");
                        ddlPayMode_C.Items.Clear();
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB where FKey = '總務請購單      ConsSheetA      PayMode' " + Environment.NewLine +
                                       " order by ClassNo";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                ddlPayMode_C.Items.Add(liTemp);
                            }
                        }
                        eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_INS");
                        eSupName = (TextBox)fvConsSheetA_Detail.FindControl("eSupName_INS");
                        vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo.ClientID + "&SupNameID=" + eSupName.ClientID + "&SupKind=C";
                        vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eSupNo.Attributes["onClick"] = vSupDataScript;
                        Label eBuDate = (Label)fvConsSheetA_Detail.FindControl("eBuDateA_INS");
                        Label eBuMan = (Label)fvConsSheetA_Detail.FindControl("eBuManA_INS");
                        Label eBuManName = (Label)fvConsSheetA_Detail.FindControl("eBuManNameA_INS");
                        eBuDate.Text = DateTime.Today.ToShortDateString();
                        eBuMan.Text = vLoginID;
                        eBuManName.Text = vLoginName;
                        Label eSheetStatus = (Label)fvConsSheetA_Detail.FindControl("eSheetStatus_INS");
                        eSheetStatus.Text = "000";
                        Label eSheetStatus_C = (Label)fvConsSheetA_Detail.FindControl("eSheetStatus_C_INS");
                        vSQLStr_Temp = "select ClassTxt from DBDICB where FKey = '總務耗材進出單  fmConsSheetA    SheetStatus' and ClassNo = '000' ";
                        eSheetStatus_C.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "ClassTxt");
                        Label eStatusDate = (Label)fvConsSheetA_Detail.FindControl("eStatusDate_INS");
                        eStatusDate.Text = DateTime.Today.ToShortDateString();
                    }
                    break;
            }
        }

        protected void bbOK_A_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_Edit");
            string vSQLStr_Temp = "";
            int vTempINT;
            DateTime vToday = DateTime.Today;
            if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    //新增異動記錄
                    string vOptionNote = "異動總務耗材進貨單" + Environment.NewLine +
                                         "單號：[ " + eSheetNo.Text.Trim() + " ]" + Environment.NewLine +
                                         "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                         "異動種類：修改" + Environment.NewLine +
                                         "ConsumablesInstore.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionNote);
                    //複製一份到歷史記錄檔
                    string vH_IndexFirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_IndexFirstCode + "%' ";
                    string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    string vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_IndexFirstCode, ""), out vTempINT) ? vH_IndexFirstCode + vTempINT.ToString("D4") : vH_IndexFirstCode + "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, ActionDate, ActionMan, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, " + Environment.NewLine +
                                   "       TaxRate, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan) " + Environment.NewLine +
                                   "select '" + vNewIndex + "', 'EDIT', GetDate, '" + vLoginID + "',  SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, " + Environment.NewLine +
                                   "       TaxRate, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + eSheetNo.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入修改資料
                    string vSheetNo = eSheetNo.Text.Trim();
                    TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
                    TextBox eAssignMan = (TextBox)fvConsSheetA_Detail.FindControl("eAssignMan_Edit");
                    TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_Edit");
                    TextBox eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_Edit");
                    TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_Edit");
                    Label eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmountA_Edit");
                    Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_Edit");
                    Label eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_Edit");
                    Label eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_Edit");
                    Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_Edit");
                    Label ePayDate = (Label)fvConsSheetA_Detail.FindControl("ePayDate_Edit");
                    vSQLStr_Temp = "update ConsSheetA " + Environment.NewLine +
                                   "   set DepNo = @DepNo, AssignMan = @AssignMan, SupNo = @SupNo, Amount = @Amount, TaxType = @TaxType, TaxRate = @TaxRate, " + Environment.NewLine +
                                   "       TaxAMT = @TaxAMT, TotalAmount = @TotalAmount, PayDate = @PayDate, PayMode = @PayMode, RemarkA = @RemarkA, " + Environment.NewLine +
                                   "       ModifyDate = GetDate(), ModifyMan = @ModifyMan " + Environment.NewLine +
                                   " where SheetNo = @SheetNo ";
                    double vTempFloat;
                    DateTime vTempDate;
                    sdsConsSheetA_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.UpdateParameters.Clear();
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan.Text.Trim() != "") ? eAssignMan.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SupNo", DbType.String, (eSupNo.Text.Trim() != "") ? eSupNo.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("Amount", DbType.Double, double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TaxType", DbType.String, (eTaxType.Text.Trim() != "") ? eTaxType.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TaxRate", DbType.Double, double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TaxAMT", DbType.Double, double.TryParse(eTaxAMT.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, double.TryParse(eTotalAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("PayDate", DbType.Date, DateTime.TryParse(ePayDate.Text.Trim(), out vTempDate) ? PF.GetAD(vTempDate.ToShortDateString()) : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("PayMode", DbType.String, (ePayMode.Text.Trim() != "") ? ePayMode.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA.Text.Trim() != "") ? eRemarkA.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetA_Detail.Update();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
            if (eDepNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_Edit");
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo DESC";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo.Text = vDepNo;
                eDepName.Text = vDepName;
            }
        }

        protected void eAssignMan_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eAssignMan = (TextBox)fvConsSheetA_Detail.FindControl("eAssignMan_Edit");
            if (eAssignMan != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eAssignManName = (Label)fvConsSheetA_Detail.FindControl("eAssignManName_Edit");
                string vAssignMan = eAssignMan.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
                string vAssignManName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vAssignManName == "")
                {
                    vAssignManName = eAssignMan.Text.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] like '%" + vAssignManName + "%' order by Assumeday DESC";
                    vAssignMan = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eAssignMan.Text = vAssignMan;
                eAssignManName.Text = vAssignManName;
            }
        }

        protected void eTaxType_C_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("eTaxType_C_Edit");
            if (ddlTaxType != null)
            {
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_Edit");
                eTaxType.Text = ddlTaxType.SelectedValue.Trim();
                TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_Edit");
                if ((ddlTaxType.SelectedValue.Trim() == "1") || (ddlTaxType.SelectedValue.Trim() == "2"))
                {
                    eTaxRate.Text = "0.05";
                }
                else
                {
                    eTaxRate.Text = "0.0";
                }
            }
        }

        protected void eTaxRate_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_Edit");
            double vTempFloat;
            double vAmount;
            double vTaxRate;
            double vTaxAMT;
            double vTotalAmount;
            if (eTaxRate != null)
            {
                vTaxRate = Double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmountA_Edit");
                vAmount = double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_Edit");
                vTaxAMT = double.TryParse(eTaxAMT.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_Edit");
                vTotalAmount = double.TryParse(eTotalAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_Edit");
                switch (eTaxType.Text.Trim())
                {
                    case "1":
                        vTaxAMT = Math.Round(vAmount * vTaxRate, 0, MidpointRounding.AwayFromZero);
                        vTotalAmount = vAmount + vTaxAMT;
                        break;
                    case "2":
                        vTaxAMT = Math.Round(vAmount / (1 + vTaxRate), 0, MidpointRounding.AwayFromZero);
                        vTotalAmount = vAmount;
                        vAmount = vTotalAmount - vTaxAMT;
                        break;

                    case "3":
                    case "4":
                        vTaxRate = 0.0;
                        vTaxAMT = 0.0;
                        vTotalAmount = vAmount;
                        break;
                }
                eTaxAMT.Text = vTaxAMT.ToString();
                eAmount.Text = vAmount.ToString();
                eTotalAmount.Text = vTotalAmount.ToString();
                eTaxRate.Text = vTaxRate.ToString();
            }
        }

        protected void ePayMode_C_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlPayMode = (DropDownList)fvConsSheetA_Detail.FindControl("ePayMode_C_Edit");
            if (ddlPayMode != null)
            {
                Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_Edit");
                ePayMode.Text = ddlPayMode.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 確定新增表頭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKA_INS_Click(object sender, EventArgs e)
        {
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_INS");
            string vSheetMode = "SI"; //進貨單的代碼 SI
            string vSQLStr_Temp = "";
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                TextBox eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_INS");
                Label eBuMan = (Label)fvConsSheetA_Detail.FindControl("eBuManA_INS");
                Label eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmountA_INS");
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_INS");
                TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_INS");
                Label eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_INS");
                Label eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_INS");
                Label eSheetStatus = (Label)fvConsSheetA_Detail.FindControl("eSheetStatus_INS");
                Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_INS");
                TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_INS");
                TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
                TextBox eAssignMan = (TextBox)fvConsSheetA_Detail.FindControl("eAssignMan_INS");
                DateTime vToday = DateTime.Today;
                string vFirstWord = (vToday.Year - 1911).ToString("D3") + vToday.Month.ToString("D2") + vSheetMode.Trim();
                vSQLStr_Temp = "select max(SheetNo) MaxNo from ConsSheetA where SheetNo like '" + vFirstWord + "%' ";
                string vOldSheetNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                int vTempINT;
                double vTempFloat;
                string vNewIndexStr = Int32.TryParse(vOldSheetNo.Replace(vFirstWord, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                string vSheetNo = vFirstWord + vNewIndexStr;
                vSQLStr_Temp = "insert into ConsSheet " + Environment.NewLine +
                               "      (SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                               "       SheetStatus, StatusDate, PayMode, RemarkA, DepNo, AssignMan) " + Environment.NewLine +
                               "values(@SheetNo, @SheetMode, @SupNo, GetDate(), @BuMan, @Amount, @TaxType, @TaxRate, @TaxAMT, @TotalAmount, " + Environment.NewLine +
                               "       @SheetStatus, GetDate(), @PayMode, @RemarkA, @DepNo, @AssignMan) ";
                sdsConsSheetA_Detail.InsertCommand = vSQLStr_Temp;
                sdsConsSheetA_Detail.InsertParameters.Clear();
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetMode", DbType.String, vSheetMode));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SupNo", DbType.String, (eSupNo.Text.Trim() != "") ? eSupNo.Text.Trim() : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("Amount", DbType.Double, double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TaxType", DbType.String, (eTaxType.Text.Trim() != "") ? eTaxType.Text.Trim() : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TaxRate", DbType.Double, double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TaxAMT", DbType.Double, double.TryParse(eTaxAMT.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, double.TryParse(eTotalAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetStatus", DbType.String, (eSheetStatus.Text.Trim() != "") ? eSheetStatus.Text.Trim() : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("PayMode", DbType.String, (ePayMode.Text.Trim() != "") ? ePayMode.Text.Trim() : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA.Text.Trim() != "") ? eRemarkA.Text.Trim() : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan.Text.Trim() != "") ? eAssignMan.Text.Trim() : String.Empty));
                sdsConsSheetA_Detail.Insert();
                gridConsSheetA_List.DataBind();
                fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                fvConsSheetA_Detail.DataBind();
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
            if (eDepNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_INS");
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo DESC";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo.Text = vDepNo;
                eDepName.Text = vDepName;
            }
        }

        protected void eAssignMan_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eAssignMan = (TextBox)fvConsSheetA_Detail.FindControl("eAssignMan_INS");
            if (eAssignMan != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eAssignManName = (Label)fvConsSheetA_Detail.FindControl("eAssignManName_INS");
                string vAssignMan = eAssignMan.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
                string vAssignManName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vAssignManName == "")
                {
                    vAssignManName = eAssignMan.Text.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] like '%" + vAssignManName + "%' order by Assumeday DESC";
                    vAssignMan = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eAssignMan.Text = vAssignMan;
                eAssignManName.Text = vAssignManName;
            }
        }

        protected void eTaxType_C_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("eTaxType_C_INS");
            if (ddlTaxType != null)
            {
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_INS");
                eTaxType.Text = ddlTaxType.SelectedValue.Trim();
                TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_INS");
                if ((ddlTaxType.SelectedValue.Trim() == "1") || (ddlTaxType.SelectedValue.Trim() == "2"))
                {
                    eTaxRate.Text = "0.05";
                }
                else
                {
                    eTaxRate.Text = "0.0";
                }
            }
        }

        protected void eTaxRate_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_INS");
            double vTempFloat;
            double vAmount;
            double vTaxRate;
            double vTaxAMT;
            double vTotalAmount;
            if (eTaxRate != null)
            {
                vTaxRate = Double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmountA_INS");
                vAmount = double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_INS");
                vTaxAMT = double.TryParse(eTaxAMT.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_INS");
                vTotalAmount = double.TryParse(eTotalAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_INS");
                switch (eTaxType.Text.Trim())
                {
                    case "1":
                        vTaxAMT = Math.Round(vAmount * vTaxRate, 0, MidpointRounding.AwayFromZero);
                        vTotalAmount = vAmount + vTaxAMT;
                        break;
                    case "2":
                        vTaxAMT = Math.Round(vAmount / (1 + vTaxRate), 0, MidpointRounding.AwayFromZero);
                        vTotalAmount = vAmount;
                        vAmount = vTotalAmount - vTaxAMT;
                        break;

                    case "3":
                    case "4":
                        vTaxRate = 0.0;
                        vTaxAMT = 0.0;
                        vTotalAmount = vAmount;
                        break;
                }
                eTaxAMT.Text = vTaxAMT.ToString();
                eAmount.Text = vAmount.ToString();
                eTotalAmount.Text = vTotalAmount.ToString();
                eTaxRate.Text = vTaxRate.ToString();
            }
        }

        protected void ePayMode_C_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlPayMode = (DropDownList)fvConsSheetA_Detail.FindControl("ePayMode_C_Edit");
            if (ddlPayMode != null)
            {
                Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_Edit");
                ePayMode.Text = ddlPayMode.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 列印入庫單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_List_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 全部入庫
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseA_List_Click(object sender, EventArgs e)
        {
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSQLStr_Temp = "update ConsSheetB set ItemStatus = '001' where ItemStatus = '000' and SheetNo = '" + eSheetNo.Text.Trim() + "' ";
                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                OpenData();
            }
        }

        protected void bbDeleteA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            string vSQLStr_Temp;
            string vSheetNoItems;
            string vH_IndexFirstCode;
            string vMaxIndex;
            string vNewIndex;
            int vTempINT;
            int vItems;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                try
                {
                    //新增異動記錄
                    string vOptionNote = "異動總務耗材進貨單" + Environment.NewLine +
                                         "單號：[ " + eSheetNo.Text.Trim() + " ]" + Environment.NewLine +
                                         "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                         "異動種類：刪除" + Environment.NewLine +
                                         "ConsumablesInstore.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionNote);
                    //複製一份到歷史記錄檔
                    vH_IndexFirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_IndexFirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_IndexFirstCode, ""), out vTempINT) ? vH_IndexFirstCode + vTempINT.ToString("D4") : vH_IndexFirstCode + "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, ActionDate, ActionMan, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, " + Environment.NewLine +
                                   "       TaxRate, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan) " + Environment.NewLine +
                                   "select '" + vNewIndex + "', 'DEL', GetDate(), '" + vLoginID + "',  SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, " + Environment.NewLine +
                                   "       TaxRate, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + eSheetNo.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    vSQLStr_Temp = "select count(Items) RCount from ConsSheetB where SheetNo = '" + eSheetNo.Text.Trim() + "' ";
                    if ((Int32.TryParse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount"), out vTempINT) == true) && (vTempINT > 0))
                    {
                        //有明細檔，先刪除明細
                        vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_IndexFirstCode + "%' ";
                        vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                        vItems = Int32.TryParse(vMaxIndex.Replace(vH_IndexFirstCode, ""), out vTempINT) ? vTempINT : 0;
                        using (SqlConnection connDelSheetB = new SqlConnection(vConnStr))
                        {
                            vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = '" + eSheetNo.Text.Trim() + "' order by Items ";
                            SqlCommand cmdDelSheetB = new SqlCommand(vSQLStr_Temp, connDelSheetB);
                            connDelSheetB.Open();
                            SqlDataReader drDelSheetB = cmdDelSheetB.ExecuteReader();
                            while (drDelSheetB.Read())
                            {
                                vItems++;
                                vNewIndex = vH_IndexFirstCode + vItems.ToString("D4");
                                vSheetNoItems = drDelSheetB["SheetNoItems"].ToString().Trim();
                                vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                               "      (H_Index, ActionMode, ActionDate, ActionMan, SheetNoItems, SheetNo, Items, ConsNo, ConsUnit, " + Environment.NewLine +
                                               "       Price, Quantity, Amount, RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate) " + Environment.NewLine +
                                               "select '" + vNewIndex + "', 'DEL', GetDate(), '" + vLoginID + "', SheetNoItems, SheetNo, Items, ConsNo, ConsUnit, " + Environment.NewLine +
                                               "       Price, Quantity, Amount, RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate " + Environment.NewLine +
                                               "  from ConsSheetB " + Environment.NewLine +
                                               " where SheetNoItems = '" + vSheetNoItems + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            }
                        }
                        vSQLStr_Temp = "delete ConsSheetB where SheetNo = '" + eSheetNo.Text.Trim() + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    }
                    vSQLStr_Temp = "delete ConsSheetA where SheetNo = @SheetNo";
                    sdsConsSheetA_Detail.DeleteCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.DeleteParameters.Clear();
                    sdsConsSheetA_Detail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, eSheetNo.Text.Trim()));
                    sdsConsSheetA_Detail.Delete();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        protected void fvConsSheetB_Detail_DataBound(object sender, EventArgs e)
        {
            string vConsDataURL;
            string vConsDataScript;
            Label eSheetNoItems;
            TextBox eConsNo;
            TextBox eConsName;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvConsSheetB_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_Edit");
                    if (eSheetNoItems != null)
                    {
                        eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
                        eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_Edit");
                        vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsPriceID=" + eConsName.ClientID;
                        vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eConsNo.Attributes["onClick"] = vConsDataScript;
                    }
                    break;
                case FormViewMode.Insert:
                    eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_INS");
                    if (eSheetNoItems != null)
                    {
                        eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                        eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_INS");
                        vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsPriceID=" + eConsName.ClientID;
                        vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eConsNo.Attributes["onClick"] = vConsDataScript;
                    }
                    break;
                default:
                    break;
            }
        }

        protected void bbOKB_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            string vSQLStr_Temp;
            string vH_IndexFirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vMaxIndex;
            string vNewIndex;
            string vSheetNoItems;
            double vTempFloat;
            int vTempINT;
            int vItems;
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_Edit");
            if (eSheetNoItems != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_Edit");
                    Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_Edit");
                    //寫入異動記錄
                    string vOptionNote = "異動總務耗材進貨單明細" + Environment.NewLine +
                                         "單號：[ " + eSheetNo.Text.Trim() + " ]" + Environment.NewLine +
                                         "項次：[ " + eItems.Text.Trim() + " ]" + Environment.NewLine +
                                         "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                         "異動種類：修改" + Environment.NewLine +
                                         "ConsumablesInstore.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionNote);
                    //複製原資料到歷史檔
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_IndexFirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vItems = Int32.TryParse(vMaxIndex.Replace(vH_IndexFirstCode, ""), out vTempINT) ? vTempINT + 1 : 1;
                    vNewIndex = vH_IndexFirstCode + vItems.ToString("D4");
                    vSheetNoItems = eSheetNoItems.Text.Trim();
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, ActionDate, ActionMan, SheetNoItems, SheetNo, Items, ConsNo, ConsUnit, " + Environment.NewLine +
                                   "       Price, Quantity, Amount, RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate) " + Environment.NewLine +
                                   "select '" + vNewIndex + "', 'DEL', GetDate(), '" + vLoginID + "', SheetNoItems, SheetNo, Items, ConsNo, ConsUnit, " + Environment.NewLine +
                                   "       Price, Quantity, Amount, RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //開始更新異動資料
                    TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
                    TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_Edit");
                    TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_Edit");
                    Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmountB_Edit");
                    TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_Edit");
                    Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_Edit");
                    vSQLStr_Temp = "update ConsSheetB " + Environment.NewLine +
                                   "   set ConsNo = @ConsNo, Price = @Price, Quantity = @Quantity, Amount = @Amount, RemarkB = @RemarkB, " + Environment.NewLine +
                                   "       QtyMode = 1, ConsUnit = @ConsUnit, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                   " where SheetNoItems = @SheetNoItems ";
                    sdsConsSheetB_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.UpdateParameters.Clear();
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ConsNo", DbType.String, (eConsNo.Text.Trim() != "") ? eConsNo.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Price", DbType.Double, double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Quantity", DbType.Double, double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Amount", DbType.Double, double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB.Text.Trim() != "") ? eRemarkB.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ConsUnit", DbType.String, (eConsUnit.Text.Trim() != "") ? eConsUnit.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.Update();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        protected void eConsName_Edit_TextChanged(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
            if ((eConsNo != null) && (eConsNo.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vConsNo = eConsNo.Text.Trim();
                double vTempFloat;
                TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_Edit");
                Label eStockQty = (Label)fvConsSheetB_Detail.FindControl("eStockQty_Edit");
                Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_Edit");
                Label eConsUnit_C = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_C_Edit");
                string vSQLStr_Temp = "select c.StockQty, c.LastInPrice, c.ConsUnit, d.ClassTxt ConsUnit_C " + Environment.NewLine +
                                      "  from Consumables c left join DBDICB d on d.ClassNo = c.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                                      " where c.ConsNo = @ConsNo";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                    connTemp.Open();
                    cmdTemp.Parameters.Clear();
                    cmdTemp.Parameters.Add(new SqlParameter("ConsNo", vConsNo));
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        ePrice.Text = Double.TryParse(drTemp["LastInPrice"].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                        eStockQty.Text = Double.TryParse(drTemp["StockQty"].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                        eConsUnit.Text = drTemp["ConsUnit"].ToString().Trim();
                        eConsUnit_C.Text = drTemp["ConsUnit_C"].ToString().Trim();
                    }
                }
            }
        }

        protected void eQuantity_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_Edit");
            if (eQuantity != null)
            {
                TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_Edit");
                Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmountB_edit");
                double vTempFloat;
                double vQuantity = double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                double vPrice = double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                eAmount.Text = (vPrice * vQuantity).ToString();
            }
        }

        protected void bbOKB_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_INS");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSheetNo = eSheetNo.Text.Trim();
                string vSheetNoItems;
                string vSQLStr_Temp;
                string vMaxIndex;
                double vTempFloat;
                int vTempINT;
                int vItems;
                try
                {
                    vSQLStr_Temp = "select max(Items) MaxIndex from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vItems = Int32.TryParse(vMaxIndex, out vTempINT) ? vTempINT + 1 : 1;
                    vSheetNoItems = vSheetNo + vItems.ToString("D4");
                    //寫入異動記錄
                    string vOptionNote = "異動總務耗材進貨單明細" + Environment.NewLine +
                                         "單號：[ " + eSheetNo.Text.Trim() + " ]" + Environment.NewLine +
                                         "項次：[ " + vItems.ToString("D4") + " ]" + Environment.NewLine +
                                         "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                         "異動種類：新增" + Environment.NewLine +
                                         "ConsumablesInstore.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionNote);
                    //寫入明細檔
                    TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                    TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_INS");
                    TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_INS");
                    Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmountB_INS");
                    TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_INS");
                    Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                    vSQLStr_Temp = "insert into ConSheetB " + Environment.NewLine +
                                   "      (SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ConsUnit) " + Environment.NewLine +
                                   "values " + Environment.NewLine +
                                   "      (@SheetNoItems, @SheetNo, @Items, @ConsNo, @Price, @Quantity, @Amount, @RemarkB, 1, '000', @BuMan, GetDate(), @ConsUnit) ";
                    sdsConsSheetB_Detail.InsertCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.InsertParameters.Clear();
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Items", DbType.String, vItems.ToString("D4")));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ConsNo", DbType.String, (eConsNo.Text.Trim() != "") ? eConsNo.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Price", DbType.Double, double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("quantity", DbType.Double, double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Amount", DbType.Double, double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB.Text.Trim() != "") ? eRemarkB.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ConsUnit", DbType.String, (eConsUnit.Text.Trim() != "") ? eConsUnit.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.Insert();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = false;
                }
            }
        }

        protected void eConsName_INS_TextChanged(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
            if ((eConsNo != null) && (eConsNo.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vConsNo = eConsNo.Text.Trim();
                double vTempFloat;
                TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_INS");
                Label eStockQty = (Label)fvConsSheetB_Detail.FindControl("eStockQty_INS");
                Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                Label eConsUnit_C = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_C_INS");
                string vSQLStr_Temp = "select c.StockQty, c.LastInPrice, c.ConsUnit, d.ClassTxt ConsUnit_C " + Environment.NewLine +
                                      "  from Consumables c left join DBDICB d on d.ClassNo = c.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                                      " where c.ConsNo = @ConsNo";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                    connTemp.Open();
                    cmdTemp.Parameters.Clear();
                    cmdTemp.Parameters.Add(new SqlParameter("ConsNo", vConsNo));
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        ePrice.Text = Double.TryParse(drTemp["LastInPrice"].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                        eStockQty.Text = Double.TryParse(drTemp["StockQty"].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                        eConsUnit.Text = drTemp["ConsUnit"].ToString().Trim();
                        eConsUnit_C.Text = drTemp["ConsUnit_C"].ToString().Trim();
                    }
                }
            }
        }

        protected void eQuantity_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_INS");
            if (eQuantity != null)
            {
                TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_INS");
                Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmountB_INS");
                double vTempFloat;
                double vQuantity = double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                double vPrice = double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                eAmount.Text = (vPrice * vQuantity).ToString();
            }
        }
    }
}