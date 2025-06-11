using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ConsSheet_Order : Page
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
                        plReport.Visible = false;
                        plSearch.Visible = true;
                        plShowData_A.Visible = true;
                        plShowDetail.Visible = true;
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

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo_Search.Text.Trim();
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            string vDepName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName.Trim() == "")
            {
                vDepName = eDepNo_Search.Text.Trim();
                vSQLStr = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo ";
                vDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Search.Text = vDepNo.Trim();
            eDepName_Search.Text = vDepName.Trim();
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
                vSQLStr = "select top 1 EmpNo from Employee where [Name] like '%" + vAssignManName + "%' order by EmpNo DESC";
                vAssignMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eAssignMan_Search.Text = vAssignMan.Trim();
            eAssignManName_Search.Text = vAssignManName.Trim();
        }

        private string GetSearchStr()
        {
            DateTime vDate_S;
            DateTime vDate_E;
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   AND a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_AssignMan = (eAssignMan_Search.Text.Trim() != "") ? "   AND a.AssignMan = '" + eAssignMan_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuDate = (DateTime.TryParse(eBuDate_S_Search.Text.Trim(), out vDate_S) && DateTime.TryParse(eBuDate_E_Search.Text.Trim(), out vDate_E)) ?
                                  "   AND a.BuDate between '" + PF.TransDateString(vDate_S, "B") + "' and '" + PF.TransDateString(vDate_E, "B") + "' " + Environment.NewLine :
                                  (!DateTime.TryParse(eBuDate_S_Search.Text.Trim(), out vDate_S) && DateTime.TryParse(eBuDate_E_Search.Text.Trim(), out vDate_E)) ?
                                  "   AND a.BuDate = '" + PF.TransDateString(vDate_E, "B") + "' " + Environment.NewLine :
                                  DateTime.TryParse(eBuDate_S_Search.Text.Trim(), out vDate_S) && !DateTime.TryParse(eBuDate_E_Search.Text.Trim(), out vDate_E) ?
                                  "   AND a.BuDate = '" + PF.TransDateString(vDate_S, "B") + "' " + Environment.NewLine : "";
            string vResultStr = "SELECT a.SheetNo, a.SheetNote, a.DepNo, d.NAME AS DepName, a.AssignMan, e.NAME AS AssignManName, " + Environment.NewLine +
                                "       CONVERT (varchar(10), a.BuDate, 111) AS BuDate, a.TotalAmount, " + Environment.NewLine +
                                "       CASE WHEN a.SheetStatus = '0' THEN '已開單' " + Environment.NewLine +
                                "            WHEN a.SheetStatus = '1' THEN '處理中' " + Environment.NewLine +
                                "            WHEN a.SheetStatus = '8' THEN '已作廢' " + Environment.NewLine +
                                "            WHEN a.SheetStatus = '9' THEN '已結案' END AS SheetStatus_C " + Environment.NewLine +
                                "  FROM ConsSheetA AS a LEFT OUTER JOIN DEPARTMENT AS d ON d.DEPNO = a.DepNo " + Environment.NewLine +
                                "                       LEFT OUTER JOIN EMPLOYEE AS e ON e.EMPNO = a.AssignMan " + Environment.NewLine +
                                " WHERE isnull(a.SheetMode, '') = 'BS' " + Environment.NewLine +
                                vWStr_DepNo +
                                vWStr_AssignMan +
                                vWStr_BuDate +
                                " ORDER BY SheetNo ";
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSearchStr();
            gridConsSheetA_List.DataSourceID = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connSearch = new SqlConnection(vConnStr))
            {
                DataTable dtSearch = new DataTable();
                SqlDataAdapter daSearch = new SqlDataAdapter(vSelectStr, connSearch);
                connSearch.Open();
                daSearch.Fill(dtSearch);
                gridConsSheetA_List.DataSource = dtSearch;
                gridConsSheetA_List.DataBind();
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 主檔資料完成繫結
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void fvConsSheetA_Detail_DataBound(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            DropDownList ddlTaxType;
            Label eTaxType;
            DropDownList ddlPayMode;
            Label ePayMode;
            TextBox eSupNo;
            TextBox eSupName;
            string vSupDataURL;
            string vSupDataScript;
            string vSQLStr;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvConsSheetA_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("eTaxType_C_Edit");
                    if (ddlTaxType != null)
                    {
                        ddlTaxType.Items.Clear();
                        eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_Edit");
                        vSQLStr = "select cast('' as varchar) ClassNo, Cast('' as varchar) ClassTxt " + Environment.NewLine +
                                  " union all " + Environment.NewLine +
                                  "select ClassNo, ClassTxt from DBDICB where FKey = '總務請購單      ConsSheetA      TAXTYPE' " + Environment.NewLine +
                                  " order by ClassNo ";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                ddlTaxType.Items.Add(liTemp);
                            }
                        }
                        ddlTaxType.SelectedIndex = ddlTaxType.Items.IndexOf(ddlTaxType.Items.FindByValue(eTaxType.Text.Trim()));
                        ddlPayMode = (DropDownList)fvConsSheetA_Detail.FindControl("ePayMode_C_Edit");
                        ddlPayMode.Items.Clear();
                        vSQLStr = "select cast('' as varchar) ClassNo, Cast('' as varchar) ClassTxt " + Environment.NewLine +
                                  " union all " + Environment.NewLine +
                                  "select ClassNo, ClassTxt from DBDICB where FKey = '總務請購單      ConsSheetA      PayMode' " + Environment.NewLine +
                                  " order by ClassNo ";
                        ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_Edit");
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                ddlPayMode.Items.Add(liTemp);
                            }
                        }
                        ddlPayMode.SelectedIndex = ddlPayMode.Items.IndexOf(ddlPayMode.Items.FindByValue(ePayMode.Text.Trim()));
                        eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_Edit");
                        eSupName = (TextBox)fvConsSheetA_Detail.FindControl("eSupName_Edit");
                        vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo.ClientID + "&SupNameID=" + eSupName.ClientID + "&SupKind=C";
                        vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eSupNo.Attributes["onClick"] = vSupDataScript;
                    }
                    break;
                case FormViewMode.Insert:
                    ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("eTaxType_C_INS");
                    if (ddlTaxType != null)
                    {
                        vSQLStr = "select cast('' as varchar) ClassNo, Cast('' as varchar) ClassTxt " + Environment.NewLine +
                                  " union all " + Environment.NewLine +
                                  "select ClassNo, ClassTxt from DBDICB where FKey = '總務請購單      ConsSheetA      TAXTYPE' " + Environment.NewLine +
                                  " order by ClassNo ";
                        ddlTaxType.Items.Clear();
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                ddlTaxType.Items.Add(liTemp);
                            }
                        }
                    }
                    ddlPayMode = (DropDownList)fvConsSheetA_Detail.FindControl("ePayMode_C_INS");
                    ddlPayMode.Items.Clear();
                    vSQLStr = "select cast('' as varchar) ClassNo, Cast('' as varchar) ClassTxt " + Environment.NewLine +
                              " union all " + Environment.NewLine +
                              "select ClassNo, ClassTxt from DBDICB where FKey = '總務請購單      ConsSheetA      PayMode' " + Environment.NewLine +
                              " order by ClassNo ";
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        while (drTemp.Read())
                        {
                            ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                            ddlPayMode.Items.Add(liTemp);
                        }
                    }
                    eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_INS");
                    eSupName = (TextBox)fvConsSheetA_Detail.FindControl("eSupName_INS");
                    vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo.ClientID + "&SupNameID=" + eSupName.ClientID + "&SupKind=C";
                    vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                    eSupNo.Attributes["onClick"] = vSupDataScript;
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 修改主檔資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            DateTime vToday = DateTime.Today;
            string vSQLStr_Temp;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vMaxIndex;
            string vNewIndex;
            double vTempFloat;
            int vTempINT;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNo_Edit");
            if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    string vSheetNo = eSheetNo.Text.Trim();
                    //寫入更新記錄
                    string vOptionString = "修改總務耗材請購單" + Environment.NewLine +
                                           "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：修改" + Environment.NewLine +
                                           "ConsSheet_Order.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份主檔資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                                   "       SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyDate, ModifyMan, DepNo, AssignMan, ActiveDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                                   "       SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyDate, ModifyMan, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    TextBox eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_Edit");
                    string vSupNo = eSupNo.Text.Trim();
                    TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
                    string vDepNo = eDepNo.Text.Trim();
                    TextBox eAssignMan = (TextBox)fvConsSheetA_Detail.FindControl("eAssignMan_Edit");
                    string vAssignMan = eAssignMan.Text.Trim();
                    Label eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmount_Edit");
                    string vAmount = Double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                    Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_Edit");
                    string vTaxType = eTaxType.Text.Trim();
                    TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_Edit");
                    string vTaxRate = Double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                    Label eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_Edit");
                    string vTaxAMT = Double.TryParse(eTaxAMT.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                    Label eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_Edit");
                    string vTotalAmount = Double.TryParse(eTotalAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                    Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_Edit");
                    string vPayMode = ePayMode.Text.Trim();
                    TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_Edit");
                    string vRemarkA = eRemarkA.Text.Trim();
                    TextBox eSheetNote = (TextBox)fvConsSheetA_Detail.FindControl("eSheetNote_Edit");
                    string vSheetNote = eSheetNote.Text.Trim();
                    //更新用的 SQL 字串
                    string vUpdateStr = "update ConsSheetA " + Environment.NewLine +
                                        "   set SupNo = @SupNo, DepNo = @DepNo, AssignMan = @AssignMan, Amount = @Amount, " + Environment.NewLine +
                                        "       TaxType = @TaxType, TaxRate = @TaxRate, TaxAMT = @TaxAMT, TotalAmount = @TotalAmount, PayMode = @PayMode, " + Environment.NewLine +
                                        "       SheetNote = @SheetNote, RemarkA = @RemarkA, ModifyDate = GetDate(), ModifyMan = @ModifyMan " + Environment.NewLine +
                                        " where SheetNo = @SheetNo ";
                    sdsConsSheetA_Detail.UpdateCommand = vUpdateStr;
                    sdsConsSheetA_Detail.UpdateParameters.Clear();
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SupNo", DbType.String, (vSupNo != "") ? vSupNo : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (vDepNo != "") ? vDepNo : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("AssignMan", DbType.String, (vAssignMan != "") ? vAssignMan : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("Amount", DbType.Double, vAmount));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TaxType", DbType.String, (vTaxType != "") ? vTaxType : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TaxRate", DbType.Double, vTaxRate));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TaxAMT", DbType.Double, vTaxAMT));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("PayMode", DbType.String, (vPayMode != "") ? vPayMode : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNote", DbType.String, (vSheetNote != "") ? vSheetNote : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("RemarkA", DbType.String, (vRemarkA != "") ? vRemarkA : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetA_Detail.Update();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.DataBind();
                    fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
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
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
            if (eDepNo != null)
            {
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_Edit");
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo DESC";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
                }
                eDepNo.Text = vDepNo.Trim();
                eDepName.Text = vDepName.Trim();
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
                string vSQLStr = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
                string vAssignManName = PF.GetValue(vConnStr, vSQLStr, "Name");
                if (vAssignManName == "")
                {
                    vAssignManName = vAssignMan.Trim();
                    vSQLStr = "select top 1 EmpNo from Employee where [Name] like '%" + vAssignManName + "%' order by Assumeday DESC";
                    vAssignMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                }
                eAssignMan.Text = vAssignMan.Trim();
                eAssignManName.Text = vAssignManName.Trim();
            }
        }

        protected void eTaxType_C_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            double vTempFloat;
            DropDownList ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("eTaxType_C_Edit");
            if (ddlTaxType != null)
            {
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_Edit");
                eTaxType.Text = ddlTaxType.SelectedValue.Trim();
                Label eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmount_Edit");
                Double vAmount = Double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_Edit");
                Double vTaxRate = double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat : 0.05;
                Label eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_Edit");
                Double vTaxAMT = Double.TryParse(eTaxAMT.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_Edit");
                Double vTotalAmount = Double.TryParse(eTotalAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                switch (ddlTaxType.SelectedValue)
                {
                    case "1":
                        vTaxAMT = Math.Round(vAmount * vTaxRate, 0, MidpointRounding.AwayFromZero);
                        vTotalAmount = vAmount + vTaxAMT;
                        break;

                    case "2":
                        vTotalAmount = vAmount;
                        vAmount = Math.Round(vTotalAmount / (1 + vTaxRate), 0, MidpointRounding.AwayFromZero);
                        vTaxAMT = vTotalAmount - vAmount;
                        break;

                    case "3":
                    case "4":
                        vTaxRate = 0.0;
                        vTaxAMT = vAmount;
                        vTotalAmount = vAmount;
                        break;
                }
                eAmount.Text = vAmount.ToString();
                eTaxRate.Text = vTaxRate.ToString();
                eTaxAMT.Text = vTaxAMT.ToString();
                eTotalAmount.Text = vTotalAmount.ToString();
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

        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            double vTempFloat;
            DateTime vToday = DateTime.Today;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNo_INS");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vFirstCode = (vToday.Year - 1911).ToString("D3") + vToday.Month.ToString("D2") + "BS";
                string vSQLStr = "select max(SheetNo) MaxNo from ConsSheetA where SheetNo like '" + vFirstCode + "%' ";
                string vOldIndex = PF.GetValue(vConnStr, vSQLStr, "MaxNo");
                Int32 vIndex = 0;
                string vNewIndes = Int32.TryParse(vOldIndex.Replace(vFirstCode, ""), out vIndex) ? (vIndex + 1).ToString("D4") : "0001";
                string vSheetNo = vFirstCode + vNewIndes;
                TextBox eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_INS");
                string vSupNo = eSupNo.Text.Trim();
                TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
                string vDepNo = eDepNo.Text.Trim();
                TextBox eAssignMan = (TextBox)fvConsSheetA_Detail.FindControl("eAssignMan_INS");
                string vAssignMan = eAssignMan.Text.Trim();
                Label eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmount_INS");
                string vAmount = Double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_INS");
                string vTaxType = eTaxType.Text.Trim();
                TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_INS");
                string vTaxRate = Double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                Label eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_INS");
                string vTaxAMT = Double.TryParse(eTaxAMT.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                Label eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_INS");
                string vTotalAmount = Double.TryParse(eTotalAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_INS");
                string vPayMode = ePayMode.Text.Trim();
                TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_INS");
                string vRemarkA = eRemarkA.Text.Trim();
                TextBox eSheetNote = (TextBox)fvConsSheetA_Detail.FindControl("eSheetNote_INS");
                string vSheetNote = eSheetNote.Text.Trim();
                vSQLStr = "insert into ConsSheetA " + Environment.NewLine +
                          "       (SheetNo, SheetMode, SupNo, DepNo, AssignMan, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, PayMode, RemarkA, SheetNote) " + Environment.NewLine +
                          "values (@SheetNo, 'BS', @SupNo, @DepNo, @AssignMan, GetDate(), @BuMan, @Amount, @TaxType, @TaxRate, @TaxAMT, @TotalAmount, @PayMode, @RemarkA, @SheetNote) ";
                sdsConsSheetA_Detail.InsertCommand = vSQLStr;
                sdsConsSheetA_Detail.InsertParameters.Clear();
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SupNo", DbType.String, (vSupNo != "") ? vSupNo : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (vDepNo != "") ? vDepNo : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("AssignMan", DbType.String, (vAssignMan != "") ? vAssignMan : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("Amount", DbType.Double, vAmount));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TaxType", DbType.String, (vTaxType != "") ? vTaxType : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TaxRate", DbType.Double, vTaxRate));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TaxAMT", DbType.Double, vTaxAMT));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TotalAmount", DbType.Double, vTotalAmount));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("PayMode", DbType.String, (vPayMode != "") ? vPayMode : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("RemarkA", DbType.String, (vRemarkA != "") ? vRemarkA : String.Empty));
                sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNote", DbType.String, (vSheetNote != "") ? vSheetNote : String.Empty));
                sdsConsSheetA_Detail.Insert();
                gridConsSheetA_List.DataBind();
                fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                fvConsSheetA_Detail.DataBind();
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
            if (eDepNo != null)
            {
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_INS");
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo DESC";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
                }
                eDepNo.Text = vDepNo.Trim();
                eDepName.Text = vDepName.Trim();
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
                string vSQLStr = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
                string vAssignManName = PF.GetValue(vConnStr, vSQLStr, "Name");
                if (vAssignManName == "")
                {
                    vAssignManName = vAssignMan.Trim();
                    vSQLStr = "select top 1 EmpNo from Employee where [Name] like '%" + vAssignManName + "%' order by Assumeday DESC";
                    vAssignMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
                }
                eAssignMan.Text = vAssignMan.Trim();
                eAssignManName.Text = vAssignManName.Trim();
            }
        }

        protected void eTaxType_C_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            double vTempFloat;
            DropDownList ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("eTaxType_C_INS");
            if (ddlTaxType != null)
            {
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_INS");
                eTaxType.Text = ddlTaxType.SelectedValue.Trim();
                Label eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmount_INS");
                Double vAmount = Double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_INS");
                Double vTaxRate = double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat : 0.05;
                Label eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_INS");
                Double vTaxAMT = Double.TryParse(eTaxAMT.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_INS");
                Double vTotalAmount = Double.TryParse(eTotalAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                switch (ddlTaxType.SelectedValue)
                {
                    case "1":
                        vTaxAMT = Math.Round(vAmount * vTaxRate, 0, MidpointRounding.AwayFromZero);
                        vTotalAmount = vAmount + vTaxAMT;
                        break;

                    case "2":
                        vTotalAmount = vAmount;
                        vAmount = Math.Round(vTotalAmount / (1 + vTaxRate), 0, MidpointRounding.AwayFromZero);
                        vTaxAMT = vTotalAmount - vAmount;
                        break;

                    case "3":
                    case "4":
                        vTaxRate = 0.0;
                        vTaxAMT = vAmount;
                        vTotalAmount = vAmount;
                        break;
                }
                eAmount.Text = vAmount.ToString();
                eTaxRate.Text = vTaxRate.ToString();
                eTaxAMT.Text = vTaxAMT.ToString();
                eTotalAmount.Text = vTotalAmount.ToString();
            }
        }

        protected void ePayMode_C_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlPayMode = (DropDownList)fvConsSheetA_Detail.FindControl("ePayMode_C_INS");
            if (ddlPayMode != null)
            {
                Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_INS");
                ePayMode.Text = ddlPayMode.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 刪除請購單主檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            //主檔刪除
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            DateTime vToday = DateTime.Today;
            string vOptionString;
            string vSQLStr_Temp;
            string vRCount_Str;
            string vH_IndexFirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vMaxIndex;
            string vNewIndex;
            int vTempINT = 0;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNo_List");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    string vSheetNo = eSheetNo.Text.Trim();
                    //寫入異動記錄
                    vOptionString = "刪除總務耗材請購單" + Environment.NewLine +
                                    "單號：[ " + eSheetNo.Text.Trim() + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：刪除" + Environment.NewLine +
                                    "ConsSheet_Order.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //檢查有沒有明細，有明細要一併刪除
                    vSQLStr_Temp = "select Count(SheetNoItems) RCount from ConsSheetB where SheetNo = '" + eSheetNo.Text.Trim() + "' ";
                    vRCount_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount");
                    if ((Int32.TryParse(vRCount_Str, out vTempINT)) && (vTempINT > 0))
                    {
                        vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = @SheetNo";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            cmdTemp.Parameters.Clear();
                            cmdTemp.Parameters.Add(new SqlParameter("SheetNo", eSheetNo.Text.Trim()));
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_IndexFirstCode + "%' ";
                                vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                                vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_IndexFirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                                vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                               "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                               "       QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionDate, ActionMan) " + Environment.NewLine +
                                               "select '" + vH_IndexFirstCode + vNewIndex + "'. 'DELA', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                               "       QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, GetDate(), '" + vLoginID + "' " + Environment.NewLine +
                                               "  from ConsSheetB " + Environment.NewLine +
                                               " where SheetNoItems = '" + drTemp["SheetNoItems"].ToString().Trim() + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            }
                        }
                        vSQLStr_Temp = "delete ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    }
                    //備份主檔資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_IndexFirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_IndexFirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                                   "       SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_IndexFirstCode + vNewIndex + "', 'DELA', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                                   "       SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //刪除資料
                    vSQLStr_Temp = "delete ConSheetA where SheetNo = @SheetNo";
                    sdsConsSheetA_Detail.DeleteCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.DeleteParameters.Clear();
                    sdsConsSheetA_Detail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
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

        /// <summary>
        /// 列印請購單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            // 列印請購單
            DataTable dtPrintA;
            DataTable dtPrintB;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNo_List");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSheetNo = eSheetNo.Text.Trim();
                string vSelectStr = "select a.SheetNo, a.DepNo, d.[Name] DepName, a.AssignMan, e.[Name] AssignManName, " + Environment.NewLine +
                                    "       convert(varchar(10), a.BuDate, 111) BuDate " + Environment.NewLine +
                                    "  from ConsSheetA a left join Department d on d.DepNo = a.DepNo " + Environment.NewLine +
                                    "                    left join Employee e on e.EmpNo = a.AssignMan " + Environment.NewLine +
                                    " where SheetNo = '" + vSheetNo + "' ";
                using (SqlConnection connSheetA=new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daSheetA = new SqlDataAdapter(vSelectStr, connSheetA);
                    connSheetA.Open();
                    dtPrintA = new DataTable();
                    daSheetA.Fill(dtPrintA);
                }
                
                string vCompanyName = PF.GetValue(vConnStr, "select [Name] from [Custom] where [Code] = 'A000' and Types = 'O' ", "Name");
                string vReportName = "請購 (修) 單";
                string vReportPath = @"Report\ConsSheet_OrderP.rdlc";
                ReportDataSource rdsPrintA = new ReportDataSource("ConsSheetA_OrderP", dtPrintA);
                //ReportDataSource rdsPrintB = new ReportDataSource("ConsSheet_OrderPB", dtPrintB);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = vReportPath;
                rvPrint.LocalReport.DataSources.Add(rdsPrintA);
                //rvPrint.LocalReport.DataSources.Add(rdsPrintB);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.Refresh();
                plReport.Visible = true;
                plSearch.Visible = false;
                plShowData_A.Visible = false;
                plShowDetail.Visible = false;
            }
        }

        protected void Unnamed_SubreportProcessing(object sender, SubreportProcessingEventArgs e)
        {
            if (vConnStr=="")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNo_List");
            string vSheetNo = eSheetNo.Text.Trim();
            DataTable dtPrintB;
            string vSelectStr = "select b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.ConsUnit, d.ClassTxt ConsUnit_C, " + Environment.NewLine +
                                "       b.Quantity, b.RemarkB " + Environment.NewLine +
                                "  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo " + Environment.NewLine +
                                "                    left join DBDICB d on d.ClassNo = b.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit'" + Environment.NewLine +
                                " where b.SheetNo = '" + vSheetNo + "' order by Items ";
            using (SqlConnection connSheetB = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintB = new SqlDataAdapter(vSelectStr, connSheetB);
                connSheetB.Open();
                dtPrintB = new DataTable();
                daPrintB.Fill(dtPrintB);
            }
            ReportDataSource rdsPrintB = new ReportDataSource("ConsSheetB_OrderP", dtPrintB);
            e.DataSources.Add(rdsPrintB);
        }

        protected void gridConsSheetA_List_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridConsSheetA_List.PageIndex = e.NewPageIndex;
            gridConsSheetA_List.DataBind();
        }

        protected void gridConsSheetB_List_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridConsSheetB_List.PageIndex = e.NewPageIndex;
            gridConsSheetB_List.DataBind();
        }

        protected void fvConsSheetB_Detail_DataBound(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            Label eSheetNoA = (Label)fvConsSheetA_Detail.FindControl("eSheetNo_List");
            Label eItemStatus;
            Label eSheetNo;
            TextBox eConsNo;
            TextBox eConsName;
            string vConsDataURL;
            string vConsDataScript;
            switch (fvConsSheetB_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    eItemStatus = (Label)fvConsSheetB_Detail.FindControl("eItemStatus_List");
                    if (eItemStatus != null)
                    {
                        Button bbEdit = (Button)fvConsSheetB_Detail.FindControl("bbEdit_B_List");
                        Button bbAbort = (Button)fvConsSheetB_Detail.FindControl("bbAbort_B_List");
                        Button bbDelete = (Button)fvConsSheetB_Detail.FindControl("bbDelete_B_List");
                        if (eItemStatus.Text.Trim() == "000")
                        {
                            bbEdit.Visible = true;
                            bbAbort.Visible = true;
                            bbDelete.Visible = true;
                        }
                        else
                        {
                            bbEdit.Visible = false;
                            bbAbort.Visible = false;
                            bbDelete.Visible = false;
                        }
                    }
                    break;
                case FormViewMode.Edit:
                    Label eSheetNoItems_Edit = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_Edit");
                    if (eSheetNoItems_Edit != null)
                    {
                        eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
                        eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_Edit");
                        vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsPriceID=" + eConsName.ClientID;
                        vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eConsNo.Attributes["onClick"] = vConsDataScript;
                    }
                    break;
                case FormViewMode.Insert:
                    Label eSheetNoItems_INS = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_INS");
                    if (eSheetNoItems_INS != null)
                    {
                        eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNo_B_INS");
                        eSheetNo.Text = eSheetNoA.Text.Trim();
                        eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                        eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_INS");
                        vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsPriceID=" + eConsName.ClientID;
                        vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eConsNo.Attributes["onClick"] = vConsDataScript;
                        Label eQtyMode = (Label)fvConsSheetB_Detail.FindControl("eQtyMode_INS");
                        eQtyMode.Text = "-1";
                        eItemStatus = (Label)fvConsSheetB_Detail.FindControl("eItemStatus_INS");
                        eItemStatus.Text = "000";
                        Label eItemStatus_C = (Label)fvConsSheetB_Detail.FindControl("eItemStatus_C_INS");
                        eItemStatus_C.Text = "已開單";
                    }
                    break;
            }
        }

        protected void bbOK_B_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            DateTime vToday = DateTime.Today;
            string vSQLStr_Temp;
            string vOptionString;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vMaxIndex;
            string vNewIndex;
            double vTempFloat = 0.0;
            int vTempINT;
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_Edit");
            if ((eSheetNoItems != null) && (eSheetNoItems.Text.Trim() != ""))
            {
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_Edit");
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_Edit");
                try
                {
                    //寫入異動記錄
                    vOptionString = "編輯總務耗材請購單明細" + Environment.NewLine +
                                "單號：[ " + eSheetNo.Text.Trim() + " ]" + Environment.NewLine +
                                "項次：[ " + eItems.Text.Trim() + " ] " + Environment.NewLine +
                                "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                "異動種類：修改" + Environment.NewLine +
                                "ConsSheet_Order.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份明細資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionMan, ActionDate) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //修改明細資料
                    TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
                    string vConsNo = eConsNo.Text.Trim();
                    TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_Edit");
                    string vPrice = double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                    TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_Edit");
                    string vQuantity = double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                    Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmount_B_Edit");
                    string vAmount = double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                    TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_Edit");
                    string vRemarkB = eRemarkB.Text.Trim();
                    //組合更新語法
                    string vUpdateStr = "update ConsSheetB set ConsNo = @ConsNo, Price = @Price, Quantity = @Quantity, " + Environment.NewLine +
                                        "                      Amount = @Amount, RemarkB = @RemarkB, ModifyMan = @ModifyMan, ModifyDate = getDate() " + Environment.NewLine +
                                        " where SheetNoItems = @SheetNoItems ";
                    //代入參數
                    sdsConsSheetB_Detail.UpdateCommand = vUpdateStr;
                    sdsConsSheetB_Detail.UpdateParameters.Clear();
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ConsNo", DbType.String, (vConsNo != "") ? vConsNo : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Price", DbType.Double, vPrice));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Quantity", DbType.Double, vQuantity));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Amount", DbType.Double, vAmount));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("RemarkB", DbType.String, (vRemarkB != "") ? vRemarkB : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    //更新資料
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
                Label ePrice = (Label)fvConsSheetB_Detail.FindControl("ePrice_Edit");
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
            if ((eQuantity != null) && (eQuantity.Text.Trim() != ""))
            {
                double vTempFloat;
                Label ePrice = (Label)fvConsSheetB_Detail.FindControl("ePrice_Edit");
                Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmount_B_Edit");
                double vPrice = double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                double vQuantity = double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                eAmount.Text = (vPrice * vQuantity).ToString();
            }
        }

        protected void bbOK_B_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            double vTempFloat = 0.0;
            Int32 vTempINT;
            Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNo_B_INS");
            if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSheetNo = eSheetNo.Text.Trim();
                string vSQLStr = "select max(Items) MaxItem from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                string vItems = Int32.TryParse(PF.GetValue(vConnStr, vSQLStr, "MaxItem"), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                string vSheetNoItems = vSheetNo + vItems;
                TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                string vConsNo = eConsNo.Text.Trim();
                Label ePrice = (Label)fvConsSheetB_Detail.FindControl("ePrice_INS");
                string vPrice = double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_INS");
                string vQuantity = double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmount_B_INS");
                string vAmount = double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_INS");
                string vRemarkB = eRemarkB.Text.Trim();
                Label eQtyMode = (Label)fvConsSheetB_Detail.FindControl("eQtyMode_INS");
                string vQtyMode = Int32.TryParse(eQtyMode.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "-1";
                Label eItemStatus = (Label)fvConsSheetB_Detail.FindControl("eItemStatus_INS");
                string vItemStatus = eItemStatus.Text.Trim();
                Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                string vConsUnit = eConsUnit.Text.Trim();
                //組織新增語法
                vSQLStr = "Insert into ConsSheetB " + Environment.NewLine +
                         "       (SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                         "        QtyMode, ItemStatus, BuMan, BuDate, ConsUnit)" + Environment.NewLine +
                         "values (@SheetNoItems, @SheetNo, @Items, @ConsNo, @Price, @Quantity, @Amount, @RemarkB, " + Environment.NewLine +
                         "        @QtyMode, @ItemStatus, @BuMan, GetDate(), @ConsUnit)";
                sdsConsSheetB_Detail.InsertCommand = vSQLStr;
                sdsConsSheetB_Detail.InsertParameters.Clear();
                //填入參數
                try
                {
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ConsNo", DbType.String, (vConsNo != "") ? vConsNo : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Price", DbType.Double, vPrice));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Quantity", DbType.Double, vQuantity));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Amount", DbType.Double, vAmount));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("RemarkB", DbType.String, (vRemarkB != "") ? vRemarkB : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("QtyMode", DbType.Int32, vQtyMode));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ItemStatus", DbType.String, (vItemStatus != "") ? vItemStatus : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ConsUnit", DbType.String, (vConsUnit != "") ? vConsUnit : String.Empty));
                    sdsConsSheetB_Detail.Insert();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                    throw;
                }
            }
        }

        protected void eConsName_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
            if ((eConsNo != null) && (eConsNo.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vConsNo = eConsNo.Text.Trim();
                double vTempFloat;
                Label ePrice = (Label)fvConsSheetB_Detail.FindControl("ePrice_INS");
                Label eStockQty = (Label)fvConsSheetB_Detail.FindControl("eStockQty_INS");
                Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                Label eConsUnit_C = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_C_INS");
                string vSQLStr_Temp = "select c.StockQty, c.LastInPrice, c.ConsUnit, d.ClassTxt ConsUnit_C " + Environment.NewLine +
                                      "  from Consumables c left join DBDICB d on d.ClassNo = c.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                                      " where c.ConsNo = @ConsNo";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                    cmdTemp.Parameters.Clear();
                    cmdTemp.Parameters.Add(new SqlParameter("ConsNo", vConsNo));
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        ePrice.Text = double.TryParse(drTemp["LastInPrice"].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                        eStockQty.Text = double.TryParse(drTemp["StockQty"].ToString().Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0";
                        eConsUnit.Text = drTemp["ConsUnit"].ToString().Trim();
                        eConsUnit_C.Text = drTemp["ConsUnit_C"].ToString().Trim();
                    }
                }
            }
        }

        protected void eQuantity_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_INS");
            if ((eQuantity != null) && (eQuantity.Text.Trim() != ""))
            {
                double vTempFloat;
                Label ePrice = (Label)fvConsSheetB_Detail.FindControl("ePrice_INS");
                Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmount_B_INS");
                double vPrice = double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                double vQuantity = double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                eAmount.Text = (vPrice * vQuantity).ToString();
            }
        }

        /// <summary>
        /// 明細作廢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbAbort_B_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_List");
            DateTime vToday = DateTime.Today;
            string vOptionString;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vSQLStr_Temp;
            string vMaxIndex;
            string vNewIndex;
            int vTempINT;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (eSheetNoItems != null)
            {
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_List");
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_List");
                try
                {
                    //寫入異動記錄
                    vOptionString = "異動總務課耗材請購單明細" + Environment.NewLine +
                                    "單號： [ " + eSheetNo.Text.Trim() + " ] " + Environment.NewLine +
                                    "項次： [ " + eItems.Text.Trim() + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：明細作廢" + Environment.NewLine +
                                    "ConsSheet_Order.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份明細資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                   "       QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionMan, ActionDate) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                   "       QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + eSheetNoItems.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入作廢資料
                    vSQLStr_Temp = "update ConsSheetB set ItemStatus = '999', ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "' " + Environment.NewLine +
                                   " where SheetNoItems = '" + eSheetNoItems.Text.Trim() + "' and ItemStatus = '000' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        /// <summary>
        /// 刪除明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_B_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            DateTime vToday = DateTime.Today;
            string vSQLStr_Temp;
            string vOptionString;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vMaxIndex;
            string vNewIndex;
            int vTempINT;
            if (vConnStr == "")

            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems != null)
            {
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_List");
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_List");
                try
                {
                    //寫入異動記錄
                    vOptionString = "刪除總務課耗材請購單明細" + Environment.NewLine +
                                    "單號： [ " + eSheetNo.Text.Trim() + " ] " + Environment.NewLine +
                                    "項次： [ " + eItems.Text.Trim() + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：刪除" + Environment.NewLine +
                                    "ConsSheet_Order.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份明細資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                   "       QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyMan, ModifyDate, ConsUnit, ActionMan, ActionDate) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'DELB', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                   "       QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyMan, ModifyDate, ConsUnit, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + eSheetNoItems.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    vSQLStr_Temp = "delete ConsSheetB where SheetNoItems = @SheetNoItems ";
                    sdsConsSheetB_Detail.DeleteCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.DeleteParameters.Clear();
                    sdsConsSheetB_Detail.DeleteParameters.Add(new Parameter("SheetNoItems", DbType.String, eSheetNoItems.Text.Trim()));
                    sdsConsSheetB_Detail.Delete();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plSearch.Visible = true;
            plShowData_A.Visible = true;
            plShowDetail.Visible = true;
            rvPrint.LocalReport.DataSources.Clear();
            plReport.Visible = false;
        }
    }
}