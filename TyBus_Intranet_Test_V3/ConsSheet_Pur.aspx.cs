using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ConsSheet_Pur : Page
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
                    //動態掛載日期輸入視窗
                    string vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDateS_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBuDateS_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDateE_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBuDateE_Search.Attributes["onClick"] = vDateScript;

                    if (!IsPostBack)
                    {
                        plSearch.Visible = true;
                        plShowDataA.Visible = true;
                        plShowDataB.Visible = true;
                        plPickupDetail.Visible = false;
                        plPrint.Visible = false;
                    }
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    OpenData();
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

        private void OpenData()
        {
            string vSelectStr = GetSelectStr();
            sdsConsSheetA_List.SelectCommand = vSelectStr;
            sdsConsSheetA_List.Select(new DataSourceSelectArguments());
            gridConsSheetA_List.DataBind();
        }

        private string GetSelectStr()
        {
            DateTime vTempDateS;
            DateTime vTempDateE;
            string vResultStr;
            string vWStr_BuDate = (DateTime.TryParse(eBuDateS_Search.Text.Trim(), out vTempDateS) && DateTime.TryParse(eBuDateE_Search.Text.Trim(), out vTempDateE)) ?
                                  "   and a.BuDate between '" + PF.GetAD(vTempDateS.ToShortDateString()) + "' and '" + PF.GetAD(vTempDateE.ToShortDateString()) + "' " + Environment.NewLine :
                                  (DateTime.TryParse(eBuDateS_Search.Text.Trim(), out vTempDateS) && !DateTime.TryParse(eBuDateE_Search.Text.Trim(), out vTempDateE)) ?
                                  "   and a.BuDate = '" + PF.GetAD(vTempDateS.ToShortDateString()) + "' " + Environment.NewLine :
                                  (!DateTime.TryParse(eBuDateS_Search.Text.Trim(), out vTempDateS) && DateTime.TryParse(eBuDateE_Search.Text.Trim(), out vTempDateE)) ?
                                  "   and a.BuDate = '" + PF.GetAD(vTempDateE.ToShortDateString()) + "' " + Environment.NewLine : "";
            string vWStr_SupName = (eSupNo_Search.Text.Trim() != "") ? "   and c.[Name] like '%" + eSupNo_Search + "%' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ?
                                    "   and a.SheetNo in (select SheetNo from ConsSheetB where ConsNo in (select ConsNo from Consumables where ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine :
                                    "";
            vResultStr = "select a.SheetNo, a.BuDate, e.[Name] BuMan_C, c.[Name] SupName, a.TotalAmount, d.ClassTxt SheetStatus_C " + Environment.NewLine +
                         "  from ConsSheetA a left join [Custom] c on c.Code = a.SupNo and c.[Types] = 'S' " + Environment.NewLine +
                         "                    left join DBDICB d on d.ClassNo = a.SheetStatus and d.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus' " + Environment.NewLine +
                         "                    left join Employee e on e.EmpNo = a.BuMan " + Environment.NewLine +
                         " where a.SheetMode = 'PS' " + Environment.NewLine +
                         vWStr_BuDate +
                         vWStr_ConsName +
                         vWStr_SupName +
                         " order by a.SheetNo ";
            return vResultStr;
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
            string vSQLStr;
            Label eSheetNo;
            TextBox eSupNo;
            TextBox eSupName;
            string vSupDataURL;
            string vSupDataScript;
            DropDownList ddlTaxType;
            Label eTaxType;
            DropDownList ddlPayMode;
            Label ePayMode;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvConsSheetA_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
                    if (eSheetNo != null)
                    {
                        string vSheetNo = eSheetNo.Text.Trim();
                        //有 "已開單" 的明細才開放【整單到貨】的按鈕
                        vSQLStr = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '000') = '000' ";
                        string vRCount_Str = PF.GetValue(vConnStr, vSQLStr, "RCount");
                        int vRCount = Int32.Parse(vRCount_Str);
                        Button bbArriveA = (Button)fvConsSheetA_Detail.FindControl("bbArriveA_List");
                        bbArriveA.Enabled = (vRCount > 0);
                        //只要有不是 "已開單" 或 "已作廢" 的明細就不開放整單作廢
                        vSQLStr = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '000') not in ('000', '999') ";
                        vRCount_Str = PF.GetValue(vConnStr, vSQLStr, "RCount");
                        vRCount = Int32.Parse(vRCount_Str);
                        Button bbAbortA = (Button)fvConsSheetA_Detail.FindControl("bbAbortA_List");
                        bbAbortA.Enabled = (vRCount <= 0);
                        //有 "已到貨" 的明細才開放【付款】按鈕
                        vSQLStr = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '000') = '001' ";
                        vRCount_Str = PF.GetValue(vConnStr, vSQLStr, "RCount");
                        vRCount = Int32.Parse(vRCount_Str);
                        Button bbCloseA = (Button)fvConsSheetA_Detail.FindControl("bbCloseA_List");
                        bbCloseA.Enabled = (vRCount > 0);
                    }
                    break;
                case FormViewMode.Edit:
                    eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_Edit");
                    if (eSheetNo != null)
                    {
                        eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_Edit");
                        eSupName = (TextBox)fvConsSheetA_Detail.FindControl("eSupName_Edit");
                        vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo.ClientID + "&SupNameID=" + eSupName.ClientID + "&SupKind=C";
                        vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eSupNo.Attributes["onClick"] = vSupDataScript;
                        ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("ddlTaxType_Edit");
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
                        ddlPayMode = (DropDownList)fvConsSheetA_Detail.FindControl("ddlPayMode_Edit");
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
                    }
                    break;
                case FormViewMode.Insert:
                    eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_INS");
                    if (eSheetNo != null)
                    {
                        eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_INS");
                        eSupName = (TextBox)fvConsSheetA_Detail.FindControl("eSupName_INS");
                        vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo.ClientID + "&SupNameID=" + eSupName.ClientID + "&SupKind=C";
                        vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eSupNo.Attributes["onClick"] = vSupDataScript;
                        ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("ddlTaxType_INS");
                        ddlTaxType.Items.Clear();
                        eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_INS");
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
                        ddlTaxType.SelectedIndex = 0;
                        ddlPayMode = (DropDownList)fvConsSheetA_Detail.FindControl("ddlPayMode_INS");
                        ddlPayMode.Items.Clear();
                        vSQLStr = "select cast('' as varchar) ClassNo, Cast('' as varchar) ClassTxt " + Environment.NewLine +
                                  " union all " + Environment.NewLine +
                                  "select ClassNo, ClassTxt from DBDICB where FKey = '總務請購單      ConsSheetA      PayMode' " + Environment.NewLine +
                                  " order by ClassNo ";
                        ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_INS");
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
                        ddlPayMode.SelectedIndex = 0;
                    }
                    break;
            }
        }

        /// <summary>
        /// 主檔確認修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKA_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_Edit");
            if (eSheetNo != null)
            {
                int vTempINT;
                double vTempFloat;
                string vSheetNo = eSheetNo.Text.Trim();
                string vSQLStr_Temp;
                string vOptionString;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vMaxIndex;
                string vNewIndex;
                try
                {
                    //寫入異動記錄
                    vOptionString = "修改總務耗材採購單" + Environment.NewLine +
                                    "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：修改" + Environment.NewLine +
                                    "ConsSheet_Pur.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份舊資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                                   "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                                   "       ModifyMan, ModifyDate, DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                                   "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                                   "       ModifyMan, ModifyDate, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入異動
                    Label eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmountA_Edit");
                    Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_Edit");
                    Label eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_Edit");
                    Label eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_Edit");
                    Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_Edit");
                    TextBox eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_Edit");
                    TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_Edit");
                    TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_Edit");
                    TextBox eSheetNote = (TextBox)fvConsSheetA_Detail.FindControl("eSheetNote_Edit");
                    TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
                    vSQLStr_Temp = "update ConsSheetA set SupNo = @SupNo, Amount = @Amount, TaxType = @TaxType, TaxRate = @TaxRate, " + Environment.NewLine +
                                   "                      TaxAMT = @TaxAMT, TotalAmount = @TotalAmount, PayMode = @PayMode, DepNo = @DepNo, " + Environment.NewLine +
                                   "                      ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                   " where SheetNo = @SheetNo ";
                    sdsConsSheetA_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.UpdateParameters.Clear();
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SupNo", DbType.String, (eSupNo.Text.Trim() != "") ? eSupNo.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("Amount", DbType.Double, double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TaxType", DbType.String, (eTaxType.Text.Trim() != "") ? eTaxType.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TaxRate", DbType.Double, double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TaxAMT", DbType.Double, double.TryParse(eTaxAMT.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Double, double.TryParse(eTotalAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("PayMode", DbType.String, (ePayMode.Text.Trim() != "") ? ePayMode.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, eSheetNo.Text.Trim()));
                    sdsConsSheetA_Detail.Update();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 計算稅額
        /// </summary>
        /// <param name="fMode"></param>
        /// <param name="fTaxRate"></param>
        /// <param name="fTaxType"></param>
        private void CalTax(FormViewMode fMode, double fTaxRate, string fTaxType)
        {
            double vTempFloat;
            Label eAmount;
            double vAmount;
            Label eTaxAMT;
            double vTaxAMT;
            Label eTotalAmount;
            double vTotalAmount;
            switch (fMode)
            {
                case FormViewMode.Edit:
                    eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmountA_Edit");
                    if (eAmount != null)
                    {
                        eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_Edit");
                        eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_Edit");
                        vAmount = double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                        if (fTaxType == "2")
                        {
                            vTaxAMT = Math.Round(vAmount / (1 + fTaxRate), 2, MidpointRounding.AwayFromZero);
                            vTotalAmount = vAmount;
                            vAmount = vTotalAmount - vTaxAMT;
                        }
                        else
                        {
                            vTaxAMT = vAmount * fTaxRate;
                            vTotalAmount = vAmount + vTaxAMT;
                        }
                        eAmount.Text = vAmount.ToString();
                        eTaxAMT.Text = vTaxAMT.ToString();
                        eTotalAmount.Text = vTotalAmount.ToString();
                    }
                    break;
                case FormViewMode.Insert:
                    eAmount = (Label)fvConsSheetA_Detail.FindControl("eAmountA_INS");
                    if (eAmount != null)
                    {
                        eTaxAMT = (Label)fvConsSheetA_Detail.FindControl("eTaxAMT_INS");
                        eTotalAmount = (Label)fvConsSheetA_Detail.FindControl("eTotalAmount_INS");
                        vAmount = double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                        if (fTaxType == "2")
                        {
                            vTaxAMT = Math.Round(vAmount / (1 + fTaxRate), 2, MidpointRounding.AwayFromZero);
                            vTotalAmount = vAmount;
                            vAmount = vTotalAmount - vTaxAMT;
                        }
                        else
                        {
                            vTaxAMT = vAmount * fTaxRate;
                            vTotalAmount = vAmount + vTaxAMT;
                        }
                        eAmount.Text = vAmount.ToString();
                        eTaxAMT.Text = vTaxAMT.ToString();
                        eTotalAmount.Text = vTotalAmount.ToString();
                    }
                    break;
            }
        }

        protected void eTaxRate_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_Edit");
            if (eTaxRate != null)
            {
                double vTempFloat;
                double vTaxRate = double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_Edit");
                string vTaxType = (eTaxType.Text.Trim() != "") ? eTaxType.Text.Trim() : "1";
                CalTax(fvConsSheetA_Detail.CurrentMode, vTaxRate, vTaxType);
            }
        }

        protected void ddlTaxType_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("ddlTaxType_Edit");
            if (ddlTaxType != null)
            {
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_Edit");
                eTaxType.Text = ddlTaxType.SelectedValue.Trim();
                TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_Edit");
                switch (eTaxType.Text.Trim())
                {
                    case "1":
                    case "2":
                        eTaxRate.Text = "0.05";
                        break;

                    default:
                        eTaxRate.Text = "0.0";
                        break;
                }
                CalTax(fvConsSheetA_Detail.CurrentMode, double.Parse(eTaxRate.Text), eTaxType.Text.Trim());
            }
        }

        protected void ddlPayMode_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlPayMode = (DropDownList)fvConsSheetA_Detail.FindControl("ddlPayMode_Edit");
            if (ddlPayMode != null)
            {
                Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_Edit");
                ePayMode.Text = ddlPayMode.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 主檔新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKA_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_INS");
            if (eSheetNo != null)
            {
                double vTempFloat;
                string vSheetNo = PF.GetConsSheetNo(vConnStr, "PS");
                string vSQLStr_Temp;
                try
                {
                    Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_INS");
                    Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_INS");
                    TextBox eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_INS");
                    TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_INS");
                    TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_INS");
                    TextBox eSheetNote = (TextBox)fvConsSheetA_Detail.FindControl("eSheetNote_INS");
                    TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
                    vSQLStr_Temp = "insert into ConsSheetA " + Environment.NewLine +
                                   "      (SheetNo, SheetMode, SupNo, BuMan, BuDate, TaxType, TaxRate, " + Environment.NewLine +
                                   "       SheetStatus, StatusDate, PayMode, RemarkA, SheetNote, DepNo) " + Environment.NewLine +
                                   "values(@SheetNo, 'PS', @SupNo, @BuMan, GetDate(), @TaxType, @TaxRate, " + Environment.NewLine +
                                   "       '000', GetDate(), @PayMode, @RemarkA, @SheetNote, @DepNo) ";
                    sdsConsSheetA_Detail.InsertCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.InsertParameters.Clear();
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SupNo", DbType.String, (eSupNo.Text.Trim() != "") ? eSupNo.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TaxType", DbType.String, (eTaxType.Text.Trim() != "") ? eTaxType.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("TaxRate", DbType.Double, double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("PayMode", DbType.String, (ePayMode.Text.Trim() != "") ? ePayMode.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA.Text.Trim() != "") ? eRemarkA.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNote", DbType.String, (eSheetNote.Text.Trim() != "") ? eSheetNote.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.Insert();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        protected void ddlTaxType_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTaxType = (DropDownList)fvConsSheetA_Detail.FindControl("ddlTaxType_INS");
            if (ddlTaxType != null)
            {
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_INS");
                eTaxType.Text = ddlTaxType.SelectedValue.Trim();
                TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_INS");
                switch (eTaxType.Text.Trim())
                {
                    case "1":
                    case "2":
                        eTaxRate.Text = "0.05";
                        break;

                    default:
                        eTaxRate.Text = "0.0";
                        break;
                }
                CalTax(fvConsSheetA_Detail.CurrentMode, double.Parse(eTaxRate.Text), eTaxType.Text.Trim());
            }
        }

        protected void eTaxRate_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eTaxRate = (TextBox)fvConsSheetA_Detail.FindControl("eTaxRate_INS");
            if (eTaxRate != null)
            {
                double vTempFloat;
                double vTaxRate = double.TryParse(eTaxRate.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                Label eTaxType = (Label)fvConsSheetA_Detail.FindControl("eTaxType_INS");
                string vTaxType = (eTaxType.Text.Trim() != "") ? eTaxType.Text.Trim() : "1";
                CalTax(fvConsSheetA_Detail.CurrentMode, vTaxRate, vTaxType);
            }
        }

        protected void ddlPayMode_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlPayMode = (DropDownList)fvConsSheetA_Detail.FindControl("ddlPayMode_INS");
            if (ddlPayMode != null)
            {
                Label ePayMode = (Label)fvConsSheetA_Detail.FindControl("ePayMode_INS");
                ePayMode.Text = ddlPayMode.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 整單到貨
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbArriveA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            int vTempINT;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vOptionString;
            string vSQLStr_Temp;
            string vMaxIndex;
            string vNewIndex;
            string vSheetNo;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                vSheetNo = eSheetNo.Text.Trim();
                //寫入異動記錄
                vOptionString = "修改總務耗材採購單" + Environment.NewLine +
                                "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                "異動種類：整單到貨" + Environment.NewLine +
                                "ConsSheet_Pur.aspx";
                try
                {
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //找出狀態是 "已開單" 的明細
                    vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = '" + vSheetNo + "' and ItemStatus = '000' ";
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        if (drTemp.HasRows)
                        {
                            string vSheetNoItems;
                            while (drTemp.Read())
                            {
                                vSheetNoItems = drTemp["SheetNoItems"].ToString().Trim();
                                vSQLStr_Temp = "update ConsSheetB set ItemStatus = '001', ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "' " + Environment.NewLine +
                                               " where SheetNoItems = '" + vSheetNoItems + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            }
                        }
                    }
                    //備份舊資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'ARI', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //更新資料
                    vSQLStr_Temp = "update ConsSheetA set SheetStatus = '010', StatusDate = GetDate(), ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "' " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = false;
                }
            }
        }

        /// <summary>
        /// 付款結案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            int vTempINT;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vOptionString;
            string vSQLStr_Temp;
            string vMaxIndex;
            string vNewIndex;
            string vSheetNo;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                vSheetNo = eSheetNo.Text.Trim();
                //寫入異動記錄
                vOptionString = "修改總務耗材採購單" + Environment.NewLine +
                                "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                "異動種類：整單付款" + Environment.NewLine +
                                "ConsSheet_Pur.aspx";
                try
                {
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //找出狀態是 "已到貨" 的明細並設定為結案
                    vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = '" + vSheetNo + "' and ItemStatus = '001' ";
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        if (drTemp.HasRows)
                        {
                            string vSheetNoItems;
                            while (drTemp.Read())
                            {
                                vSheetNoItems = drTemp["SheetNoItems"].ToString().Trim();
                                vSQLStr_Temp = "update ConsSheetB set ItemStatus = '998', ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "' " + Environment.NewLine +
                                               " where SheetNoItems = '" + vSheetNoItems + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            }
                        }
                    }
                    //備份舊資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'PAID', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //更新資料
                    vSQLStr_Temp = "update ConsSheetA set SheetStatus = '001', StatusDate = GetDate(), ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "' " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = false;
                }
            }
        }

        /// <summary>
        /// 列印採購單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPtint_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                //產生參數需要的值
                Label eBuDate = (Label)fvConsSheetA_Detail.FindControl("eBuDateA_List");
                Label eSupName = (Label)fvConsSheetA_Detail.FindControl("eSupName_List");
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_List");
                DateTime vTempDate;
                DataTable dtPrint;
                string vSheetNo = eSheetNo.Text.Trim();
                string vReportName = "訂購單"; //報表名稱
                string vCompanyName = PF.GetValue(vConnStr, "select [Name] from [Custom] where [Code] = 'A000' and Types = 'O' ", "Name"); //公司名稱
                string vBuildDate = DateTime.TryParse(eBuDate.Text.Trim(), out vTempDate) ?
                                    (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd") :
                                    (vToday.Year - 1911).ToString("D4") + "/" + vToday.ToString("MM/dd"); //用民國年顯示訂購日期
                string vSupplierName = eSupName.Text.Trim(); //廠商名稱
                string vOrderName = eDepName.Text.Trim(); ;
                //產生報表資料
                string vSelectStr = "select b.SheetNoItems, b.ConsNo, c.ConsName, d.ClassTxt ConsUnit_C, b.Quantity, b.Price, b.Amount, b.RemarkB " + Environment.NewLine +
                                    "  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo " + Environment.NewLine +
                                    "                    left join DBDICB d on d.ClassNo = b.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                                    " where SheetNo = '" + vSheetNo + "' " + Environment.NewLine +
                                    " order by Items ";
                using (SqlConnection connPrint = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                    connPrint.Open();
                    dtPrint = new DataTable();
                    daPrint.Fill(dtPrint);
                }
                string vReportPath = @"Report\ConsSheet_PurP.rdlc"; //報表路徑

                ReportDataSource rdsPrint = new ReportDataSource("ConsSheet_PurP", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = vReportPath;
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("BuildDate", vBuildDate));
                rvPrint.LocalReport.SetParameters(new ReportParameter("SupplierName", vSupplierName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("OrderName", vOrderName));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plSearch.Visible = false;
                plShowDataA.Visible = false;
                plShowDataB.Visible = false;
                plPickupDetail.Visible = false;
            }
        }

        /// <summary>
        /// 整單作廢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbAbortA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            int vTempINT;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vOptionString;
            string vSQLStr_Temp;
            string vMaxIndex;
            string vNewIndex;
            string vSheetNo;
            string vRCount_S;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                vSheetNo = eSheetNo.Text.Trim();
                //先找出明細有沒有狀態不是【已開單】或【已作廢】的，只要有一筆就不能作廢
                vSQLStr_Temp = "select count(Items) RCount from ConsSheetB " + Environment.NewLine +
                               " where SheetNo = '" + vSheetNo + "' " + Environment.NewLine +
                               "   and ItemStatus not in ('000', '999') ";
                try
                {
                    vRCount_S = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount");
                    vTempINT = Int32.Parse(vRCount_S);
                    if (vTempINT == 0)
                    {
                        //先把明細作廢
                        vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = '" + vSheetNo + "' and ItemStatus = '000' ";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            if (drTemp.HasRows)
                            {
                                string vSheetNoItems;
                                while (drTemp.Read())
                                {
                                    vSheetNoItems = drTemp["SheetNoItems"].ToString().Trim();
                                    vSQLStr_Temp = "update ConsSheetB set ItemStatus = '999', ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "' where SheetNoItems = '" + vSheetNoItems + "' ";
                                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                                }
                            }
                        }
                        //寫入異動記錄
                        vOptionString = "作廢總務耗材採購單" + Environment.NewLine +
                                        "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                        "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                        "異動種類：整單作廢" + Environment.NewLine +
                                        "ConsSheet_Pur.aspx";
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                        //備份舊資料
                        vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                        vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                        vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                        vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                       "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                       "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                       "       DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                       "select '" + vH_FirstCode + vNewIndex + "', 'ABR', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                       "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                       "       DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                       " where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        //更新資料
                        vSQLStr_Temp = "update ConsSheetA set SheetStatus = '999', StatusDate = GetDate(), ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "' " + Environment.NewLine +
                                       " where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    }
                    else
                    {
                        eErrorMSG_Main.Text = "單據明細已有部份到貨或結案，本單不可作廢！";
                        eErrorMSG_Main.Visible = true;
                    }
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = false;
                }
            }
        }

        /// <summary>
        /// 刪除主檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            int vTempINT;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vOptionString;
            string vSQLStr_Temp;
            string vSheetNoItems;
            string vMaxIndex;
            string vNewIndex;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                string vSheetNo = eSheetNo.Text.Trim();
                try
                {
                    //寫入異動記錄
                    vOptionString = "刪除總務耗材採購單" + Environment.NewLine +
                                    "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：整單刪除" + Environment.NewLine +
                                    "ConsSheet_Pur.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //檢查有沒有明細
                    vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        while (drTemp.Read())
                        {
                            vSheetNoItems = drTemp["SheetNoItems"].ToString().Trim();
                            //備份明細資料
                            vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                            vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                            vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                            vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                           "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, " + Environment.NewLine +
                                           "       Amount, RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, " + Environment.NewLine +
                                           "       ConsUnit, ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                           "select '" + vH_FirstCode + vNewIndex + "', 'DELA', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, " + Environment.NewLine +
                                           "       Amount, RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, " + Environment.NewLine +
                                           "       ConsUnit, GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                           "  from ConsSheetB " + Environment.NewLine +
                                           " where SheetNoItems = '" + vSheetNoItems + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            //刪除明細資料
                            vSQLStr_Temp = "delete ConsSheetB where SheetNoItems = '" + vSheetNoItems + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        }
                    }
                    //備份主檔資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'DELA', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //刪除主檔資料
                    vSQLStr_Temp = "delete ConsSheetA where SheetNo = @SheetNo";
                    sdsConsSheetA_Detail.DeleteCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.DeleteParameters.Clear();
                    sdsConsSheetA_Detail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetA_Detail.Delete();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = false;
                }
            }
        }

        protected void fvConsSheetB_Detail_DataBound(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            string vSQLStr_Temp;
            Label eSheetNoItems;
            string vConsDataURL;
            string vConsDataScript;
            TextBox eConsNo;
            TextBox eConsName;
            TextBox ePrice;
            DropDownList ddlConsUnit;
            TextBox eConsUnit;
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
                        try
                        {
                            eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
                            eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_Edit");
                            eConsUnit = (TextBox)fvConsSheetB_Detail.FindControl("eConsUnit_Edit");
                            ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_Edit");
                            vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsNameID=" + eConsName.ClientID +
                                           "&ConsPriceID=" + ePrice.ClientID + "&ConsUnitID=" + eConsUnit.ClientID;
                            vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                            eConsNo.Attributes["onClick"] = vConsDataScript;
                            ddlConsUnit = (DropDownList)fvConsSheetB_Detail.FindControl("ddlConsUnit_Edit");
                            vSQLStr_Temp = "select Cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                           " union all " + Environment.NewLine +
                                           "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                                           " order by ClassNo";
                            using (SqlConnection connTemp = new SqlConnection(vConnStr))
                            {
                                SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                                connTemp.Open();
                                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                                if (drTemp.HasRows)
                                {
                                    ddlConsUnit.Items.Clear();
                                    while (drTemp.Read())
                                    {
                                        ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                        ddlConsUnit.Items.Add(liTemp);
                                    }
                                }
                            }
                            if (eConsUnit.Text.Trim() != "")
                            {
                                ddlConsUnit.SelectedIndex = ddlConsUnit.Items.IndexOf(ddlConsUnit.Items.FindByValue(eConsUnit.Text.Trim()));
                            }
                            else
                            {
                                ddlConsUnit.SelectedIndex = 0;
                            }
                        }
                        catch (Exception eMessage)
                        {
                            eErrorMSG_Main.Text = eMessage.Message;
                            eErrorMSG_Main.Visible = true;
                        }
                    }
                    break;
                case FormViewMode.Insert:
                    eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_INS");
                    if (eSheetNoItems != null)
                    {
                        try
                        {
                            Label eSheetNoA = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
                            Label eSheetNoB = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_INS");
                            eSheetNoB.Text = eSheetNoA.Text.Trim();
                            eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                            eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_INS");
                            eConsUnit = (TextBox)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                            ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_INS");
                            vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsNameID=" + eConsName.ClientID +
                                           "&ConsPriceID=" + ePrice.ClientID + "&ConsUnitID=" + eConsUnit.ClientID;
                            vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                            eConsNo.Attributes["onClick"] = vConsDataScript;
                            ddlConsUnit = (DropDownList)fvConsSheetB_Detail.FindControl("ddlConsUnit_INS");
                            vSQLStr_Temp = "select Cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                           " union all " + Environment.NewLine +
                                           "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                                           " order by ClassNo";
                            using (SqlConnection connTemp = new SqlConnection(vConnStr))
                            {
                                SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                                connTemp.Open();
                                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                                if (drTemp.HasRows)
                                {
                                    ddlConsUnit.Items.Clear();
                                    while (drTemp.Read())
                                    {
                                        ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                        ddlConsUnit.Items.Add(liTemp);
                                    }
                                }
                            }
                            ddlConsUnit.SelectedIndex = 0;
                            eConsUnit.Text = "";
                        }
                        catch (Exception eMessage)
                        {
                            eErrorMSG_Main.Text = eMessage.Message;
                            eErrorMSG_Main.Visible = true;
                        }
                    }
                    break;
            }
        }

        protected void bbOKB_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_Edit");
            if (eSheetNoItems != null)
            {
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_Edit");
                string vSheetNo = eSheetNo.Text.Trim();
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_Edit");
                string vItems = eItems.Text.Trim();
                //先寫入異動記錄
                try
                {
                    string vOptionString = "修改總務耗材採購單明細" + Environment.NewLine +
                                           "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                           "項次：[ " + vItems + " ]" + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：修改" + Environment.NewLine +
                                           "ConsSheet_Pur.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份舊資料
                    string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    string vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    int vTempINT;
                    string vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //更新資料
                    TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
                    Label ePrice = (Label)fvConsSheetB_Detail.FindControl("ePrice_Edit");
                    TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_Edit");
                    Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmountB_Edit");
                    Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_Edit");
                    TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_Edit");
                    string vConsNo = eConsNo.Text.Trim();
                    double vTempFloat;
                    string vConsUnit = eConsUnit.Text.Trim();
                    string vRemarkB = eRemarkB.Text.Trim();
                    vSQLStr_Temp = "update ConsSheetB " + Environment.NewLine +
                                   "   set ConsNo = @ConsNo, Price = @Price, Quantity = @Quantity, Amount = @Amount, ConsUnit = @ConsUnit, " + Environment.NewLine +
                                   "       RemarkB = @RemarkB, ModifyDate = GetDate(), ModifyMan = @ModifyMan " + Environment.NewLine +
                                   " where SheetNoItems = @SheetNoItems ";
                    sdsConsSheetB_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.UpdateParameters.Clear();
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ConsNo", DbType.String, (vConsNo != "") ? vConsNo : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Price", DbType.Double, double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Quantity", DbType.Double, double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Amount", DbType.Double, double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ConsUnit", DbType.String, (vConsUnit != "") ? vConsUnit : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("RemarkB", DbType.String, (vRemarkB != "") ? vRemarkB : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.Update();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        protected void ePrice_Edit_TextChanged(object sender, EventArgs e)
        {
            double vTempFloat;
            TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_Edit");
            if (ePrice != null)
            {
                TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_Edit");
                Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmountB_Edit");
                double vPrice = double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                double vQuantity = double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                double vAmount = vPrice * vQuantity;
                ePrice.Text = vPrice.ToString();
                eQuantity.Text = vQuantity.ToString();
                eAmount.Text = vAmount.ToString();
            }
        }

        protected void ddlConsUnit_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlConsUnit = (DropDownList)fvConsSheetB_Detail.FindControl("ddlConsUnit_Edit");
            if (ddlConsUnit != null)
            {
                Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_Edit");
                eConsUnit.Text = ddlConsUnit.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 明細新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKB_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_INS");
            if (eSheetNoItems != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_INS");
                string vSheetNo = eSheetNo.Text.Trim();
                if (vSheetNo == "")
                {
                    Label eSheetNoA = (Label)fvConsSheetB_Detail.FindControl("eSheetNoA_List");
                    vSheetNo = eSheetNoA.Text.Trim();
                }
                int vTempINT;
                string vSQLStr_Temp = "select max(Items) MaxIndex from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                string vItems = Int32.TryParse(vMaxIndex, out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                string vSheetNoItems = vSheetNo + vItems;
                TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_INS");
                TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_INS");
                TextBox eConsUnit = (TextBox)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_INS");
                Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmountB_INS");
                double vTempFloat;
                vSQLStr_Temp = "insert into ConsSheetB " + Environment.NewLine +
                               "      (SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, ConsUnit, " + Environment.NewLine +
                               "       RemarkB, ItemStatus, BuMan, BuDate) " + Environment.NewLine +
                               "values(@SheetNoItems, @SheetNo, @Items, @ConsNo, @Price, @Quantity, @Amount, @ConsUnit, " + Environment.NewLine +
                               "       @RemarkB, '000', @BuMan, GetDate()) ";
                try
                {
                    sdsConsSheetB_Detail.InsertCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.InsertParameters.Clear();
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ConsNo", DbType.String, (eConsNo.Text.Trim() != "") ? eConsNo.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Price", DbType.Double, double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Quantity", DbType.Double, double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Amount", DbType.Double, double.TryParse(eAmount.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ConsUnit", DbType.String, (eConsUnit.Text.Trim() != "") ? eConsUnit.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB.Text.Trim() != "") ? eRemarkB.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsConsSheetB_Detail.Insert();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        protected void ePrice_INS_TextChanged(object sender, EventArgs e)
        {
            double vTempFloat;
            TextBox ePrice = (TextBox)fvConsSheetB_Detail.FindControl("ePrice_INS");
            if (ePrice != null)
            {
                TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_INS");
                Label eAmount = (Label)fvConsSheetB_Detail.FindControl("eAmountB_INS");
                double vPrice = double.TryParse(ePrice.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                double vQuantity = double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat : 0.0;
                double vAmount = vPrice * vQuantity;
                ePrice.Text = vPrice.ToString();
                eQuantity.Text = vQuantity.ToString();
                eAmount.Text = vAmount.ToString();
            }
        }

        protected void ddlConsUnit_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlConsUnit = (DropDownList)fvConsSheetB_Detail.FindControl("ddlConsUnit_INS");
            if (ddlConsUnit != null)
            {
                Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                eConsUnit.Text = ddlConsUnit.SelectedValue.Trim();
            }
        }

        /// <summary>
        /// 明細到貨
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbAriveB_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems != null)
            {
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_List");
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_List");
                string vSheetNo = eSheetNo.Text.Trim();
                string vItems = eItems.Text.Trim();
                //寫入異動記錄
                string vOptionString = "修改總務耗材採購單明細" + Environment.NewLine +
                                       "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                       "項次：[ " + vItems + " ]" + Environment.NewLine +
                                       "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                       "異動種類：設定到貨" + Environment.NewLine +
                                       "ConsSheet_Pur.aspx";
                try
                {
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //修改明細狀態
                    string vSQLStr_Temp = "update ConsSheetB set ItemStatus = '001', ModifyMan = '" + vLoginID + "', ModifyDate = GetDate() where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
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
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems != null)
            {
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_List");
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_List");
                string vSheetNo = eSheetNo.Text.Trim();
                string vItems = eItems.Text.Trim();
                //寫入異動記錄
                string vOptionString = "修改總務耗材採購單明細" + Environment.NewLine +
                                       "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                       "項次：[ " + vItems + " ]" + Environment.NewLine +
                                       "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                       "異動種類：明細作廢" + Environment.NewLine +
                                       "ConsSheet_Pur.aspx";
                try
                {
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份舊資料
                    int vTempINT;
                    string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    string vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    string vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'ABR', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //修改明細狀態
                    vSQLStr_Temp = "update ConsSheetB set ItemStatus = '999', ModifyMan = '" + vLoginID + "', ModifyDate = GetDate() where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
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
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems != null)
            {
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_List");
                string vSheetNo = eSheetNo.Text.Trim();
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_List");
                string vItems = eItems.Text.Trim();
                Label eSourceNo = (Label)fvConsSheetB_Detail.FindControl("eSourceNo_List");
                string vSourceNo = eSourceNo.Text.Trim();
                //先寫入異動記錄
                string vOptionString = "刪除總務耗材採購單明細" + Environment.NewLine +
                                       "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                       "項次：[ " + vItems + " ]" + Environment.NewLine +
                                       "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                       "異動種類：明細刪除" + Environment.NewLine +
                                       "ConsSheet_Pur.aspx";
                try
                {
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份舊資料
                    int vTempINT;
                    string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    string vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    string vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'DELB', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //修改來源狀態
                    if (vSourceNo!="")
                    {
                        vSQLStr_Temp = "update ConsSheetB set ItemStatus = '000' where SheetNoItems = '" + vSourceNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    }
                    //刪除明細
                    vSQLStr_Temp = "delete ConsSheetB where SheetNoItems = @SheetNoItems";
                    sdsConsSheetB_Detail.DeleteCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.DeleteParameters.Clear();
                    sdsConsSheetB_Detail.DeleteParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.Delete();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = false;
                }
            }
        }

        /// <summary>
        /// 請購單匯入明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExportDetail_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            string vSQLStr_Temp;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
            {
                Label eSupNo = (Label)fvConsSheetA_Detail.FindControl("eSupNo_List");
                Label eSupName = (Label)fvConsSheetA_Detail.FindControl("eSupName_List");
                string vSupNo = eSupNo.Text.Trim();
                vSQLStr_Temp = "select b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.Quantity, c.StockQty, b.RemarkB, b.Price " + Environment.NewLine +
                               "  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo " + Environment.NewLine +
                               " where b.ItemStatus = '000' " + Environment.NewLine +
                               "   and b.SheetNo like '%BS%' " + Environment.NewLine +
                               "   and c.ConsNo in (select ConsNo from ConsSupList where SupNo = '" + vSupNo + "') " + Environment.NewLine +
                               " union all " + Environment.NewLine +
                               "select b.SheetNoItems, b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.Quantity, c.StockQty, b.RemarkB, b.Price " + Environment.NewLine +
                               "  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo " + Environment.NewLine +
                               " where b.ItemStatus = '000' " + Environment.NewLine +
                               "   and b.SheetNo like '%BS%' " + Environment.NewLine +
                               "   and c.ConsNo not in (select Distinct ConsNo from ConsSupList) ";
                sdsPickup_List.SelectCommand = vSQLStr_Temp;
                sdsPickup_List.Select(new DataSourceSelectArguments());
                gridPickup_List.DataBind();
                plPickupDetail.Visible = true;
                plShowDataB.Visible = false;
            }

        }

        /// <summary>
        /// 全選
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSelectAll_Order_Click(object sender, EventArgs e)
        {
            foreach (GridViewRow vCurrentRows in gridPickup_List.Rows)
            {
                CheckBox cbTemp = (CheckBox)vCurrentRows.FindControl("cbChoise");
                if (cbTemp != null) cbTemp.Checked = true;
            }
        }

        /// <summary>
        /// 取消全選
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbUnselAll_Order_Click(object sender, EventArgs e)
        {
            foreach (GridViewRow vCurrentRows in gridPickup_List.Rows)
            {
                CheckBox cbTemp = (CheckBox)vCurrentRows.FindControl("cbChoise");
                if (cbTemp != null) cbTemp.Checked = false;
            }
        }

        /// <summary>
        /// 取回勾選的請購單明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKC_Order_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                int vTempINT;
                string vSQLStr_Temp;
                string vSheetNoItems;
                string vSourceNo;
                string vMaxIndex;
                string vNewIndex;
                string vSheetNo = eSheetNo.Text.Trim();
                try
                {
                    foreach (GridViewRow vCurrentRows in gridPickup_List.Rows)
                    {
                        CheckBox cbTemp = (CheckBox)vCurrentRows.FindControl("cbChoise");
                        if (cbTemp != null && cbTemp.Checked)
                        {
                            vSQLStr_Temp = "select max(Items) MaxIndex from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                            vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                            vNewIndex = Int32.TryParse(vMaxIndex, out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                            vSheetNoItems = vSheetNo + vNewIndex;
                            vSourceNo = gridPickup_List.DataKeys[vCurrentRows.RowIndex].Value.ToString().Trim();
                            vSQLStr_Temp = "insert into ConsSheetB " + Environment.NewLine +
                                           "      (SheetNoItems, SheetNo, Items, ConsNo, ItemStatus, Quantity, ConsUnit, RemarkB, BuDate, BuMan, SourceNo) " + Environment.NewLine +
                                           "select '" + vSheetNoItems + "', '" + vSheetNo + "', '" + vNewIndex + "', ConsNo, '000', Quantity, ConsUnit, RemarkB, GetDate(), " + Environment.NewLine +
                                           "       '" + vLoginID + "', SheetNoItems " + Environment.NewLine +
                                           "  from ConsSheetB " + Environment.NewLine +
                                           " where SheetNoItems = '" + vSourceNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            vSQLStr_Temp = "update ConsSheetB set ItemStatus = '002' where SheetNoItems = '" + vSourceNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        }
                    }
                    gridConsSheetB_List.DataBind();
                    plShowDataB.Visible = true;
                    plPickupDetail.Visible = false;
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        protected void bbCancelC_Order_Click(object sender, EventArgs e)
        {
            plShowDataB.Visible = true;
            plPickupDetail.Visible = false;
        }

        /// <summary>
        /// 關閉預覽畫面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plShowDataA.Visible = true;
            plShowDataB.Visible = true;
            plSearch.Visible = true;
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
            if (eDepNo != null)
            {
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_Edit");
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo ";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo.Text = vDepNo;
                eDepName.Text = vDepName;
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
            if (eDepNo != null)
            {
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_INS");
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo ";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo.Text = vDepNo;
                eDepName.Text = vDepName;
            }
        }
    }
}