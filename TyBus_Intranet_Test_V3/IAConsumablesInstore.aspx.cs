using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class IAConsumablesInstore : Page
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
                    string vCaseDateURL = "InputDate.aspx?TextboxID=" + eBuDate_S_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eBuDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {

                    }
                    else
                    {
                        //把 OpenData 寫在這裡，確保第一次打開頁面時不會自動查詢資料，但是每次查詢資料或是切換頁面後會更新內容
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
            string vResultStr = "";

            String vBuDate_S = (eBuDate_S_Search.Text.Trim() != "") ? PF.TransDateString(eBuDate_S_Search.Text.Trim(), "B") : DateTime.Now.ToShortDateString();
            string vBuDate_E = (eBuDate_E_Search.Text.Trim() != "") ? PF.TransDateString(eBuDate_E_Search.Text.Trim(), "B") : DateTime.Now.ToShortDateString();

            string vWStr_SheetNo = ((eSheetNo_S_Search.Text.Trim() != "") && (eSheetNo_E_Search.Text.Trim() != "")) ? "   and a.SheetNo between '" + eSheetNo_S_Search.Text.Trim() + "' and '" + eSheetNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                   ((eSheetNo_S_Search.Text.Trim() != "") && (eSheetNo_E_Search.Text.Trim() == "")) ? "   and a.SheetNo = '" + eSheetNo_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                   ((eSheetNo_S_Search.Text.Trim() == "") && (eSheetNo_E_Search.Text.Trim() != "")) ? "   and a.SheetNo = '" + eSheetNo_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_SupNo = (eSupNo_Search.Text.Trim() != "") ? "   and a.SupNo = '" + eSupNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_SheetStatus = (eSheetStatus_Search.Text.Trim() != "") ? "   and a.SheetStatus = '" + eSheetStatus_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuMan = (eBuMan_Search.Text.Trim() != "") ? "   and a.BuMan = '" + eBuMan_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuDate = ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() != "")) ? "   and a.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() == "")) ? "   and a.BuDate = '" + vBuDate_S + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() != "")) ? "   and a.BuDate = '" + vBuDate_E + "' " + Environment.NewLine : "";
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


            vResultStr = "SELECT SheetNo, SupNo, (SELECT name FROM CUSTOM WHERE (types = 'S') AND (code = a.SupNo)) AS SupName, BuDate, TotalAmount, SheetNote, SheetStatus, " + Environment.NewLine +
                         "       (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      SheetStatus') AND (CLASSNO = a.SheetStatus)) AS SheetStatus_C, RemarkA " + Environment.NewLine +
                         "  FROM IASheetA AS a " + Environment.NewLine +
                         " WHERE ISNULL(a.SheetNo, '') <> '' " + Environment.NewLine +
                         "   AND SheetMode = 'SI' " + Environment.NewLine +
                         vWStr_SheetNo +
                         vWStr_SupNo +
                         vWStr_SheetStatus +
                         vWStr_BuMan +
                         vWStr_BuDate +
                         vWStr_SheetNote +
                         vWStr_Substr +
                         " ORDER BY a.SheetNo DESC ";
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSelStr();
            dsIASheetA_List.SelectCommand = vSelectStr;
            gridIASheetList.DataBind();
        }

        protected void eSupNo_Search_TextChanged(object sender, EventArgs e)
        {
            string vSupNo_Temp = eSupNo_Search.Text.Trim();
            string vSQLStr = "select [Name] from Custom where Code = '" + vSupNo_Temp + "' and [Types] = 'S' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
        }

        protected void ddlSheetStatus_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eSheetStatus_Search.Text = ddlSheetStatus_Search.SelectedValue;
        }

        protected void eBuMan_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo_Temp = eBuMan_Search.Text.Trim();
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp.Trim();
                vSQLStr = "select top 1 EmpNo from Employee where [Name] like '%" + vEmpName_Temp + "%' order by EmpNo DESC ";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eBuMan_Search.Text = vEmpNo_Temp.Trim();
            eBuManName_Search.Text = vEmpName_Temp.Trim();
        }

        protected void bbOK_Search_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo_Edit = (Label)fvIASheetA_Detail.FindControl("eSheetNo_Edit");
            if (eSheetNo_Edit != null)
            {
                TextBox eSupNo_Edit = (TextBox)fvIASheetA_Detail.FindControl("eSupNo_Edit");
                Label ePayMode_Edit = (Label)fvIASheetA_Detail.FindControl("ePayMode_Edit");
                Label eTaxType_Edit = (Label)fvIASheetA_Detail.FindControl("eTaxType_Edit");
                TextBox eTaxRate_Edit = (TextBox)fvIASheetA_Detail.FindControl("eTaxRate_Edit");
                TextBox eTaxAMT_Edit = (TextBox)fvIASheetA_Detail.FindControl("eTaxAMT_Edit");
                Label eTotalAmount_Edit = (Label)fvIASheetA_Detail.FindControl("eTotalAmount_Edit");
                TextBox eSheetNote_Edit = (TextBox)fvIASheetA_Detail.FindControl("eSheetNote_Edit");
                TextBox eRemarkA_Edit = (TextBox)fvIASheetA_Detail.FindControl("eRemarkA_Edit");
                try
                {
                    string vSQLStr_Temp = "update IASheetA " + Environment.NewLine +
                                          "   set SupNo = @SupNo, PayMode = @PayMode, TaxType = @TaxType, TaxRate = @TaxRate, TaxAMT = @TaxAMT, TotalAmount = @TotalAmount, " + Environment.NewLine +
                                          "       SheetNote = @SheetNote, RemarkA = @RemarkA, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                          " where SheetNo = @SheetNo";
                    dsIASheetA_Detail.UpdateCommand = vSQLStr_Temp;
                    dsIASheetA_Detail.UpdateParameters.Clear();
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("SupNo", DbType.String, (eSupNo_Edit.Text.Trim() != "") ? eSupNo_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("PayMode", DbType.String, (ePayMode_Edit.Text.Trim() != "") ? ePayMode_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("TaxType", DbType.String, (eTaxType_Edit.Text.Trim() != "") ? eTaxType_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("TaxRate", DbType.Double, (eTaxRate_Edit.Text.Trim() != "") ? eTaxRate_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("TaxAMT", DbType.Int32, (eTaxAMT_Edit.Text.Trim() != "") ? eTaxAMT_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Int32, (eTotalAmount_Edit.Text.Trim() != "") ? eTotalAmount_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("SheetNote", DbType.String, (eSheetNote_Edit.Text.Trim() != "") ? eSheetNote_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA_Edit.Text.Trim() != "") ? eRemarkA_Edit.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    dsIASheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, eSheetNo_Edit.Text.Trim()));
                    dsIASheetA_Detail.Update();
                    gridIASheetList.DataBind();
                    fvIASheetA_Detail.DataBind();
                    fvIASheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
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

        protected void eSupNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eSupNo_Edit = (TextBox)fvIASheetA_Detail.FindControl("eSupNo_Edit");
            Label eSupName_Edit = (Label)fvIASheetA_Detail.FindControl("eSupName_Edit");
            string vSupNo_Temp = eSupNo_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Custom where Code = '" + vSupNo_Temp + "' and Types = 'S' ";
            string vSupName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vSupName_Temp.Trim() == "")
            {
                vSupName_Temp = vSupNo_Temp.Trim();
                vSQLStr_Temp = "select Top 1 Code from Custom where Types = 'S' and [Name] like '%" + vSupName_Temp.Trim() + "%' order by Code";
                vSupNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Code");
            }
            eSupNo_Edit.Text = vSupNo_Temp.Trim();
            eSupName_Edit.Text = vSupName_Temp.Trim();
        }

        protected void ddlPayMode_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlPayMode_Edit = (DropDownList)fvIASheetA_Detail.FindControl("ddlPayMode_Edit");
            Label ePayMode_Edit = (Label)fvIASheetA_Detail.FindControl("ePayMode_Edit");
            ePayMode_Edit.Text = ddlPayMode_Edit.SelectedValue.Trim();
        }

        protected void ddlTaxType_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTaxType_Edit = (DropDownList)fvIASheetA_Detail.FindControl("ddlTaxType_Edit");
            Label eTaxType_Edit = (Label)fvIASheetA_Detail.FindControl("eTaxType_Edit");
            eTaxType_Edit.Text = ddlTaxType_Edit.SelectedValue.Trim();
            TextBox eTaxRate_Edit = (TextBox)fvIASheetA_Detail.FindControl("eTaxRate_Edit");
            if (eTaxRate_Edit.Text.Trim() != "")
            {
                double vTaxRate_Temp = double.Parse(eTaxRate_Edit.Text.Trim()) / 100.0;
                Label eAmount_Edit = (Label)fvIASheetA_Detail.FindControl("eAmount_Edit");
                int vAmount_Temp = (eAmount_Edit.Text.Trim() != "") ? Int32.Parse(eAmount_Edit.Text.Trim()) : 0;
                Label eTotalAmount_Edit = (Label)fvIASheetA_Detail.FindControl("eTotalAmount_Edit");
                TextBox eTaxAMT_Edit = (TextBox)fvIASheetA_Detail.FindControl("eTaxAMT_Edit");
                int vTaxAMT_Temp = 0;
                switch (ddlTaxType_Edit.SelectedValue)
                {
                    case "1": //應稅
                        vTaxAMT_Temp = (int)Math.Round(vAmount_Temp * vTaxRate_Temp, 0, MidpointRounding.AwayFromZero);
                        eTaxAMT_Edit.Text = vTaxAMT_Temp.ToString();
                        eTotalAmount_Edit.Text = (vAmount_Temp + vTaxAMT_Temp).ToString();
                        eTaxAMT_Edit.Enabled = true;
                        eTaxRate_Edit.Enabled = true;
                        break;
                    case "2": //應稅內含
                        eTotalAmount_Edit.Text = eAmount_Edit.Text.Trim();
                        vTaxAMT_Temp = (int)Math.Round((vAmount_Temp / (1 + vTaxRate_Temp)), 0, MidpointRounding.AwayFromZero);
                        eTaxAMT_Edit.Text = vTaxAMT_Temp.ToString("D0");
                        eAmount_Edit.Text = (vAmount_Temp - vTaxAMT_Temp).ToString();
                        eTaxAMT_Edit.Enabled = true;
                        eTaxRate_Edit.Enabled = true;
                        break;
                    case "3": //零稅率
                        eTotalAmount_Edit.Text = eAmount_Edit.Text.Trim();
                        eTaxAMT_Edit.Text = "0";
                        eTaxRate_Edit.Text = "0";
                        eTaxAMT_Edit.Enabled = false;
                        eTaxRate_Edit.Enabled = false;
                        break;
                    case "4": //免稅
                        eTotalAmount_Edit.Text = eAmount_Edit.Text.Trim();
                        eTaxAMT_Edit.Text = "0";
                        eTaxRate_Edit.Text = "";
                        eTaxAMT_Edit.Enabled = false;
                        eTaxRate_Edit.Enabled = false;
                        break;
                }
            }
        }

        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "";
            Label eSheetNo_INS = (Label)fvIASheetA_Detail.FindControl("eSheetNo_INS");
            if (eSheetNo_INS != null)
            {
                //產生單號
                string vSheetNo = PF.GetIASheetNo(vConnStr, "SI");


                TextBox eSupNo_INS = (TextBox)fvIASheetA_Detail.FindControl("eSupNo_INS");
                Label ePayMode_INS = (Label)fvIASheetA_Detail.FindControl("ePayMode_INS");
                Label eTaxType_INS = (Label)fvIASheetA_Detail.FindControl("eTaxType_INS");
                TextBox eTaxRate_INS = (TextBox)fvIASheetA_Detail.FindControl("eTaxRate_INS");
                TextBox eSheetNote_INS = (TextBox)fvIASheetA_Detail.FindControl("eSheetNote_INS");
                TextBox eRemarkA_INS = (TextBox)fvIASheetA_Detail.FindControl("eRemarkA_INS");
                try
                {
                    vSQLStr_Temp = "insert into IASheetA " + Environment.NewLine +
                                   "       (SheetNo, SheetMode, SupNo, PayMode, TaxType, TaxRate, TaxAMT, Amount, TotalAmount, " + Environment.NewLine +
                                   "        SheetNote, RemarkA, BuMan, BuDate, SheetStatus) " + Environment.NewLine +
                                   "values (@SheetNo, 'SI', @SupNo, @PayMode, @TaxType, @TaxRate, 0, 0, 0, @SheetNote, @RemarkA, @BuMan, GetDate(), '000')";
                    dsIASheetA_Detail.InsertCommand = vSQLStr_Temp;
                    dsIASheetA_Detail.InsertParameters.Clear();
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("SupNo", DbType.String, (eSupNo_INS.Text.Trim() != "") ? eSupNo_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("PayMode", DbType.String, (ePayMode_INS.Text.Trim() != "") ? ePayMode_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("TaxType", DbType.String, (eTaxType_INS.Text.Trim() != "") ? eTaxType_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("TaxRate", DbType.Double, (eTaxRate_INS.Text.Trim() != "") ? eTaxRate_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("SheetNote", DbType.String, (eSheetNote_INS.Text.Trim() != "") ? eSheetNote_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA_INS.Text.Trim() != "") ? eRemarkA_INS.Text.Trim() : String.Empty));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    dsIASheetA_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    dsIASheetA_Detail.Insert();
                    gridIASheetList.DataBind();
                    fvIASheetA_Detail.DataBind();
                    fvIASheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
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

        protected void eSupNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eSupNo_INS = (TextBox)fvIASheetA_Detail.FindControl("eSupNo_INS");
            Label eSupName_INS = (Label)fvIASheetA_Detail.FindControl("eSupName_INS");
            string vSupNo_Temp = eSupNo_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Custom where Code = '" + vSupNo_Temp + "' and Types = 'S' ";
            string vSupName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vSupName_Temp.Trim() == "")
            {
                vSupName_Temp = vSupNo_Temp.Trim();
                vSQLStr_Temp = "select Top 1 Code from Custom where Types = 'S' and [Name] like '%" + vSupName_Temp.Trim() + "%' order by Code";
                vSupNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Code");
            }
            eSupNo_INS.Text = vSupNo_Temp.Trim();
            eSupName_INS.Text = vSupName_Temp.Trim();
        }

        protected void ddlPayMode_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlPayMode_INS = (DropDownList)fvIASheetA_Detail.FindControl("ddlPayMode_INS");
            Label ePayMode_INS = (Label)fvIASheetA_Detail.FindControl("ePayMode_INS");
            ePayMode_INS.Text = ddlPayMode_INS.SelectedValue.Trim();
        }

        protected void ddlTaxType_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTaxType_INS = (DropDownList)fvIASheetA_Detail.FindControl("ddlTaxType_INS");
            Label eTaxType_INS = (Label)fvIASheetA_Detail.FindControl("eTaxType_INS");
            eTaxType_INS.Text = ddlTaxType_INS.SelectedValue.Trim();
            TextBox eTaxRate_INS = (TextBox)fvIASheetA_Detail.FindControl("eTaxRate_INS");
            if (eTaxRate_INS.Text.Trim() != "")
            {
                double vTaxRate_Temp = double.Parse(eTaxRate_INS.Text.Trim()) / 100.0;
                Label eAmount_INS = (Label)fvIASheetA_Detail.FindControl("eAmount_INS");
                int vAmount_Temp = (eAmount_INS.Text.Trim() != "") ? Int32.Parse(eAmount_INS.Text.Trim()) : 0;
                Label eTotalAmount_INS = (Label)fvIASheetA_Detail.FindControl("eTotalAmount_INS");
                TextBox eTaxAMT_INS = (TextBox)fvIASheetA_Detail.FindControl("eTaxAMT_INS");
                int vTaxAMT_Temp = 0;
                switch (ddlTaxType_INS.SelectedValue)
                {
                    case "1": //應稅
                        vTaxAMT_Temp = (int)Math.Round(vAmount_Temp * vTaxRate_Temp, 0, MidpointRounding.AwayFromZero);
                        eTaxAMT_INS.Text = vTaxAMT_Temp.ToString();
                        eTotalAmount_INS.Text = (vAmount_Temp + vTaxAMT_Temp).ToString();
                        break;
                    case "2": //應稅內含
                        eTotalAmount_INS.Text = eAmount_INS.Text.Trim();
                        vTaxAMT_Temp = (int)Math.Round((vAmount_Temp / (1 + vTaxRate_Temp)), 0, MidpointRounding.AwayFromZero);
                        eTaxAMT_INS.Text = vTaxAMT_Temp.ToString("D0");
                        eAmount_INS.Text = (vAmount_Temp - vTaxAMT_Temp).ToString();
                        eTaxAMT_INS.Enabled = true;
                        eTaxRate_INS.Enabled = true;
                        break;
                    case "3": //零稅率
                        eTotalAmount_INS.Text = eAmount_INS.Text.Trim();
                        eTaxAMT_INS.Text = "0";
                        eTaxRate_INS.Text = "0";
                        eTaxAMT_INS.Enabled = false;
                        eTaxRate_INS.Enabled = false;
                        break;
                    case "4": //免稅
                        eTotalAmount_INS.Text = eAmount_INS.Text.Trim();
                        eTaxAMT_INS.Text = "0";
                        eTaxRate_INS.Text = "";
                        eTaxAMT_INS.Enabled = false;
                        eTaxRate_INS.Enabled = false;
                        break;
                }
            }
        }

        protected void fvIASheetA_Detail_DataBound(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSupDataURL = "";
            string vSupDataScript = "";

            switch (fvIASheetA_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    //2021.12.23 新增
                    gridIASheetB_List.Visible = true;
                    fvIASheetB_Detail.Visible = true;

                    Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
                    Label eSheetStatus_List = (Label)fvIASheetA_Detail.FindControl("eSheetStatus_List");
                    Button bbIsArrived_List = (Button)fvIASheetA_Detail.FindControl("bbIsArrived_List");
                    Button bbIsPaid_List = (Button)fvIASheetA_Detail.FindControl("bbIsPaid_List");
                    Button bbIsAbort_List = (Button)fvIASheetA_Detail.FindControl("bbIsAbort_List");
                    Button bbIsClosed_List = (Button)fvIASheetA_Detail.FindControl("bbIsClosed_List");
                    Button bbDelete_List = (Button)fvIASheetA_Detail.FindControl("bbDelete_List");
                    Button bbModify_List = (Button)fvIASheetA_Detail.FindControl("bbModify_List");

                    if (eSheetNo_List != null)
                    {
                        switch (eSheetStatus_List.Text.Trim())
                        {
                            case "000": //已開單
                                bbIsArrived_List.Enabled = true;
                                bbIsPaid_List.Enabled = true;
                                bbIsAbort_List.Enabled = true;
                                bbIsClosed_List.Enabled = true;
                                break;
                            case "010": //已付款
                            case "002": //已到貨
                                bbIsArrived_List.Enabled = false;
                                bbIsPaid_List.Enabled = false;
                                bbIsAbort_List.Enabled = true;
                                bbIsClosed_List.Enabled = true;
                                break;
                            case "003": //部份到貨
                                bbIsArrived_List.Enabled = true;
                                bbIsPaid_List.Enabled = false;
                                bbIsAbort_List.Enabled = true;
                                bbIsClosed_List.Enabled = true;
                                break;
                            case "998": //已結案
                            case "999": //已作廢
                                bbModify_List.Enabled = false;
                                bbDelete_List.Enabled = false;
                                break;
                            default:
                                bbIsArrived_List.Enabled = false;
                                bbIsPaid_List.Enabled = false;
                                bbIsAbort_List.Enabled = false;
                                bbIsClosed_List.Enabled = false;
                                break;
                        }
                    }
                    break;
                case FormViewMode.Edit:
                    //2021.12.23 新增
                    gridIASheetB_List.Visible = false;
                    fvIASheetB_Detail.Visible = false;

                    Label eSheetNo_Edit = (Label)fvIASheetA_Detail.FindControl("eSheetNo_Edit");
                    if (eSheetNo_Edit != null)
                    {
                        Label eModifyMan_Edit = (Label)fvIASheetA_Detail.FindControl("eModifyMan_Edit");
                        /*
                        eModifyMan_Edit.Text = vLoginID;
                        Label eModifyManName_Edit = (Label)fvIASheetA_Detail.FindControl("eModifyManName_Edit");
                        eModifyManName_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + eModifyMan_Edit.Text.Trim() + "' ", "Name");
                        Label eModifyDate_Edit = (Label)fvIASheetA_Detail.FindControl("eModifyDate_Edit");
                        eModifyDate_Edit.Text = DateTime.Now.ToShortDateString(); //*/

                        DropDownList ddlPayMode_Edit = (DropDownList)fvIASheetA_Detail.FindControl("ddlPayMode_Edit");
                        Label ePayMode_Edit = (Label)fvIASheetA_Detail.FindControl("ePayMode_Edit");
                        int vRCount_Edit = 0;
                        for (int i = 0; i < ddlPayMode_Edit.Items.Count; i++)
                        {
                            if (ddlPayMode_Edit.Items[i].Value == ePayMode_Edit.Text.Trim())
                            {
                                vRCount_Edit = i;
                            }
                        }
                        ddlPayMode_Edit.SelectedIndex = vRCount_Edit;

                        DropDownList ddlTaxType_Edit = (DropDownList)fvIASheetA_Detail.FindControl("ddlTaxType_Edit");
                        Label eTaxType_Edit = (Label)fvIASheetA_Detail.FindControl("eTaxType_Edit");
                        vRCount_Edit = 0;
                        for (int i = 0; i < ddlTaxType_Edit.Items.Count; i++)
                        {
                            if (ddlTaxType_Edit.Items[i].Value == eTaxType_Edit.Text.Trim())
                            {
                                vRCount_Edit = i;
                            }
                        }
                        ddlTaxType_Edit.SelectedIndex = vRCount_Edit;
                    }
                    TextBox eSupNo_Edit = (TextBox)fvIASheetA_Detail.FindControl("eSupNo_Edit");
                    vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo_Edit.ClientID;
                    vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                    eSupNo_Edit.Attributes["onClick"] = vSupDataScript;
                    break;
                case FormViewMode.Insert:
                    //2021.12.23 新增
                    gridIASheetB_List.Visible = false;
                    fvIASheetB_Detail.Visible = false;

                    Label eSheetNo_INS = (Label)fvIASheetA_Detail.FindControl("eSheetNo_INS");
                    if (eSheetNo_INS != null)
                    {
                        Label eBuMan_INS = (Label)fvIASheetA_Detail.FindControl("eBuMan_INS");
                        eBuMan_INS.Text = vLoginID;
                        Label eBuManName_INS = (Label)fvIASheetA_Detail.FindControl("eBuManNAme_INS");
                        eBuManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                        Label eBuDate_INS = (Label)fvIASheetA_Detail.FindControl("eBuDate_INS");
                        eBuDate_INS.Text = DateTime.Now.ToShortDateString();

                        DropDownList ddlPayMode_INS = (DropDownList)fvIASheetA_Detail.FindControl("ddlPayMode_INS");
                        Label ePayMode_INS = (Label)fvIASheetA_Detail.FindControl("ePayMode_INS");
                        int vRCount_INS = 0;
                        for (int i = 0; i < ddlPayMode_INS.Items.Count; i++)
                        {
                            if (ddlPayMode_INS.Items[i].Value == ePayMode_INS.Text.Trim())
                            {
                                vRCount_INS = i;
                            }
                        }
                        ddlPayMode_INS.SelectedIndex = vRCount_INS;

                        DropDownList ddlTaxType_INS = (DropDownList)fvIASheetA_Detail.FindControl("ddlTaxType_INS");
                        Label eTaxType_INS = (Label)fvIASheetA_Detail.FindControl("eTaxType_INS");
                        vRCount_INS = 0;
                        for (int i = 0; i < ddlTaxType_INS.Items.Count; i++)
                        {
                            if (ddlTaxType_INS.Items[i].Value == eTaxType_INS.Text.Trim())
                            {
                                vRCount_INS = i;
                            }
                        }
                        ddlTaxType_INS.SelectedIndex = vRCount_INS;
                    }
                    TextBox eSupNo_INS = (TextBox)fvIASheetA_Detail.FindControl("eSupNo_INS");
                    vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo_INS.ClientID;
                    vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                    eSupNo_INS.Attributes["onClick"] = vSupDataScript;
                    break;
                default:
                    break;
            }
        }

        protected void eTaxAMT_Edit_TextChanged(object sender, EventArgs e)
        {
            DropDownList ddlTaxType_Edit = (DropDownList)fvIASheetA_Detail.FindControl("ddlTaxType_Edit");
            if (ddlTaxType_Edit != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                TextBox eTaxAMT_Edit = (TextBox)fvIASheetA_Detail.FindControl("eTaxAMT_Edit");
                TextBox eTaxRate_Edit = (TextBox)fvIASheetA_Detail.FindControl("eTaxRate_Edit");
                Label eSheetNo_Edit = (Label)fvIASheetA_Detail.FindControl("eSheetNo_Edit");
                Label eAmount_Edit = (Label)fvIASheetA_Detail.FindControl("eAmount_Edit");
                Label eTotalAmount_Edit = (Label)fvIASheetA_Detail.FindControl("eTotalAmount_Edit");
                string vSheetNo = eSheetNo_Edit.Text.Trim();
                string vSQLStr_Temp = "select sum(Amount) TAmount from IASheetB where SheetNo = '" + vSheetNo + "' ";
                string vTamount_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "TAmount");
                int vTAmount = (vTamount_Str.Trim() != "") ? Int32.Parse(vTamount_Str) : 0;
                int vTaxAMT = (eTaxAMT_Edit.Text.Trim() != "") ? Int32.Parse(eTaxAMT_Edit.Text.Trim()) : 0;
                double vTaxRate = (eTaxRate_Edit.Text.Trim() != "") ? double.Parse(eTaxRate_Edit.Text.Trim()) / 100.0 : 0.05; //預設稅率 5%
                switch (ddlTaxType_Edit.SelectedValue)
                {
                    case "1": //應稅
                        eAmount_Edit.Text = vTAmount.ToString();
                        if (vTaxAMT == 0)
                        {
                            vTaxAMT = (int)Math.Round((double)vTAmount * vTaxRate, MidpointRounding.AwayFromZero);
                            eTaxAMT_Edit.Text = (vTaxAMT * 100).ToString();
                        }
                        eTotalAmount_Edit.Text = (vTAmount + vTaxAMT).ToString();
                        break;
                    case "2": //應稅內含
                        eTotalAmount_Edit.Text = vTAmount.ToString();
                        if (vTaxAMT == 0)
                        {
                            vTaxAMT = (int)Math.Round((double)vTAmount / (1.0 + vTaxRate), MidpointRounding.AwayFromZero);
                            eTaxAMT_Edit.Text = (vTaxAMT * 100).ToString();
                        }
                        eAmount_Edit.Text = (vTAmount - vTaxAMT).ToString();
                        break;
                    default: //零稅或免稅
                        eTotalAmount_Edit.Text = vTAmount.ToString();
                        eAmount_Edit.Text = vTAmount.ToString();
                        eTaxAMT_Edit.Text = "0";
                        break;
                }
            }
        }

        protected void eTaxAMT_INS_TextChanged(object sender, EventArgs e)
        {
            DropDownList ddlTaxType_INS = (DropDownList)fvIASheetA_Detail.FindControl("ddlTaxType_INS");
            if (ddlTaxType_INS != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                TextBox eTaxAMT_INS = (TextBox)fvIASheetA_Detail.FindControl("eTaxAMT_INS");
                TextBox eTaxRate_INS = (TextBox)fvIASheetA_Detail.FindControl("eTaxRate_INS");
                Label eSheetNo_INS = (Label)fvIASheetA_Detail.FindControl("eSheetNo_INS");
                Label eAmount_INS = (Label)fvIASheetA_Detail.FindControl("eAmount_INS");
                Label eTotalAmount_INS = (Label)fvIASheetA_Detail.FindControl("eTotalAmount_INS");
                string vSheetNo = eSheetNo_INS.Text.Trim();
                string vSQLStr_Temp = "select sum(Amount) TAmount from IASheetB where SheetNo = '" + vSheetNo + "' ";
                string vTamount_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "TAmount");
                int vTAmount = (vTamount_Str.Trim() != "") ? Int32.Parse(vTamount_Str) : 0;
                int vTaxAMT = (eTaxAMT_INS.Text.Trim() != "") ? Int32.Parse(eTaxAMT_INS.Text.Trim()) : 0;
                double vTaxRate = (eTaxRate_INS.Text.Trim() != "") ? double.Parse(eTaxRate_INS.Text.Trim()) / 100.0 : 0.05; //預設稅率 5%
                switch (ddlTaxType_INS.SelectedValue)
                {
                    case "1": //應稅
                        eAmount_INS.Text = vTAmount.ToString();
                        if (vTaxAMT == 0)
                        {
                            vTaxAMT = (int)Math.Round((double)vTAmount * vTaxRate, MidpointRounding.AwayFromZero);
                            eTaxAMT_INS.Text = (vTaxAMT * 100).ToString();
                        }
                        eTotalAmount_INS.Text = (vTAmount + vTaxAMT).ToString();
                        break;
                    case "2": //應稅內含
                        eTotalAmount_INS.Text = vTAmount.ToString();
                        if (vTaxAMT == 0)
                        {
                            vTaxAMT = (int)Math.Round((double)vTAmount / (1.0 + vTaxRate), MidpointRounding.AwayFromZero);
                            eTaxAMT_INS.Text = (vTaxAMT * 100).ToString();
                        }
                        eAmount_INS.Text = (vTAmount - vTaxAMT).ToString();
                        break;
                    default: //零稅或免稅
                        eTotalAmount_INS.Text = vTAmount.ToString();
                        eAmount_INS.Text = vTAmount.ToString();
                        eTaxAMT_INS.Text = "0";
                        break;
                }
            }
        }

        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "";
            Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
            if (eSheetNo_List != null)
            {
                try
                {
                    string vSheetNo = eSheetNo_List.Text.Trim();
                    string vSheetNoItems = "";
                    //檢查有沒有明細項目
                    //vSQLStr_Temp = "select count(SheetNo) RCount from IASheetB where SheetNo = '" + vSheetNo + "' ";
                    //int vRCount = int.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount"));
                    //if (vRCount > 0)
                    //{
                    /* 2022.07.14 計算數量交給資料庫的 Trigger 做
                    //先計算每一個耗材扣除這個單據之後的總數量
                    using (SqlConnection connDelDetail = new SqlConnection())
                    {
                        connDelDetail.ConnectionString = vConnStr;
                        SqlCommand cmdDelDetail = new SqlCommand("select ConsNo from IASheetB where SheetNo = '" + vSheetNo + "' order by Items ", connDelDetail);
                        connDelDetail.Open();
                        SqlDataReader drDelDetail = cmdDelDetail.ExecuteReader();
                        while (drDelDetail.Read())
                        {
                            PF.CalIAStoreQuantity(drDelDetail["ConsNo"].ToString().Trim(), vConnStr, vSheetNo);
                        }
                    } //*/
                    //再實際刪除明細
                    //vSQLStr_Temp = "delete IASheetB where SheetNo = '" + vSheetNo + "' ";
                    //PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //}
                    //2022.08.30 更換做法
                    using (SqlConnection connDelete = new SqlConnection())
                    {
                        connDelete.ConnectionString = vConnStr;
                        SqlCommand cmdDelete = new SqlCommand("select SheetNoItems from IASheetB where SheetNo = '" + vSheetNo + "' order by Items", connDelete);
                        connDelete.Open();
                        SqlDataReader drDelete = cmdDelete.ExecuteReader();
                        while (drDelete.Read())
                        {
                            vSheetNoItems = drDelete["SheetNoItems"].ToString().Trim();
                            vSQLStr_Temp = "delete IASheetB where SheetNoItems = '" + vSheetNoItems + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        }
                    }
                    //確定沒有明細之後把主檔刪除
                    dsIASheetA_Detail.DeleteParameters.Clear();
                    dsIASheetA_Detail.DeleteCommand = "delete IASheetA where SheetNo = @SheetNo";
                    dsIASheetA_Detail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    dsIASheetA_Detail.Delete();
                    gridIASheetList.DataBind();
                    fvIASheetA_Detail.DataBind();
                    fvIASheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
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
        /// 設定整張進貨單到貨
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbIsArrived_List_Click(object sender, EventArgs e) //已到貨
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "";
            string vSheetStatus = "";
            //string vConsNo = "";
            Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
            if (eSheetNo_List != null)
            {
                string vSheetNo = eSheetNo_List.Text.Trim();
                string vSheetNoItems = "";
                string vQuantity = "";
                if (vSheetNo != "")
                {
                    vSQLStr_Temp = "select SheetStatus from IASheetA where SheetNo = '" + vSheetNo + "' ";
                    vSheetStatus = PF.GetValue(vConnStr, vSQLStr_Temp, "SheetStatus").Trim();
                    if ((vSheetStatus == "000") || (vSheetStatus == "020")) //單頭狀態是 "已開單" 或 "部份到貨"
                    {
                        try
                        {
                            //vSQLStr_Temp = "update IASheetB set ItemStatus = '001' where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '') = '' ";
                            //PF.ExecSQL(vConnStr, vSQLStr_Temp); //把明細沒有作廢或到貨的都設為已到貨
                            using (SqlConnection connChangeTypeB = new SqlConnection())
                            {
                                connChangeTypeB.ConnectionString = vConnStr;
                                //2023.01.12 修改
                                //SqlCommand cmdChangeTypeB = new SqlCommand("select SheetNoItems from IASheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '') = '' order by Items", connChangeTypeB);
                                vSQLStr_Temp = "select SheetNoItems, Quantity from IASheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '000') = '000' order by Items";
                                SqlCommand cmdChangeTypeB = new SqlCommand(vSQLStr_Temp, connChangeTypeB);
                                //============================================================================================
                                connChangeTypeB.Open();
                                SqlDataReader drChangeTypeB = cmdChangeTypeB.ExecuteReader();
                                while (drChangeTypeB.Read())
                                {
                                    vSheetNoItems = drChangeTypeB["SheetNoItems"].ToString().Trim();
                                    //2023.01.12 修改
                                    vQuantity = drChangeTypeB["Quantity"].ToString().Trim();
                                    //vSQLStr_Temp = "update IASheetB set ItemStatus = '001' where SheetNoItems = '" + vSheetNoItems + "' ";
                                    vSQLStr_Temp = "update IASheetB " + Environment.NewLine +
                                                   "   set ItemStatus = '001', LastIndate = GetDate(), LastInQty = " + vQuantity + ", " + Environment.NewLine +
                                                   "       ModifyDate = GetDate(), ModifyMan  = '" + vLoginID + "' " + Environment.NewLine +
                                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                                }
                            }
                            vSQLStr_Temp = "update IASheetA set SheetStatus = '010' where SheetNo = '" + vSheetNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp); //把單頭設為已到貨
                        }
                        catch (Exception eMessage)
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('" + eMessage.Message + "')");
                            Response.Write("</" + "Script>");
                        }
                    }
                    else if (vSheetStatus == "001") //單頭狀態是 "已付款"
                    {
                        try
                        {
                            //vSQLStr_Temp = "update IASheetB set ItemStatus = '001' where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '') = '' ";
                            //PF.ExecSQL(vConnStr, vSQLStr_Temp); //把明細沒有作廢或到貨的都設為已到貨
                            using (SqlConnection connChangeTypeB = new SqlConnection())
                            {
                                connChangeTypeB.ConnectionString = vConnStr;
                                //2023.01.12 修正
                                //SqlCommand cmdChangeTypeB = new SqlCommand("select SheetNoItems from IASheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '') = '' order by Items", connChangeTypeB);
                                vSQLStr_Temp = "select SheetNoItems, Quantity from IASheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '000') = '000' order by Items";
                                SqlCommand cmdChangeTypeB = new SqlCommand(vSQLStr_Temp, connChangeTypeB);
                                //==========================================================================
                                connChangeTypeB.Open();
                                SqlDataReader drChangeTypeB = cmdChangeTypeB.ExecuteReader();
                                while (drChangeTypeB.Read())
                                {
                                    vSheetNoItems = drChangeTypeB["SheetNoItems"].ToString().Trim();
                                    //2023.01.12 修改
                                    vQuantity = drChangeTypeB["Quantity"].ToString().Trim();
                                    //vSQLStr_Temp = "update IASheetB set ItemStatus = '001' where SheetNoItems = '" + vSheetNoItems + "' ";
                                    vSQLStr_Temp = "update IASheetB " + Environment.NewLine +
                                                   "   set ItemStatus = '001', LastIndate = GetDate(), LastInQty = " + vQuantity + ", " + Environment.NewLine +
                                                   "       ModifyDate = GetDate(), ModifyMan  = '" + vLoginID + "' " + Environment.NewLine +
                                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                                }
                            }
                            vSQLStr_Temp = "update IASheetA set SheetStatus = '998' where SheetNo = '" + vSheetNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp); //因為付款跟到貨都完成了，把單頭設為已結案
                        }
                        catch (Exception eMessage)
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('" + eMessage.Message + "')");
                            Response.Write("</" + "Script>");
                        }
                    }
                    //計算並更新耗材數量
                    /* 2022.07.14 計算數量交給資料庫的 Trigger 做
                    using (SqlConnection connCalQuantity = new SqlConnection())
                    {
                        connCalQuantity.ConnectionString = vConnStr;
                        SqlCommand cmdCalQty = new SqlCommand("select ConsNo from IASheetB where SheetNo = '" + vSheetNo + "' ", connCalQuantity);
                        connCalQuantity.Open();
                        SqlDataReader drCalQty = cmdCalQty.ExecuteReader();
                        while (drCalQty.Read())
                        {
                            vConsNo = drCalQty["ConsNo"].ToString().Trim();
                            PF.CalIAStoreQuantity(vConsNo, vConnStr, "");
                        }
                    } //*/
                }
            }
        }

        protected void bbIsPaid_List_Click(object sender, EventArgs e) //已付款
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "";
            string vSheetStatus = "";
            Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
            if (eSheetNo_List != null)
            {
                string vSheetNo = eSheetNo_List.Text.Trim();
                if (vSheetNo != "")
                {
                    vSQLStr_Temp = "select SheetStatus from IASheetA where SheetNo = '" + vSheetNo + "' ";
                    vSheetStatus = PF.GetValue(vConnStr, vSQLStr_Temp, "SheetStatus").Trim();
                    if ((vSheetStatus == "000") || (vSheetStatus == "020")) //單頭狀態是 "已開單" 或 "部份到貨"
                    {
                        vSQLStr_Temp = "update IASheetA set SheetStatus = '001' where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp); //把單頭設為已付款
                    }
                    else if (vSheetStatus == "010") //單頭狀態是 "已到貨"
                    {
                        vSQLStr_Temp = "update IASheetA set SheetStatus = '998' where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp); //因為付款跟到貨都完成了，把單頭設為已結案
                    }
                }
            }
        }

        protected void bbIsAbort_List_Click(object sender, EventArgs e) //本單作廢
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "";
            int vRCount = 0;
            Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
            if (eSheetNo_List != null)
            {
                string vSheetNo = eSheetNo_List.Text.Trim();
                if (vSheetNo != "")
                {
                    vSQLStr_Temp = "select count(SheetNo) RCount from IASheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '999') = '001'";
                    if (int.TryParse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount"), out vRCount) && vRCount > 0)
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('進貨單號 [" + vSheetNo + "] 有部份明細已到貨，不可作廢！')");
                        Response.Write("</" + "Script>");
                    }
                    else
                    {
                        vSQLStr_Temp = "update IASheetB set ItemStatus = '999' where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp); //先把明細都作廢
                        vSQLStr_Temp = "update IASheetA set SheetStatus = '999' where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp); //再把單頭作廢
                    }
                }
            }
        }

        protected void bbIsClosed_List_Click(object sender, EventArgs e) //本單結案
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr_Temp = "";
            Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
            if (eSheetNo_List != null)
            {
                string vSheetNo = eSheetNo_List.Text.Trim();
                if (vSheetNo != "")
                {
                    //vSQLStr_Temp = "update IASheetB set ItemStatus = '999' where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '') = '' ";
                    vSQLStr_Temp = "update IASheetB set ItemStatus = '999' where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '000') = '000' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp); //先把明細沒有到貨的設為作廢
                    vSQLStr_Temp = "update IASheetA set SheetStatus = '998' where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp); //再把單頭結案
                }
            }
        }

        protected void fvIASheetBDetail_DataBound(object sender, EventArgs e)
        {
            string vProductDataURL = "";
            string vProductDataScript = "";
            switch (fvIASheetB_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Label eSheetNoItems_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_List");
                    if (eSheetNoItems_List != null)
                    {

                    }
                    break;
                case FormViewMode.Edit:
                    Label eSheetNoItems_Edit = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_Edit");
                    if (eSheetNoItems_Edit != null)
                    {
                        Label eModifyDateB_Edit = (Label)fvIASheetB_Detail.FindControl("eModifyDateB_Edit");
                        eModifyDateB_Edit.Text = DateTime.Now.ToShortDateString();
                        Label eModifyManB_Edit = (Label)fvIASheetB_Detail.FindControl("eModifyManB_Edit");
                        Label eModifyManNameB_Edit = (Label)fvIASheetB_Detail.FindControl("eModifyManNameB_Edit");
                        eModifyManB_Edit.Text = vLoginID.Trim();
                        eModifyManNameB_Edit.Text = vLoginName.Trim();
                        TextBox eConsNo_Edit = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_Edit");
                        //2023.01.12 修改
                        TextBox ePrice_Edit = (TextBox)fvIASheetB_Detail.FindControl("ePrice_Edit");
                        //vProductDataURL = "SearchProduct.aspx?TextBoxID=" + eConsNo_Edit.ClientID;
                        vProductDataURL = "SearchProduct.aspx?ConsNoID=" + eConsNo_Edit.ClientID + "&ConsPriceID=" + ePrice_Edit.ClientID;
                        //============================================================================================
                        vProductDataScript = "window.open('" + vProductDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eConsNo_Edit.Attributes["onClick"] = vProductDataScript;
                    }
                    break;
                case FormViewMode.Insert:
                    Label eSheetNo_List = (Label)fvIASheetA_Detail.FindControl("eSheetNo_List");
                    if (eSheetNo_List != null)
                    {
                        Label eSheetNoItems_INS = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_INS");
                        Label eSheetNoB_INS = (Label)fvIASheetB_Detail.FindControl("eSheetNoB_INS");
                        eSheetNoB_INS.Text = eSheetNo_List.Text.Trim();
                        Label eBuDateB_INS = (Label)fvIASheetB_Detail.FindControl("eBuDateB_INS");
                        eBuDateB_INS.Text = DateTime.Now.ToShortDateString();
                        Label eBuManB_INS = (Label)fvIASheetB_Detail.FindControl("eBuManB_INS");
                        Label eBuManNameB_INS = (Label)fvIASheetB_Detail.FindControl("eBuManNameB_INS");
                        eBuManB_INS.Text = vLoginID.Trim();
                        eBuManNameB_INS.Text = vLoginName.Trim();
                        TextBox eConsNo_INS = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_INS");
                        //2023.01.12 修改
                        TextBox ePrice_INS = (TextBox)fvIASheetB_Detail.FindControl("ePrice_INS");
                        //vProductDataURL = "SearchProduct.aspx?TextBoxID=" + eConsNo_INS.ClientID;
                        vProductDataURL = "SearchProduct.aspx?ConsNoID=" + eConsNo_INS.ClientID + "&ConsPriceID=" + ePrice_INS.ClientID;
                        //============================================================================================
                        vProductDataScript = "window.open('" + vProductDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        Button bbGetConsNo_INS = (Button)fvIASheetB_Detail.FindControl("bbGetConsNo_INS");
                        bbGetConsNo_INS.Attributes["onClick"] = vProductDataScript;
                    }
                    break;
                default:
                    break;
            }
        }

        protected void bbOKB_Edit_Click(object sender, EventArgs e)
        {
            Label eSheetNoItems_Edit = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_Edit");
            if (eSheetNoItems_Edit != null)
            {
                string vSheetNoItems = eSheetNoItems_Edit.Text.Trim();
                if (vSheetNoItems != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    TextBox eConsNo_Edit = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_Edit");
                    TextBox ePrice_Edit = (TextBox)fvIASheetB_Detail.FindControl("ePrice_Edit");
                    TextBox eQuantity_Edit = (TextBox)fvIASheetB_Detail.FindControl("eQuantity_Edit");
                    Label eAmount_Edit = (Label)fvIASheetB_Detail.FindControl("eAmountB_Edit");
                    TextBox eRemarkB_Edit = (TextBox)fvIASheetB_Detail.FindControl("eRemarkB_Edit");
                    //Label eUnit_Edit = (Label)fvIASheetB_Detail.FindControl("eUnit_Edit");
                    Label eSheetNoB_Edit = (Label)fvIASheetB_Detail.FindControl("eSheetNoB_Edit");
                    try
                    {
                        dsIASheetB_Detail.UpdateParameters.Clear();
                        //dsIASheetB_Detail.UpdateCommand = "update IASheetB set ConsNo = @ConsNo, Price = @Price, Quantity = @Quantity, Amount = @Amount, " + Environment.NewLine +
                        //                                  "       Unit = @Unit, RemarkB = @RemarkB, ModifyMAn = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                        //                                  " where SheetNoItems = @SheetNoItems ";
                        dsIASheetB_Detail.UpdateCommand = "update IASheetB set ConsNo = @ConsNo, Price = @Price, Quantity = @Quantity, Amount = @Amount, " + Environment.NewLine +
                                                          "       RemarkB = @RemarkB, ModifyMAn = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                          " where SheetNoItems = @SheetNoItems ";
                        dsIASheetB_Detail.UpdateParameters.Add(new Parameter("ConsNo", DbType.String, (eConsNo_Edit.Text.Trim() != "") ? eConsNo_Edit.Text.Trim() : String.Empty));
                        dsIASheetB_Detail.UpdateParameters.Add(new Parameter("Price", DbType.Int32, (ePrice_Edit.Text.Trim() != "") ? ePrice_Edit.Text.Trim() : "0"));
                        dsIASheetB_Detail.UpdateParameters.Add(new Parameter("Quantity", DbType.Int32, (eQuantity_Edit.Text.Trim() != "") ? eQuantity_Edit.Text.Trim() : "0"));
                        dsIASheetB_Detail.UpdateParameters.Add(new Parameter("Amount", DbType.Int32, (eAmount_Edit.Text.Trim() != "") ? eAmount_Edit.Text.Trim() : "0"));
                        //dsIASheetB_Detail.UpdateParameters.Add(new Parameter("Unit", DbType.String, (eUnit_Edit.Text.Trim() != "") ? eUnit_Edit.Text.Trim() : String.Empty));
                        dsIASheetB_Detail.UpdateParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB_Edit.Text.Trim() != "") ? eRemarkB_Edit.Text.Trim() : String.Empty));
                        dsIASheetB_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                        dsIASheetB_Detail.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                        dsIASheetB_Detail.Update();
                        CalSheetA_Tax(eSheetNoB_Edit.Text.Trim());
                        /* 2022.07.14 計算數量交給資料庫的 Trigger 做
                        if (eConsNo_Edit.Text.Trim() != "")
                        {
                            PF.CalIAStoreQuantity(eConsNo_Edit.Text.Trim(), vConnStr, "");
                        } //*/
                        gridIASheetB_List.DataBind();
                        fvIASheetB_Detail.DataBind();
                        fvIASheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    }
                    catch (Exception eMessage)
                    {
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('" + eMessage.Message + "')");
                        Response.Write("</" + "Script>");
                    }
                }
            }
        }

        protected void eConsNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eConsNo_Edit = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_Edit");
            if (eConsNo_Edit != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eConsName_Edit = (Label)fvIASheetB_Detail.FindControl("eConsName_Edit");
                string vConsNo_Temp = eConsNo_Edit.Text.Trim();
                string vSQLStr_Temp = "select ConsName from IAConsumables where ConsNo = '" + vConsNo_Temp.Trim() + "' ";
                string vConsName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "ConsName");
                if (vConsName_Temp.Trim() == "")
                {
                    vConsName_Temp = vConsNo_Temp.Trim();
                    vSQLStr_Temp = "select top 1 ConsNo from IAConsumables where COnsName like '" + vConsName_Temp.Trim() + "%' order by ConsNo";
                    vConsNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "ConsNo");
                }
                if (vConsNo_Temp.Trim() != "")
                {
                    vSQLStr_Temp = "select (unit + ',' + (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Unit' and ClassNo = a.Unit)) as UnitData " + Environment.NewLine +
                                   "  from IAConsumables as a " + Environment.NewLine +
                                   " where ConsNo = '" + vConsNo_Temp.Trim() + "' ";
                    string vTempStr = PF.GetValue(vConnStr, vSQLStr_Temp, "UnitData");
                    string[] vaTempStr = vTempStr.Split(',');
                    Label eUnit_Edit = (Label)fvIASheetB_Detail.FindControl("eUnit_Edit");
                    Label eUnit_C_Edit = (Label)fvIASheetB_Detail.FindControl("eUnit_C_Edit");
                    eUnit_Edit.Text = vaTempStr[0].Trim();
                    eUnit_C_Edit.Text = vaTempStr[1].Trim();
                }
                eConsNo_Edit.Text = vConsNo_Temp.Trim();
                eConsName_Edit.Text = vConsName_Temp.Trim();
            }
        }

        protected void ePrice_Edit_TextChanged(object sender, EventArgs e)
        {
            Label eAmountB_Edit = (Label)fvIASheetB_Detail.FindControl("eAmountB_Edit");
            if (eAmountB_Edit != null)
            {
                TextBox ePrice_Edit = (TextBox)fvIASheetB_Detail.FindControl("ePrice_Edit");
                TextBox eQuantity_Edit = (TextBox)fvIASheetB_Detail.FindControl("eQuantity_Edit");
                double vPrice_Edit = 0.0;
                if (!double.TryParse(ePrice_Edit.Text.Trim(), out vPrice_Edit))
                {
                    vPrice_Edit = 0.0;
                }
                double vQuantity_Edit = 0.0;
                if (!double.TryParse(eQuantity_Edit.Text.Trim(), out vQuantity_Edit))
                {
                    vQuantity_Edit = 0.0;
                }
                int vAmount = (int)Math.Round(vPrice_Edit * vQuantity_Edit, 0, MidpointRounding.AwayFromZero);
                eAmountB_Edit.Text = vAmount.ToString();
            }
        }

        /// <summary>
        /// 明細新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKB_INS_Click(object sender, EventArgs e)
        {
            string vSQLStr_Temp = "";
            Label eSheetNoItems_INS = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_INS");
            if (eSheetNoItems_INS != null)
            {
                Label eSheetNoB_INS = (Label)fvIASheetB_Detail.FindControl("eSheetNoB_INS");

                string vSheetNo = eSheetNoB_INS.Text.Trim();
                if (vSheetNo != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    TextBox eConsNo_INS = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_INS");
                    TextBox ePrice_INS = (TextBox)fvIASheetB_Detail.FindControl("ePrice_INS");
                    TextBox eQuantity_INS = (TextBox)fvIASheetB_Detail.FindControl("eQuantity_INS");
                    Label eAmountB_INS = (Label)fvIASheetB_Detail.FindControl("eAmountB_INS");
                    TextBox eRemarkB_INS = (TextBox)fvIASheetB_Detail.FindControl("eRemarkB_INS");

                    string vItems = "";
                    string vOldItems = "";
                    int vNewIndex = 0;
                    string vSheetNoItems = "";
                    vSQLStr_Temp = "select max(Items) MaxItems from IASheetB where SheetNo = '" + vSheetNo + "' ";
                    vOldItems = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems").Trim();
                    if (Int32.TryParse(vOldItems, out vNewIndex))
                    {
                        vNewIndex++;
                    }
                    else
                    {
                        vNewIndex = 1;
                    }
                    vItems = vNewIndex.ToString("D4");
                    vSheetNoItems = vSheetNo.Trim() + vItems.Trim();
                    if (vSheetNoItems.Trim() != "")
                    {
                        dsIASheetB_Detail.InsertCommand = "insert into IASheetB (SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, QtyMode, RemarkB, ItemStatus, BuMan, BuDate) " + Environment.NewLine +
                                                          "values (@SheetNoItems, @SheetNo, @Items, @ConsNo, @Price, @Quantity, @Amount, 1, @RemarkB, '000', @BuMan, GetDate()) ";
                        dsIASheetB_Detail.InsertParameters.Clear();
                        dsIASheetB_Detail.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems.Trim()));
                        dsIASheetB_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo.Trim()));
                        dsIASheetB_Detail.InsertParameters.Add(new Parameter("Items", DbType.String, vItems.Trim()));
                        dsIASheetB_Detail.InsertParameters.Add(new Parameter("ConsNo", DbType.String, (eConsNo_INS.Text.Trim() != "") ? eConsNo_INS.Text.Trim() : String.Empty));
                        dsIASheetB_Detail.InsertParameters.Add(new Parameter("Price", DbType.Double, (ePrice_INS.Text.Trim() != "") ? ePrice_INS.Text.Trim() : String.Empty));
                        dsIASheetB_Detail.InsertParameters.Add(new Parameter("Quantity", DbType.Double, (eQuantity_INS.Text.Trim() != "") ? eQuantity_INS.Text.Trim() : String.Empty));
                        dsIASheetB_Detail.InsertParameters.Add(new Parameter("Amount", DbType.Int32, (eAmountB_INS.Text.Trim() != "") ? eAmountB_INS.Text.Trim() : String.Empty));
                        dsIASheetB_Detail.InsertParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB_INS.Text.Trim() != "") ? eRemarkB_INS.Text.Trim() : String.Empty));
                        dsIASheetB_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                        try
                        {
                            dsIASheetB_Detail.Insert();
                            vSQLStr_Temp = "update IASheetA set SheetStatus = '000' where SheetNo = '" + vSheetNo + "' and isnull(SheetStatus, '') = '' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            vSQLStr_Temp = "update IAConsumables set IsInorder = 1 where ConsNo = '" + eConsNo_INS.Text.Trim() + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            CalSheetA_Tax(vSheetNo);
                            // 2022.07.14 計算數量交給資料庫的 Trigger 做
                            //PF.CalIAStoreQuantity(eConsNo_INS.Text.Trim(), vConnStr, "");
                            gridIASheetB_List.DataBind();
                            fvIASheetB_Detail.DataBind();
                            fvIASheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
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
            }
        }

        protected void eConsNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eConsNo_INS = (TextBox)fvIASheetB_Detail.FindControl("eConsNo_INS");
            if (eConsNo_INS != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eConsName_INS = (Label)fvIASheetB_Detail.FindControl("eConsName_INS");
                string vConsNo_Temp = eConsNo_INS.Text.Trim();
                string vSQLStr_Temp = "select ConsName from IAConsumables where ConsNo = '" + vConsNo_Temp.Trim() + "' ";
                string vConsName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "ConsName");
                if (vConsName_Temp.Trim() == "")
                {
                    vConsName_Temp = vConsNo_Temp.Trim();
                    vSQLStr_Temp = "select top 1 ConsNo from IAConsumables where COnsName like '" + vConsName_Temp.Trim() + "%' order by ConsNo";
                    vConsNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "ConsNo");
                }
                if (vConsNo_Temp.Trim() != "")
                {
                    vSQLStr_Temp = "select (unit + ',' + (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Unit' and ClassNo = a.Unit)) as UnitData " + Environment.NewLine +
                                   "  from IAConsumables as a " + Environment.NewLine +
                                   " where ConsNo = '" + vConsNo_Temp.Trim() + "' ";
                    string vTempStr = PF.GetValue(vConnStr, vSQLStr_Temp, "UnitData");
                    string[] vaTempStr = vTempStr.Split(',');
                    Label eUnit_INS = (Label)fvIASheetB_Detail.FindControl("eUnit_INS");
                    Label eUnit_C_INS = (Label)fvIASheetB_Detail.FindControl("eUnit_C_INS");
                    eUnit_INS.Text = vaTempStr[0].Trim();
                    eUnit_C_INS.Text = vaTempStr[1].Trim();
                }
                eConsNo_INS.Text = vConsNo_Temp.Trim();
                eConsName_INS.Text = vConsName_Temp.Trim();
            }
        }

        protected void ePrice_INS_TextChanged(object sender, EventArgs e)
        {
            Label eAmountB_INS = (Label)fvIASheetB_Detail.FindControl("eAmountB_INS");
            if (eAmountB_INS != null)
            {
                TextBox ePrice_INS = (TextBox)fvIASheetB_Detail.FindControl("ePrice_INS");
                TextBox eQuantity_INS = (TextBox)fvIASheetB_Detail.FindControl("eQuantity_INS");
                double vPrice_INS = 0.0;
                if (!double.TryParse(ePrice_INS.Text.Trim(), out vPrice_INS))
                {
                    vPrice_INS = 0.0;
                }
                double vQuantity_INS = 0.0;
                if (!double.TryParse(eQuantity_INS.Text.Trim(), out vQuantity_INS))
                {
                    vQuantity_INS = 0.0;
                }
                int vAmount = (int)Math.Round(vPrice_INS * vQuantity_INS, 0, MidpointRounding.AwayFromZero);
                eAmountB_INS.Text = vAmount.ToString();
            }
        }

        /// <summary>
        /// 移除耗材 "進貨中" 旗標
        /// </summary>
        /// <param name="fConsNo">耗材料號</param>
        private void SetConsumablesStatus(string fConsNo)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Int32 vRCount = 0;
            string vSQLStr_Temp = "select count(SheetNoItems) RCount from IASheetB where ConsNo = '" + fConsNo + "' and isnull(ItemStatus, '000') = '000' ";
            if ((!Int32.TryParse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount"), out vRCount)) || vRCount == 0)
            {
                vSQLStr_Temp = "update IAConsumables set IsInorder = 0 where ConsNo = '" + fConsNo + "' ";
                PF.ExecSQL(vConnStr, vSQLStr_Temp);
            }
        }

        /// <summary>
        /// 明細刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDeleteB_List_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems_List != null)
            {
                Label eSheetNoB_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoB_List");
                Label eConsNo_List = (Label)fvIASheetB_Detail.FindControl("eConsNo_List");
                string vSheetNoItems = eSheetNoItems_List.Text.Trim();
                string vSQLStr_Temp = "delete IASheetB where SheetNoItems = @SheetNoItems";
                string vConsNo_Temp = eConsNo_List.Text.Trim();
                dsIASheetB_Detail.DeleteCommand = vSQLStr_Temp;
                dsIASheetB_Detail.DeleteParameters.Clear();
                dsIASheetB_Detail.DeleteParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                try
                {
                    dsIASheetB_Detail.Delete();
                    // 2022.07.14 計算數量交給資料庫的 Trigger 做
                    //PF.CalIAStoreQuantity(vConsNo_Temp, vConnStr, "");
                    SetConsumablesStatus(vConsNo_Temp);
                    CalSheetA_Tax(eSheetNoB_List.Text.Trim());
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
        /// 明細到貨處理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbArriveB_List_Click(object sender, EventArgs e)
        {
            Label eSheetNoItems_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems_List != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eSheetNoB_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoB_List");
                Label eConsNo_List = (Label)fvIASheetB_Detail.FindControl("eConsNo_List");
                string vSheetNoItems = eSheetNoItems_List.Text.Trim();
                string vSheetNo_Temp = eSheetNoB_List.Text.Trim();
                string vConsNo_Temp = eConsNo_List.Text.Trim();
                //2023.01.12 新增
                Label eQuantity_List = (Label)fvIASheetB_Detail.FindControl("eQuantity_List");
                string vLastQtyStr = (Int32.TryParse(eQuantity_List.Text.Trim(), out int vLastQty)) ? vLastQty.ToString() : "0";
                string vSQLStr_Temp = "select Remark from IAConsumables where ConsNo = '" + vConsNo_Temp + "' ";
                string vConsRemark = PF.GetValue(vConnStr, vSQLStr_Temp, "Remark");
                string vRemarkB_List = (vConsRemark.Trim() != "") ? DateTime.Today.ToShortDateString() + "進貨作業" + Environment.NewLine + vConsRemark.Trim() : DateTime.Today.ToShortDateString() + "進貨作業";
                //=========================================================================================
                vSQLStr_Temp = "update IASheetB set ItemStatus = '001' where SheetNoItems = '" + vSheetNoItems + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    vSQLStr_Temp = "select Count(SheetNoItems) as RCount from IASheetB where SheetNo = '" + vSheetNo_Temp.Trim() + "' and isnull(ItemStatus, '') not in ('001', '999') ";
                    int vNotArriveCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount"));
                    if (vNotArriveCount == 0)
                    {
                        vSQLStr_Temp = "update IASheetA set SheetStatus = '010' where SheetNo = '" + vSheetNo_Temp + "' ";
                    }
                    else
                    {
                        vSQLStr_Temp = "update IASheetA set SheetStatus = '020' where SheetNo = '" + vSheetNo_Temp + "' ";
                    }
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //2023.01.12 寫入最後進庫資料
                    vSQLStr_Temp = "update IAConsumables " + Environment.NewLine +
                                   "   set LastIndate = GetDate(), LastInQty = " + vLastQtyStr + ", " + Environment.NewLine +
                                   "       ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "', " + Environment.NewLine +
                                   "       Remark = '" + vRemarkB_List + "' " + Environment.NewLine +
                                   " where ConsNo = '" + vConsNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    // 2022.07.14 計算數量交給資料庫的 Trigger 做
                    //PF.CalIAStoreQuantity(vConsNo_Temp, vConnStr, "");
                    SetConsumablesStatus(vConsNo_Temp);
                    gridIASheetList.DataBind();
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
        /// 明細作廢處理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbIsAbortB_List_Click(object sender, EventArgs e)
        {
            Label eSheetNoItems_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems_List != null)
            {
                Label eSheetNoB_List = (Label)fvIASheetB_Detail.FindControl("eSheetNoB_List");
                Label eConsNo_List = (Label)fvIASheetB_Detail.FindControl("eConsNo_List");
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSheetNoItems = eSheetNoItems_List.Text.Trim();
                string vConsNo_Temp = eConsNo_List.Text.Trim();
                string vSQLStr_Temp = "update IASheetB set ItemStatus = '999' where SheetNoItems = '" + vSheetNoItems + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    CalSheetA_Tax(eSheetNoB_List.Text.Trim());
                    // 2022.07.14 計算數量交給資料庫的 Trigger 做
                    //PF.CalIAStoreQuantity(vConsNo_Temp, vConnStr, "");
                    SetConsumablesStatus(vConsNo_Temp);
                    gridIASheetList.DataBind();
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
        /// 計算表頭稅額
        /// </summary>
        /// <param name="fSheetNo"></param>
        private void CalSheetA_Tax(string fSheetNo)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            int vAmount = 0;
            int vTaxAMT = 0;
            int vTAmount = 0;

            string vSQLStr_Temp = "select count(SheetNoItems) RCount from IASheetB where SheetNo = '" + fSheetNo + "' and isnull(ItemStatus, '999') <> '999' ";
            int vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount"));
            if (vRCount > 0)
            {
                string vTaxType = "";
                double vTaxRate = 0.0;

                vSQLStr_Temp = "select sum(Amount) TAmount from IASheetB where SheetNo = '" + fSheetNo + "' and isnull(ItemStatus, '999') <> '999' ";
                vTAmount = Int32.Parse(PF.GetValue(vConnStr, vSQLStr_Temp, "TAmount"));
                vSQLStr_Temp = "select TaxRate from IASheetA where SheetNo = '" + fSheetNo + "' ";
                if (!double.TryParse(PF.GetValue(vConnStr, vSQLStr_Temp, "TaxRate"), out vTaxRate))
                {
                    vTaxRate = 5;
                }
                vTaxRate = vTaxRate / 100.0;
                vSQLStr_Temp = "select TaxType from IASheetA where SheetNo = '" + fSheetNo + "' ";
                vTaxType = PF.GetValue(vConnStr, vSQLStr_Temp, "TaxType");
                switch (vTaxType)
                {
                    case "1": //應稅
                        vAmount = vTAmount;
                        vTaxAMT = (int)Math.Round((double)vAmount * vTaxRate, MidpointRounding.AwayFromZero);
                        vTAmount = (vAmount + vTaxAMT);
                        break;
                    case "2": //應稅內含
                        vTaxAMT = (int)Math.Round((double)vTAmount / (1.0 + vTaxRate), MidpointRounding.AwayFromZero);
                        vAmount = vTAmount - vTaxAMT;
                        break;
                    case "3": //零稅或免稅
                    case "4":
                        vAmount = vTAmount;
                        vTaxAMT = 0;
                        break;
                    default:
                        break;
                }
            }
            using (SqlDataSource dsTemp = new SqlDataSource())
            {
                dsTemp.ConnectionString = vConnStr;
                dsTemp.UpdateCommand = "update IASheetA set Amount = @Amount, TaxAMT = @TaxAMT, TotalAmount = @TotalAmount where SheetNo = @SheetNo";
                dsTemp.UpdateParameters.Clear();
                dsTemp.UpdateParameters.Add(new Parameter("Amount", DbType.Int32, vAmount.ToString()));
                dsTemp.UpdateParameters.Add(new Parameter("TaxAMT", DbType.Int32, vTaxAMT.ToString()));
                dsTemp.UpdateParameters.Add(new Parameter("TotalAmount", DbType.Int32, vTAmount.ToString()));
                dsTemp.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, fSheetNo));
                try
                {
                    dsTemp.Update();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message + "')");
                    Response.Write("</" + "Script>");
                }
            }
        }
    }
}