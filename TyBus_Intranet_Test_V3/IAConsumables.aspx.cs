using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class IAConsumables : Page
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
                    if (!IsPostBack)
                    {
                        plShowData.Visible = true;
                        plReport.Visible = false;
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

        /// <summary>
        /// 取得資料查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSearchStr()
        {
            string vResultStr = "";
            string vWStr_ConsNo = ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and IA.ConsNo between '" + eConsNo_S_Search.Text.Trim() + "' and '" + eConsNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() == "")) ?
                                  "   and IA.ConsNo like '" + eConsNo_S_Search.Text.Trim() + "%' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() == "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and IA.ConsNo like '" + eConsNo_E_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Brand = (eBrand_Search.SelectedValue.Trim() != "") ? "   and IA.Brand = '" + eBrand_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CorrespondModel = (eCorrespondModel_Search.Text.Trim() != "") ? "   and IA.CorrespondModel like '%" + eCorrespondModel_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Systematics = (eSystematics_Search.SelectedValue.Trim() != "") ? "   and IA.Systematics = '" + eSystematics_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "   and IA.ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_ConsStatus = (eConsStatus_Search.SelectedValue == "S01") ? "   and IA.Quantity > IA.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S02") ? "   and IA.Quantity < IA.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S03") ? "   and IA.Quantity = 0 " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S04") ? "   and IA.IsInorder = 1 " + Environment.NewLine : "";

            vResultStr = "SELECT IA.ConsNo, (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Systematics' and ClassNo = IA.Systematics) as Systematics_C, " + Environment.NewLine +
                         "       (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Brand' and ClassNo = IA.Brand) as Brand_C, " + Environment.NewLine +
                         "       IA.ConsName, IA.OriModelNo, IA.Quantity, " + Environment.NewLine +
                         "       (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Unit' and ClassNo = IA.Unit) as Unit_C, " + Environment.NewLine +
                         "       IA.SaveQty, IA.Position, IA.Spec_Color, IA.Spec_Other, IA.CorrespondModel, IA.LastIndate, IA.LastInQty, IA.IsInorder, IA.Remark, " + Environment.NewLine +
                         "       IA.BuMan, (select [Name] from Employee where EmpNo = IA.BuMan) as BuManName, IA.BuDate, " + Environment.NewLine +
                         "       IA.ModifyMan, (select [Name] from Employee where EmpNo = IA.ModifyMan) as ModifyManName, IA.ModifyDate, IA.Price " + Environment.NewLine +
                         "  FROM IAConsumables as IA  " + Environment.NewLine +
                         " WHERE isnull(IA.ConsNo, '') <> '' " + Environment.NewLine +
                         vWStr_ConsNo +
                         vWStr_Brand +
                         vWStr_CorrespondModel +
                         vWStr_Systematics +
                         vWStr_ConsName +
                         vWStr_ConsStatus +
                         " Order by ConsNo DESC";
            return vResultStr;
        }

        /// <summary>
        /// 取回列表用的查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSearchStr_Print()
        {
            string vResultStr = "";
            string vWStr_ConsNo = ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and IA.ConsNo between '" + eConsNo_S_Search.Text.Trim() + "' and '" + eConsNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() != "") && (eConsNo_E_Search.Text.Trim() == "")) ?
                                  "   and IA.ConsNo like '" + eConsNo_S_Search.Text.Trim() + "%' " + Environment.NewLine :
                                  ((eConsNo_S_Search.Text.Trim() == "") && (eConsNo_E_Search.Text.Trim() != "")) ?
                                  "   and IA.ConsNo like '" + eConsNo_E_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Brand = (eBrand_Search.SelectedValue.Trim() != "") ? "   and IA.Brand = '" + eBrand_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_CorrespondModel = (eCorrespondModel_Search.Text.Trim() != "") ? "   and IA.CorrespondModel like '%" + eCorrespondModel_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_Systematics = (eSystematics_Search.SelectedValue.Trim() != "") ? "   and IA.Systematics = '" + eSystematics_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "   and IA.ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_ConsStatus = (eConsStatus_Search.SelectedValue == "S01") ? "   and IA.Quantity > IA.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S02") ? "   and IA.Quantity < IA.SaveQty " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S03") ? "   and IA.Quantity = 0 " + Environment.NewLine :
                                      (eConsStatus_Search.SelectedValue == "S04") ? "   and IA.IsInorder = 1 " + Environment.NewLine : "";
            vResultStr = "SELECT IA.ConsNo, (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Systematics' and ClassNo = IA.Systematics) as Systematics_C, " + Environment.NewLine +
                         "       (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Brand' and ClassNo = IA.Brand) as Brand_C, " + Environment.NewLine +
                         "       IA.ConsName, IA.OriModelNo, IA.Quantity, " + Environment.NewLine +
                         "       (select ClassTxt from DBDICB where FKey = '電腦課耗材管理  fmIAConsumables Unit' and ClassNo = IA.Unit) as Unit_C, " + Environment.NewLine +
                         "       IA.SaveQty, IA.Position, IA.Spec_Color, IA.Spec_Other, IA.CorrespondModel, IA.LastIndate, IA.LastInQty, " + Environment.NewLine +
                         "       case when isnull(IA.IsInorder, 0) = 1 then 'V' else 'X' end IsInorder, " + Environment.NewLine +
                         "       case when isnull(IA.IsStopUse, 0) = 1 then 'V' else 'X' end IsStopUse, " + Environment.NewLine +
                         "       IA.Remark, IA.BuMan, (select[Name] from Employee where EmpNo = IA.BuMan) as BuManName, IA.BuDate, " + Environment.NewLine +
                         "       IA.ModifyMan, (select[Name] from Employee where EmpNo = IA.ModifyMan) as ModifyManName, IA.ModifyDate, IA.Price " + Environment.NewLine +
                         "  FROM IAConsumables as IA " + Environment.NewLine +
                         " WHERE isnull(IA.ConsNo, '') <> '' " + Environment.NewLine +
                         vWStr_ConsNo +
                         vWStr_Brand +
                         vWStr_CorrespondModel +
                         vWStr_Systematics +
                         vWStr_ConsName +
                         vWStr_ConsStatus +
                         " Order by ConsNo DESC";
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSearchStr();
            dsIAConsumablesList.SelectCommand = vSelectStr;
            gridIAConsumablesList.DataBind();
        }

        protected void bbOK_Search_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbCancel_Search_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvIAConsumablesMamager_DataBound(object sender, EventArgs e)
        {
            switch (fvIAConsumablesMamager.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Label eConsNo_List = (Label)fvIAConsumablesMamager.FindControl("eConsNo_List");
                    if (eConsNo_List != null)
                    {
                        CheckBox cbIsStopuse_List = (CheckBox)fvIAConsumablesMamager.FindControl("cbIsStopuse_List");
                        Label eIsStopuse_List = (Label)fvIAConsumablesMamager.FindControl("eIsStopuse_List");
                        cbIsStopuse_List.Checked = (eIsStopuse_List.Text.Trim().ToUpper() == "TRUE");
                    }
                    break;
                case FormViewMode.Edit:
                    Label eConsNo_Edit = (Label)fvIAConsumablesMamager.FindControl("eConsNo_Edit");
                    if (eConsNo_Edit != null)
                    {
                        CheckBox cbIsStopUse_Edit = (CheckBox)fvIAConsumablesMamager.FindControl("cbIsStopuse_Edit");
                        Label eIsStopuse_Edit = (Label)fvIAConsumablesMamager.FindControl("eIsStopuse_Edit");
                        cbIsStopUse_Edit.Checked = (eIsStopuse_Edit.Text.Trim().ToUpper() == "TRUE");

                        DropDownList ddlUnit_Edit = (DropDownList)fvIAConsumablesMamager.FindControl("ddlUnit_Edit");
                        Label eUnit_Edit = (Label)fvIAConsumablesMamager.FindControl("eUnit_Edit");
                        int vIndex_Edit = 0;
                        for (int i = 0; i < ddlUnit_Edit.Items.Count; i++)
                        {
                            if (ddlUnit_Edit.Items[i].Value.Trim() == eUnit_Edit.Text.Trim())
                            {
                                vIndex_Edit = i;
                            }
                        }
                        ddlUnit_Edit.SelectedIndex = vIndex_Edit;
                    }
                    break;
                case FormViewMode.Insert:
                    Label eConsNo_INS = (Label)fvIAConsumablesMamager.FindControl("eConsNo_INS");
                    if (eConsNo_INS != null)
                    {
                        DropDownList ddlUnit_INS = (DropDownList)fvIAConsumablesMamager.FindControl("ddlUnit_INS");
                        Label eUnit_INS = (Label)fvIAConsumablesMamager.FindControl("eUnit_INS");
                        int vIndex_INS = 0;
                        for (int i = 0; i < ddlUnit_INS.Items.Count; i++)
                        {
                            if (ddlUnit_INS.Items[i].Value.Trim() == eUnit_INS.Text.Trim())
                            {
                                vIndex_INS = i;
                            }
                        }
                        ddlUnit_INS.SelectedIndex = vIndex_INS;
                    }
                    break;
                default:
                    break;
            }
        }

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            Label eConsNo_Edit = (Label)fvIAConsumablesMamager.FindControl("eConsNo_Edit");
            if ((eConsNo_Edit != null) && (eConsNo_Edit.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vConsNo_Temp = eConsNo_Edit.Text.Trim();
                TextBox eConsName_Edit = (TextBox)fvIAConsumablesMamager.FindControl("eConsName_Edit");
                string vConsName_Temp = eConsName_Edit.Text.Trim();
                TextBox eOriModelNo_Edit = (TextBox)fvIAConsumablesMamager.FindControl("eOriModelNo_Edit");
                string vOriModelNo_Temp = eOriModelNo_Edit.Text.Trim();
                TextBox eCorrespondModel_Edit = (TextBox)fvIAConsumablesMamager.FindControl("eCorrespondModel_Edit");
                string vCorrespondModel_Temp = eCorrespondModel_Edit.Text.Trim();
                TextBox eSpec_Color_Edit = (TextBox)fvIAConsumablesMamager.FindControl("eSpec_Color_Edit");
                string vSpec_Color_Temp = eSpec_Color_Edit.Text.Trim();
                TextBox eSpec_Other_Edit = (TextBox)fvIAConsumablesMamager.FindControl("eSpec_Other_Edit");
                string vSpec_Other_Temp = eSpec_Other_Edit.Text.Trim();
                TextBox ePosition_Edit = (TextBox)fvIAConsumablesMamager.FindControl("ePosition_Edit");
                string vPosition_Temp = ePosition_Edit.Text.Trim();
                TextBox eRemark_Edit = (TextBox)fvIAConsumablesMamager.FindControl("eRemark_Edit");
                string vRemark_Temp = eRemark_Edit.Text.Trim();
                Label eIsStopuse_Edit = (Label)fvIAConsumablesMamager.FindControl("eIsStopuse_Edit");
                string vIsStopuse_Temp = eIsStopuse_Edit.Text.Trim();
                Label eUnit_Edit = (Label)fvIAConsumablesMamager.FindControl("eUnit_Edit");
                string vUnit_Temp = eUnit_Edit.Text.Trim();
                //2023.01.11 新增
                TextBox ePrice_Edit = (TextBox)fvIAConsumablesMamager.FindControl("ePrice_Edit");
                string vPrice_Temp = (ePrice_Edit.Text.Trim() != "") ? ePrice_Edit.Text.Trim() : "0";
                //==========================================================================================================
                try
                {
                    string vRecordNote = "更新耗材資料_" + vConsNo_Temp + "==" + vConsName_Temp + Environment.NewLine +
                                         "IAConsumables.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    dsIAConsumablesDetail.UpdateCommand = "UPDATE IAConsumables " + Environment.NewLine +
                                                          "   SET ConsName = @ConsName, OriModelNo = @OriModelNo, CorrespondModel = @CorrespondModel, Unit = @Unit, " + Environment.NewLine +
                                                          "       Spec_Color = @Spec_Color, Spec_Other = @Spec_Other, Position = @Position, Remark = @Remark, " + Environment.NewLine +
                                                          //2023.01.11 修改
                                                          //"       IsStopuse = @IsStopuse, ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                                          "       IsStopuse = @IsStopuse, ModifyMan = @ModifyMan, ModifyDate = GetDate(), Price = @Price " + Environment.NewLine +
                                                          " WHERE (ConsNo = @ConsNo)";
                    dsIAConsumablesDetail.UpdateParameters.Clear();
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("ConsName", DbType.String, vConsName_Temp));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("OriModelNo", DbType.String, vOriModelNo_Temp));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("CorrespondModel", DbType.String, vCorrespondModel_Temp));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("Unit", DbType.String, vUnit_Temp));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("Spec_Color", DbType.String, vSpec_Color_Temp));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("Spec_Other", DbType.String, vSpec_Other_Temp));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("Position", DbType.String, vPosition_Temp));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, vRemark_Temp));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("IsStopuse", DbType.Boolean, vIsStopuse_Temp));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("Price", DbType.Double, vPrice_Temp)); //2023.01.11 新增
                    dsIAConsumablesDetail.UpdateParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo_Temp));
                    dsIAConsumablesDetail.Update();
                    gridIAConsumablesList.DataBind();
                    fvIAConsumablesMamager.DataBind();
                    fvIAConsumablesMamager.ChangeMode(FormViewMode.ReadOnly);
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

        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            Label eConsNo_INS = (Label)fvIAConsumablesMamager.FindControl("eConsNo_INS");
            if (eConsNo_INS != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vConsNo_Old = "";
                string vConsNo_Temp = "";
                TextBox eConsName_INS = (TextBox)fvIAConsumablesMamager.FindControl("eConsName_INS");
                string vConsName_Temp = eConsName_INS.Text.Trim();
                Label eSystematics_INS = (Label)fvIAConsumablesMamager.FindControl("eSystematics_INS");
                string vSystematics_Temp = eSystematics_INS.Text.Trim();
                Label eBrand_INS = (Label)fvIAConsumablesMamager.FindControl("eBrand_INS");
                string vBrand_Temp = eBrand_INS.Text.Trim();
                TextBox eOriModelNo_INS = (TextBox)fvIAConsumablesMamager.FindControl("eOriModelNo_INS");
                string vOriModel_Temp = eOriModelNo_INS.Text.Trim();
                TextBox eCorrespondModel_INS = (TextBox)fvIAConsumablesMamager.FindControl("eCorrespondModel_INS");
                string vCorrespondModel_Temp = eCorrespondModel_INS.Text.Trim();
                Label eUnit_INS = (Label)fvIAConsumablesMamager.FindControl("eUnit_INS");
                string vUnit_Temp = eUnit_INS.Text.Trim();
                TextBox eSpec_Color_INS = (TextBox)fvIAConsumablesMamager.FindControl("eSpec_Color_INS");
                string vSpec_Color_Temp = eSpec_Color_INS.Text.Trim();
                TextBox eSpec_Other_INS = (TextBox)fvIAConsumablesMamager.FindControl("eSpec_Other_INS");
                string vSpec_Other_Temp = eSpec_Other_INS.Text.Trim();
                TextBox ePosition_INS = (TextBox)fvIAConsumablesMamager.FindControl("ePosition_INS");
                string vPosition_Temp = ePosition_INS.Text.Trim();
                TextBox eRemark_INS = (TextBox)fvIAConsumablesMamager.FindControl("eRemark_INS");
                string vRemark_Temp = eRemark_INS.Text.Trim();
                string vBuMan_Temp = vLoginID;
                vConsNo_Old = vSystematics_Temp + "-" + (vBrand_Temp.Trim() + "------").Substring(0, 6);
                int vConsIndex = 0;
                string vLastConsNo = "";
                string vSQLStr = "select MAX(ConsNo) ConsNo_Max from IAConsumables where ConsNo like '" + vConsNo_Old.Trim() + "%' ";
                vLastConsNo = PF.GetValue(vConnStr, vSQLStr, "ConsNo_Max");
                vConsIndex = (vLastConsNo.Trim() != "") ? Int32.Parse(vLastConsNo.Replace(vConsNo_Old.Trim(), "")) + 1 : 1;
                vConsNo_Temp = vConsNo_Old + vConsIndex.ToString("D5");
                //2023.01.11 新增
                TextBox ePrice_INS = (TextBox)fvIAConsumablesMamager.FindControl("ePrice_INS");
                string vPrice_Temp = (ePrice_INS.Text.Trim() != "") ? ePrice_INS.Text.Trim() : "0";
                //==================================================================================================================
                try
                {
                    string vRecordNote = "建立耗材資料_" + vConsNo_Temp + "==" + vConsName_Temp + Environment.NewLine +
                                         "IAConsumables.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //2023.01.11 修改
                    //dsIAConsumablesDetail.InsertCommand = "INSERT INTO IAConsumables(ConsNo, ConsName, Systematics, Brand, OriModelNo, CorrespondModel, Unit, Spec_Color, Spec_Other, Position, Remark, BuMan, BuDate) " + Environment.NewLine +
                    //                                      "VALUES (@ConsNo, @ConsName, @Systematics, @Brand, @OriModelNo, @CorrespondModel, @Unit, @Spec_Color, @Spec_Other, @Position, @Remark, @BuMan, GetDate())";
                    dsIAConsumablesDetail.InsertCommand = "INSERT INTO IAConsumables(ConsNo, ConsName, Systematics, Brand, OriModelNo, CorrespondModel, Unit, Price, Spec_Color, Spec_Other, Position, Remark, BuMan, BuDate) " + Environment.NewLine +
                                                          "VALUES (@ConsNo, @ConsName, @Systematics, @Brand, @OriModelNo, @CorrespondModel, @Unit, @Price, @Spec_Color, @Spec_Other, @Position, @Remark, @BuMan, GetDate())";
                    dsIAConsumablesDetail.InsertParameters.Clear();
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("ConsName", DbType.String, vConsName_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("Systematics", DbType.String, vSystematics_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("Brand", DbType.String, vBrand_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("OriModelNo", DbType.String, vOriModel_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("CorrespondModel", DbType.String, vCorrespondModel_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("Unit", DbType.String, vUnit_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("Price", DbType.Double, vPrice_Temp)); //2023.01.11 新增
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("Spec_Color", DbType.String, vSpec_Color_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("Spec_Other", DbType.String, vSpec_Other_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("Position", DbType.String, vPosition_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, vRemark_Temp));
                    dsIAConsumablesDetail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vBuMan_Temp));
                    dsIAConsumablesDetail.Insert();
                    gridIAConsumablesList.DataBind();
                    fvIAConsumablesMamager.DataBind();
                    fvIAConsumablesMamager.ChangeMode(FormViewMode.ReadOnly);
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

        protected void ddlSystematics_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTemp = (DropDownList)fvIAConsumablesMamager.FindControl("ddlSystematics_INS");
            Label eTemp = (Label)fvIAConsumablesMamager.FindControl("eSystematics_INS");
            eTemp.Text = ddlTemp.SelectedValue.ToString().Trim();
        }

        protected void ddlBrand_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTemp = (DropDownList)fvIAConsumablesMamager.FindControl("ddlBrand_INS");
            Label eTemp = (Label)fvIAConsumablesMamager.FindControl("eBrand_INS");
            eTemp.Text = ddlTemp.SelectedValue.ToString().Trim();
        }

        protected void cbIsStopuse_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbTemp = (CheckBox)fvIAConsumablesMamager.FindControl("cbIsStopuse_Edit");
            Label eTemp = (Label)fvIAConsumablesMamager.FindControl("eIsStopuse_Edit");
            eTemp.Text = (cbTemp.Checked == true) ? "true" : "false";
        }

        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eConsNo_List = (Label)fvIAConsumablesMamager.FindControl("eConsNo_List");
            if (eConsNo_List != null)
            {
                dsIAConsumablesDetail.DeleteCommand = "delete IAConsumables where ConsNo = @ConsNo";
                dsIAConsumablesDetail.DeleteParameters.Clear();
                dsIAConsumablesDetail.DeleteParameters.Add(new Parameter("ConsNo", DbType.String, eConsNo_List.Text.Trim()));
                dsIAConsumablesDetail.Delete();
                gridIAConsumablesList.DataBind();
                fvIAConsumablesMamager.DataBind();
            }
        }

        protected void ddlUnit_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlUnit_INS = (DropDownList)fvIAConsumablesMamager.FindControl("ddlUnit_INS");
            Label eUnit_INS = (Label)fvIAConsumablesMamager.FindControl("eUnit_INS");
            eUnit_INS.Text = ddlUnit_INS.SelectedValue.Trim();
        }

        protected void ddlUnit_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlUnit_Edit = (DropDownList)fvIAConsumablesMamager.FindControl("ddlUnit_Edit");
            Label eUnit_Edit = (Label)fvIAConsumablesMamager.FindControl("eUnit_Edit");
            eUnit_Edit.Text = ddlUnit_Edit.SelectedValue.Trim();
        }

        /// <summary>
        /// 產生盤點單 (不列印，直接出 EXCEL 檔)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbStoreCheckList_Search_Click(object sender, EventArgs e)
        {
            Int32 vTempINT = 0;
            string vSQLStrTemp = "select ConsNo, Brand, ConsName, OriModelNo, isnull(CorrespondModel, '') CorrespondModal, isnull(Spec_Color, '') Spec_Color, " + Environment.NewLine +
                                 "       isnull(Quantity, '') Quantity, isnull(SaveQty, '') SaveQty, cast('' as varchar) as InventoryQty, Position " + Environment.NewLine +
                                 "  from IAConsumables " + Environment.NewLine +
                                 " order by Systematics, ConsNo";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connStoreList = new SqlConnection())
            {
                connStoreList.ConnectionString = vConnStr;
                SqlCommand cmdStoreList = new SqlCommand(vSQLStrTemp, connStoreList);
                connStoreList.Open();
                SqlDataReader drStoreList = cmdStoreList.ExecuteReader();
                if (drStoreList.HasRows)
                {
                    //有資料才做下面的動作
                    HSSFWorkbook wbExcel = new HSSFWorkbook();
                    //Excel 工作表
                    HSSFSheet wsExcel;
                    //設定標題欄位的格式
                    HSSFCellStyle csTitle = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csTitle.Alignment = HorizontalAlignment.Center; //水平置中

                    //設定字體格式
                    HSSFFont fontTitle = (HSSFFont)wbExcel.CreateFont();
                    //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
                    fontTitle.IsBold = true;
                    fontTitle.FontHeightInPoints = 12; //字體大小
                    csTitle.SetFont(fontTitle);

                    //設定資料內容欄位的格式
                    HSSFCellStyle csData = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                    HSSFFont fontData = (HSSFFont)wbExcel.CreateFont();
                    fontData.FontHeightInPoints = 12;
                    csData.SetFont(fontData);

                    HSSFCellStyle csData_Red = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData_Red.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                    HSSFFont fontData_Red = (HSSFFont)wbExcel.CreateFont();
                    fontData_Red.FontHeightInPoints = 12; //字體大小
                    fontData_Red.Color = HSSFColor.Red.Index; //字體顏色
                    csData_Red.SetFont(fontData_Red);

                    HSSFCellStyle csData_Int = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData_Int.VerticalAlignment = VerticalAlignment.Center; //垂直置中

                    HSSFDataFormat format = (HSSFDataFormat)wbExcel.CreateDataFormat();
                    csData_Int.DataFormat = format.GetFormat("###,##0");

                    HSSFCellStyle csData_Float = (HSSFCellStyle)wbExcel.CreateCellStyle();
                    csData_Float.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                    csData_Float.Alignment = HorizontalAlignment.Right;

                    HSSFDataFormat format_Float = (HSSFDataFormat)wbExcel.CreateDataFormat();
                    csData_Float.DataFormat = format.GetFormat("##0.00");

                    string vFileName = "桃園汽車客運電腦課庫存耗材盤點表";
                    string vHeaderText = "";
                    int vLinesNo = 0;
                    //新增一個工作表
                    wsExcel = (HSSFSheet)wbExcel.CreateSheet(vFileName);
                    //寫入標題列
                    wsExcel.CreateRow(vLinesNo);
                    for (int vCellCount = 0; vCellCount < drStoreList.FieldCount; vCellCount++)
                    {
                        vHeaderText = (drStoreList.GetName(vCellCount).ToUpper() == "CONSNO") ? "庫存編號" :
                                      (drStoreList.GetName(vCellCount).ToUpper() == "BRAND") ? "廠牌" :
                                      (drStoreList.GetName(vCellCount).ToUpper() == "CONSNAME") ? "庫存品項" :
                                      (drStoreList.GetName(vCellCount).ToUpper() == "ORIMODELNO") ? "原廠料號" :
                                      (drStoreList.GetName(vCellCount).ToUpper() == "CORRESPONDMODAL") ? "適用機型" :
                                      (drStoreList.GetName(vCellCount).ToUpper() == "SPEC_COLOR") ? "顏色" :
                                      (drStoreList.GetName(vCellCount).ToUpper() == "QUANTITY") ? "現有庫存量" :
                                      (drStoreList.GetName(vCellCount).ToUpper() == "SAVEQTY") ? "安全庫存量" :
                                      (drStoreList.GetName(vCellCount).ToUpper() == "INVENTORYQTY") ? "實際盤點量" :
                                      (drStoreList.GetName(vCellCount).ToUpper() == "POSITION") ? "庫存位置" :
                                      drStoreList.GetName(vCellCount).Trim();
                        wsExcel.GetRow(vLinesNo).CreateCell(vCellCount).SetCellValue(vHeaderText);
                        wsExcel.GetRow(vLinesNo).GetCell(vCellCount).CellStyle = csTitle;
                    }
                    while (drStoreList.Read())
                    {
                        vLinesNo++;
                        wsExcel.CreateRow(vLinesNo);
                        for (int i = 0; i < drStoreList.FieldCount; i++)
                        {
                            wsExcel.GetRow(vLinesNo).CreateCell(i);
                            vHeaderText = drStoreList.GetName(i).ToUpper();
                            if (drStoreList[i].ToString().Trim() != "")
                            {
                                if ((vHeaderText == "QUANTITY") ||
                                    (vHeaderText == "SAVEQTY") ||
                                    (vHeaderText == "INVENTORYQTY"))
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.Numeric);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(Int32.TryParse(drStoreList[i].ToString(), out vTempINT) ? vTempINT : 0);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData_Int;
                                }
                                else
                                {
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                    wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(drStoreList[i].ToString());
                                    wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                                }
                            }
                            else
                            {
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellType(CellType.String);
                                wsExcel.GetRow(vLinesNo).GetCell(i).SetCellValue(String.Empty);
                                wsExcel.GetRow(vLinesNo).GetCell(i).CellStyle = csData;
                            }
                        }
                    }
                    try
                    {
                        var msTarget = new NPOIMemoryStream();
                        msTarget.AllowClose = false;
                        wbExcel.Write(msTarget);
                        msTarget.Flush();
                        msTarget.Seek(0, SeekOrigin.Begin);
                        msTarget.AllowClose = true;

                        if (msTarget.Length > 0)
                        {
                            string vRecordNote = "匯出檔案_" + vFileName + Environment.NewLine +
                                                 "IAConsumables.aspx";
                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                            //先判斷是不是IE
                            HttpContext.Current.Response.ContentType = "application/octet-stream";
                            HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                            string TourVerision = brObject.Type;
                            if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                            {
                                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                            }
                            else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                            {
                                // 設定強制下載標頭
                                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
                            }
                            // 輸出檔案
                            Response.BinaryWrite(msTarget.ToArray());
                            msTarget.Close();

                            Response.End();
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
        }

        /// <summary>
        /// 匯入盤點量
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbUpdateReserve_Search_Click(object sender, EventArgs e)
        {
            string vSQLStrTemp = "";
            string vConsNo = "";
            Int32 vQuantity = 0;
            Int32 vSaveQty = 0;
            Int32 vInventoryQty = 0;
            Int32 vQtyDift = 0;
            Int32 vQtyMode = 1;
            Int32 vIndex_Count = 0;
            string vItems = "";
            string vRemarkB = "";
            Int32 vTempINT = 0;
            string vSheetNo = "";
            string vSheetNoItems = "";
            string vRemarkA = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (fuExcel.FileName.Trim() != "")
            {
                string vExtName = Path.GetExtension(fuExcel.FileName);
                switch (vExtName)
                {
                    case ".xls": //舊版 EXCEL (97-2003) 格式
                        HSSFWorkbook wbExcel_H = new HSSFWorkbook(fuExcel.FileContent);
                        HSSFSheet sheetExcel_H = (HSSFSheet)wbExcel_H.GetSheetAt(0); //取回第一個工作表
                        if (sheetExcel_H.LastRowNum > 0)
                        {
                            using (SqlDataSource dsTempA = new SqlDataSource())
                            {
                                vSheetNo = PF.GetIASheetNo(vConnStr, "SS");
                                vRemarkA = DateTime.Today.ToShortDateString() + " 盤點作業";
                                dsTempA.ConnectionString = vConnStr;
                                dsTempA.InsertCommand = "insert into IASheetA(SheetNo, SheetMode, BuDate, BuMan, SheetStatus, StatusDate, RemarkA)" + Environment.NewLine +
                                                        "values(@SheetNo, 'SS', GetDate(), @BuMan, '998', GetDate(), @RemarkA)";
                                dsTempA.InsertParameters.Clear();
                                dsTempA.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                dsTempA.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                dsTempA.InsertParameters.Add(new Parameter("RemarkA", DbType.String, vRemarkA));
                                dsTempA.Insert();
                            }
                            for (int i = 0; i <= sheetExcel_H.LastRowNum; i++)
                            {
                                HSSFRow vRowExcel_H = (HSSFRow)sheetExcel_H.GetRow(i);
                                if ((vRowExcel_H != null) && (vRowExcel_H.Cells[0].StringCellValue.Trim() != "庫存編號"))
                                {
                                    vConsNo = vRowExcel_H.Cells[0].StringCellValue.Trim();
                                    vQuantity = Int32.TryParse(vRowExcel_H.Cells[6].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vSaveQty = Int32.TryParse(vRowExcel_H.Cells[7].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vInventoryQty = Int32.TryParse(vRowExcel_H.Cells[8].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vQtyMode = ((vInventoryQty - vQuantity) < 0) ? -1 : 1;
                                    vQtyDift = Math.Abs(vInventoryQty - vQuantity);
                                    vSQLStrTemp = "select MAX(Items) MaxItem from IASheetB where SheetNo = '" + vSheetNo + "' ";
                                    vSQLStrTemp = "select MAX(Items) MaxItem from IASheetB where SheetNo = '" + vSheetNo + "' ";
                                    vIndex_Count = Int32.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "MaxItem"), out vTempINT) ? vTempINT + 1 : 1;
                                    vItems = vIndex_Count.ToString("D4").Trim();
                                    vSheetNoItems = vSheetNo + vItems;
                                    vRemarkB = DateTime.Today.ToShortDateString() + " 盤點數量更新" + Environment.NewLine;
                                    using (SqlDataSource dsTempB = new SqlDataSource())
                                    {
                                        dsTempB.ConnectionString = vConnStr;
                                        dsTempB.InsertCommand = "insert into IASheetB(SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                                                "QtyMode, ItemStatus, BuMan, BuDate)" + Environment.NewLine +
                                                                "values (@SheetNoItems, @SheetNo, @Items, @ConsNo, 0, @Quantity, 0, @RemarkB, " + Environment.NewLine +
                                                                "        @QtyMode, '001', @BuMan, GetDate())";
                                        dsTempB.InsertParameters.Clear();
                                        dsTempB.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                                        dsTempB.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                        dsTempB.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                                        dsTempB.InsertParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo));
                                        dsTempB.InsertParameters.Add(new Parameter("Quantity", DbType.Int32, vQtyDift.ToString()));
                                        dsTempB.InsertParameters.Add(new Parameter("RemarkB", DbType.String, vRemarkB));
                                        dsTempB.InsertParameters.Add(new Parameter("QtyMode", DbType.Int32, vQtyMode.ToString()));
                                        dsTempB.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                        try
                                        {
                                            dsTempB.Insert();
                                            string vRecordNote = "匯入庫存量" + Environment.NewLine +
                                                                 "IAConsumables.aspx";
                                            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                                            // 2022.07.14 計算數量交給資料庫的 Trigger 做
                                            //PF.CalIAStoreQuantity(vConsNo, vConnStr, "");
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
                        else
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('查無資料。')");
                            Response.Write("</" + "Script>");
                        }
                        break;

                    case ".xlsx": //新版 EXCEL (2010 以後) 格式
                        XSSFWorkbook wbExcel_X = new XSSFWorkbook(fuExcel.FileContent);
                        XSSFSheet sheetExcel_X = (XSSFSheet)wbExcel_X.GetSheetAt(0); //取回第一個工作表
                        if (sheetExcel_X.LastRowNum > 0)
                        {
                            using (SqlDataSource dsTempA = new SqlDataSource())
                            {
                                vSheetNo = PF.GetIASheetNo(vConnStr, "SS");
                                vRemarkA = DateTime.Today.ToShortDateString() + " 盤點作業";
                                dsTempA.ConnectionString = vConnStr;
                                dsTempA.InsertCommand = "insert into IASheetA(SheetNo, SheetMode, BuDate, BuMan, SheetStatus, StatusDate, RemarkA)" + Environment.NewLine +
                                                        "values(@SheetNo, 'SS', GetDate(), @BuMan, '998', GetDate(), @RemarkA)";
                                dsTempA.InsertParameters.Clear();
                                dsTempA.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                dsTempA.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                dsTempA.InsertParameters.Add(new Parameter("RemarkA", DbType.String, vRemarkA));
                                dsTempA.Insert();
                            }
                            for (int i = 0; i <= sheetExcel_X.LastRowNum; i++)
                            {
                                XSSFRow vRowExcel_X = (XSSFRow)sheetExcel_X.GetRow(i);
                                if ((vRowExcel_X != null) && (vRowExcel_X.Cells[0].StringCellValue.Trim() != "庫存編號"))
                                {
                                    vConsNo = vRowExcel_X.Cells[0].StringCellValue.Trim();
                                    vQuantity = Int32.TryParse(vRowExcel_X.Cells[6].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vSaveQty = Int32.TryParse(vRowExcel_X.Cells[7].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vInventoryQty = Int32.TryParse(vRowExcel_X.Cells[8].ToString().Trim(), out vTempINT) ? vTempINT : 0;
                                    vQtyMode = ((vInventoryQty - vQuantity) < 0) ? -1 : 1;
                                    vQtyDift = Math.Abs(vInventoryQty - vQuantity);
                                    vSQLStrTemp = "select MAX(Items) MaxItem from IASheetB where SheetNo = '" + vSheetNo + "' ";
                                    vIndex_Count = Int32.TryParse(PF.GetValue(vConnStr, vSQLStrTemp, "MaxItems"), out vTempINT) ? vTempINT + 1 : 1;
                                    vItems = vIndex_Count.ToString("D4").Trim();
                                    vSheetNoItems = vSheetNo + vItems;
                                    vRemarkB = DateTime.Today.ToShortDateString() + " 盤點數量更新" + Environment.NewLine;
                                    using (SqlDataSource dsTempB = new SqlDataSource())
                                    {
                                        dsTempB.ConnectionString = vConnStr;
                                        dsTempB.InsertCommand = "insert into IASheetB(SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                                                "QtyMode, ItemStatus, BuMan, BuDate)" + Environment.NewLine +
                                                                "values (@SheetNoItems, @SheetNo, @Items, @ConsNo, 0, @Quantity, 0, @RemarkB, " + Environment.NewLine +
                                                                "        @QtyMode, '001', @BuMan, GetDate())";
                                        dsTempB.InsertParameters.Clear();
                                        dsTempB.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                                        dsTempB.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                                        dsTempB.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                                        dsTempB.InsertParameters.Add(new Parameter("ConsNo", DbType.String, vConsNo));
                                        dsTempB.InsertParameters.Add(new Parameter("Quantity", DbType.Int32, vQtyDift.ToString()));
                                        dsTempB.InsertParameters.Add(new Parameter("RemarkB", DbType.String, vRemarkB));
                                        dsTempB.InsertParameters.Add(new Parameter("QtyMode", DbType.Int32, vQtyMode.ToString()));
                                        dsTempB.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                                        try
                                        {
                                            dsTempB.Insert();
                                            // 2022.07.14 計算數量交給資料庫的 Trigger 做
                                            //PF.CalIAStoreQuantity(vConsNo, vConnStr, "");
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
                        else
                        {
                            Response.Write("<Script language='Javascript'>");
                            Response.Write("alert('查無資料。')");
                            Response.Write("</" + "Script>");
                        }
                        break;

                    default:
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您指定的檔案不是 EXCEL 檔，請重新確認檔案格式。')");
                        Response.Write("</" + "Script>");
                        break;
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇盤點結果檔')");
                Response.Write("</" + "Script>");
            }
        }

        /// <summary>
        /// 庫存量報表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void IAStoreList_Search_Click(object sender, EventArgs e)
        {
            /* 2023.08.14 把庫存報表移到另一個專門的頁面
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSelectStr = GetSearchStr_Print();
            using (SqlConnection connPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrint = new SqlDataAdapter(vSelectStr, connPrint);
                connPrint.Open();
                DataTable dtPrint = new DataTable("IAConsumablesP");
                daPrint.Fill(dtPrint);
                if (dtPrint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("IAConsumablesP", dtPrint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\IAConsumablesP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "電腦課耗材庫存清單"));
                    rvPrint.LocalReport.Refresh();
                    plReport.Visible = true;
                    plShowData.Visible = false;
                }
            } //*/
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plShowData.Visible = true;
            plReport.Visible = false;
        }
    }
}