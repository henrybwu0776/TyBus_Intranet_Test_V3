using Amaterasu_Function;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class TicketStore : System.Web.UI.Page
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
                    if (!IsPostBack)
                    {
                        if ((vLoginDepNo != "") && (vLoginDepNo != "09"))
                        {
                            eDepNo.Text = vLoginDepNo;
                            eDepNo.Enabled = false;
                        }
                        else
                        {
                            eDepNo.Text = "";
                            eDepNo.Enabled = true;
                        }
                    }
                    WarehouseDataBind();
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
            string vWStr_DepNo = (eDepNo.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_EmpNo = (eAdminEmpNo.Text.Trim() != "") ? "   and a.AdminEmpNo = '" + eAdminEmpNo.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_IsInused = (rbSearchMode.SelectedValue == "Z") ?
                                    " where isnull(IsInused,'X') <> '' " + Environment.NewLine :
                                    " where isnull(IsInused, 'X') = '" + rbSearchMode.SelectedValue.Trim() + "' " + Environment.NewLine;
            string vSelStr = "select a.WarehouseNo, a.DepNo, a.DepName, a.AdminEmpNo, a.AdminEmpName, a.WarehousePosition, a.Remark, " + Environment.NewLine +
                             "       a.BuDate, a.BuMan, (select [Name] from Employee where EmpNo = a.BuMan) BuMan_C, " + Environment.NewLine +
                             "       a.ModifyDate, a.ModifyMan, (select [Name] from Employee where EmpNo = a.ModifyMan) ModifyMan_C, " + Environment.NewLine +
                             "       case when isnull(IsInused, 'X') = 'X' then '已停用' else '使用中' end IsInused " + Environment.NewLine +
                             "  from TicketStore a " + Environment.NewLine +
                             vWStr_IsInused +
                             vWStr_DepNo +
                             vWStr_EmpNo +
                             " order by a.WarehouseNo ";
            return vSelStr;
        }

        private void WarehouseDataBind()
        {
            string vSelectStr = GetSelStr();
            sdsTicketStoreList.SelectCommand = vSelectStr;
            gridTicketStoreList.DataBind();
        }

        private void BeginUpdateMode()
        {
            RadioButtonList rbIsInused_Edit = (RadioButtonList)fvTicketStoreDetail.FindControl("rbIsInused_Edit");
            TextBox eIsInused_Edit = (TextBox)fvTicketStoreDetail.FindControl("eIsInused_Edit");
            rbIsInused_Edit.SelectedIndex = (eIsInused_Edit.Text.Trim() == "V") ? 0 : 1;
        }

        private void BeginInsertMode()
        {
            RadioButtonList rbIsInused_Insert = (RadioButtonList)fvTicketStoreDetail.FindControl("rbIsInused_Insert");
            TextBox eIsInused_Insert = (TextBox)fvTicketStoreDetail.FindControl("eIsInused_Insert");
            rbIsInused_Insert.SelectedIndex = 0;
            eIsInused_Insert.Text = rbIsInused_Insert.SelectedValue.Trim();
        }

        protected void eDepNo_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo.Text.Trim();
            string vDepName = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            vDepName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName == "")
            {
                vDepName = vDepNo;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName + "' ";
                vDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo.Text = vDepNo;
            lbDepName.Text = vDepName;
        }

        protected void eAdminEmpNo_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vEmpNo = eAdminEmpNo.Text.Trim();
            string vEmpName = "";
            string vSQLStr = "select [Name] from EMployee where EMpNo = '" + vEmpNo + "' and LeaveDay is null";
            vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName + "' and LeaveDay is null";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eAdminEmpNo.Text = vEmpNo;
            eAdminEmpName.Text = vEmpName;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            WarehouseDataBind();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvTicketStoreDetail_DataBound(object sender, EventArgs e)
        {
            switch (fvTicketStoreDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    BeginUpdateMode();
                    break;
                case FormViewMode.Insert:
                    BeginInsertMode();
                    break;
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Edit = (TextBox)fvTicketStoreDetail.FindControl("eDepNo_Edit");
            TextBox eDepName_Edit = (TextBox)fvTicketStoreDetail.FindControl("eDepName_Edit");
            string vDepNo_Temp = eDepNo_Edit.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Edit.Text = vDepNo_Temp;
            eDepName_Edit.Text = vDepName_Temp;
        }

        protected void eAdminEmpNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eEmpNo = (TextBox)fvTicketStoreDetail.FindControl("eAdminEmpNo_Edit");
            TextBox eEmpName = (TextBox)fvTicketStoreDetail.FindControl("eAdminEmpName_Edit");
            string vEmpNo_Temp = eEmpNo.Text.Trim();
            string vEmpName_Temp = "";
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' and LeaveDay is null";
            vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName_Temp + "' and LeaveDay is null ";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eEmpNo.Text = vEmpNo_Temp;
            eEmpName.Text = (vEmpNo_Temp != "") ? vEmpName_Temp : "";
        }

        protected void rbIsInused_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            RadioButtonList rbIsInused_Edit = (RadioButtonList)fvTicketStoreDetail.FindControl("rbIsInused_Edit");
            TextBox eIsInused_Edit = (TextBox)fvTicketStoreDetail.FindControl("eIsInused_Edit");
            eIsInused_Edit.Text = rbIsInused_Edit.SelectedValue.Trim();
        }

        protected void eDepNo_Insert_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Insert = (TextBox)fvTicketStoreDetail.FindControl("eDepNo_Insert");
            TextBox eDepName_Insert = (TextBox)fvTicketStoreDetail.FindControl("eDepName_Insert");
            string vDepNo_Temp = eDepNo_Insert.Text.Trim();
            string vDepName_Temp = "";
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select DepNo from Department where [Name] = '" + vDepName_Temp + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Insert.Text = vDepNo_Temp;
            eDepName_Insert.Text = vDepName_Temp;
        }

        protected void eAdminEmpNo_Insert_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eEmpNo = (TextBox)fvTicketStoreDetail.FindControl("eAdminEmpNo_Insert");
            TextBox eEmpName = (TextBox)fvTicketStoreDetail.FindControl("eAdminEmpName_Insert");
            string vEmpNo_Temp = eEmpNo.Text.Trim();
            string vEmpName_Temp = "";
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' and LeaveDay is null";
            vEmpName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName_Temp == "")
            {
                vEmpName_Temp = vEmpNo_Temp;
                vSQLStr = "select EmpNo from Employee where [Name] = '" + vEmpName_Temp + "' and LeaveDay is null ";
                vEmpNo_Temp = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eEmpNo.Text = vEmpNo_Temp;
            eEmpName.Text = (vEmpNo_Temp != "") ? vEmpName_Temp : "";
        }

        protected void rbIsInused_Insert_SelectedIndexChanged(object sender, EventArgs e)
        {
            RadioButtonList rbIsInused_Insert = (RadioButtonList)fvTicketStoreDetail.FindControl("rbIsInused_Insert");
            TextBox eIsInused_Insert = (TextBox)fvTicketStoreDetail.FindControl("eIsInused_Insert");
            eIsInused_Insert.Text = rbIsInused_Insert.SelectedValue.Trim();
        }

        protected void sdsTicketStoreDetail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                WarehouseDataBind();
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    string vSQLStr = "select WarehouseNo from TicketStore where IsInused = 'V' order by WarehouseNo";
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    if (drTemp.HasRows)
                    {
                        string vWarehouseNoList = "";
                        while (drTemp.Read())
                        {
                            vWarehouseNoList = (vWarehouseNoList == "") ? drTemp["WarehouseNo"].ToString().Trim() : vWarehouseNoList + "," + drTemp["WarehouseNo"].ToString().Trim();
                        }
                        vSQLStr = "update SysFlag set Content = '" + vWarehouseNoList + "' where FormName = 'unSysflag' and ControlItme = 'a_shkNposbset' ";
                        PF.ExecSQL(vConnStr, vSQLStr);
                    }
                }
            }
        }

        protected void sdsTicketStoreDetail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                TextBox eWarehouseNo_Insert = (TextBox)fvTicketStoreDetail.FindControl("eWarehouseNo_Insert");
                TextBox eAdminEmpNo_Insert = (TextBox)fvTicketStoreDetail.FindControl("eAdminEmpNo_Insert");
                string vSQLStr = "update Employee set SellPosB = '" + eWarehouseNo_Insert.Text.Trim() + "' where EmpNo = '" + eAdminEmpNo_Insert.Text.Trim() + "' ";
                PF.ExecSQL(vConnStr, vSQLStr);
                WarehouseDataBind();
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    vSQLStr = "select WarehouseNo from TicketStore where IsInused = 'V' order by WarehouseNo";
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    if (drTemp.HasRows)
                    {
                        string vWarehouseNoList = "";
                        while (drTemp.Read())
                        {
                            vWarehouseNoList = (vWarehouseNoList == "") ? drTemp["WarehouseNo"].ToString().Trim() : vWarehouseNoList + "," + drTemp["WarehouseNo"].ToString().Trim();
                        }
                        vSQLStr = "update SysFlag set Content = '" + vWarehouseNoList + "' where FormName = 'unSysflag' and ControlItme = 'a_shkNposbset' ";
                        PF.ExecSQL(vConnStr, vSQLStr);
                    }
                }
            }
        }

        protected void sdsTicketStoreDetail_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            TextBox eIsInused_Insert = (TextBox)fvTicketStoreDetail.FindControl("eIsInused_Insert");
            e.Command.Parameters["@IsInused"].Value = eIsInused_Insert.Text.Trim();
            e.Command.Parameters["@BuDate"].Value = DateTime.Today;
            e.Command.Parameters["@BuMan"].Value = vLoginID;
        }

        protected void sdsTicketStoreDetail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eWarehouseNo_Edit = (Label)fvTicketStoreDetail.FindControl("eWarehouseNo_Edit");
                TextBox eAdminEmpNo_Edit = (TextBox)fvTicketStoreDetail.FindControl("eAdminEmpNo_Edit");
                string vSQLStr = "update Employee set SellPosB is null where SellPosB = '" + eWarehouseNo_Edit.Text.Trim() + "' ";
                PF.ExecSQL(vConnStr, vSQLStr);
                vSQLStr = "update Employee set SellPosB = '" + eWarehouseNo_Edit.Text.Trim() + "' where EmpNo = '" + eAdminEmpNo_Edit.Text.Trim() + "' ";
                PF.ExecSQL(vConnStr, vSQLStr);
                WarehouseDataBind();
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    vSQLStr = "select WarehouseNo from TicketStore where IsInused = 'V' order by WarehouseNo";
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    if (drTemp.HasRows)
                    {
                        string vWarehouseNoList = "";
                        while (drTemp.Read())
                        {
                            vWarehouseNoList = (vWarehouseNoList == "") ? drTemp["WarehouseNo"].ToString().Trim() : vWarehouseNoList + "," + drTemp["WarehouseNo"].ToString().Trim();
                        }
                        vSQLStr = "update SysFlag set Content = '" + vWarehouseNoList + "' where FormName = 'unSysflag' and ControlItme = 'a_shkNposbset' ";
                        PF.ExecSQL(vConnStr, vSQLStr);
                    }
                }
            }
        }

        protected void sdsTicketStoreDetail_Updating(object sender, SqlDataSourceCommandEventArgs e)
        {
            TextBox eIsInused_Edit = (TextBox)fvTicketStoreDetail.FindControl("eIsInused_Edit");
            e.Command.Parameters["@IsInused"].Value = eIsInused_Edit.Text.Trim();
            e.Command.Parameters["@ModifyDate"].Value = DateTime.Today;
            e.Command.Parameters["@ModifyMan"].Value = vLoginID;
        }
    }
}