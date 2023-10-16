using Amaterasu_Function;
using System;
using System.Data.SqlClient;

namespace TyBus_Intranet_Test_V3
{
    public partial class AlertTimeSetting : System.Web.UI.Page
    {
        PublicFunction PF = new PublicFunction(); //加入公用程式碼參考
        private string vLoginID = "";
        private string vLoginName = "";
        private string vLoginDepNo = "";
        private string vLoginDepName = "";
        private string vLoginTitle = "";
        private string vLoginTitleName = "";
        private string vLoginEmpType = "";
        private string vConnStr = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                vLoginID = (Session["LoginID"] != null) ? Session["LoginID"].ToString().Trim() : "";
                vLoginName = (Session["LoginName"] != null) ? Session["LoginName"].ToString().Trim() : "";
                vLoginDepNo = (Session["LoginDepNo"] != null) ? Session["LoginDepNo"].ToString().Trim() : "";
                vLoginDepName = (Session["LoginDepName"] != null) ? Session["LoginDepName"].ToString().Trim() : "";
                vLoginTitle = (Session["LoginTitle"] != null) ? Session["LoginTitle"].ToString().Trim() : "";
                vLoginTitleName = (Session["LoginTitleName"] != null) ? Session["LoginTitleName"].ToString().Trim() : "";
                vLoginEmpType = (Session["LoginEmpType"] != null) ? Session["LoginEmpType"].ToString().Trim() : "";

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        SetEditMode(false);
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

        /// <summary>
        /// 設定 TextBox 是不是可以輸入
        /// </summary>
        /// <param name="IsEnable"></param>
        private void SetEditMode(bool IsEnable)
        {
            eAlertTime01.Enabled = IsEnable;
            eAlertTime02.Enabled = IsEnable;
            eAlertTime03.Enabled = IsEnable;
            eAlertTime04.Enabled = IsEnable;
            bbCancel.Enabled = IsEnable;
            bbOK.Enabled = IsEnable;
            bbModify.Enabled = !IsEnable;
            bbExit.Enabled = !IsEnable;
        }

        /// <summary>
        /// 回傳查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            string vSelectStr = "select MAX(case when ControlItem = 'a_AlertTime01' then Content else NULL end) as AlertTime01, " + Environment.NewLine +
                                "       MAX(case when ControlItem = 'a_AlertTime02' then Content else NULL end) as AlertTime02, " + Environment.NewLine +
                                "       MAX(case when ControlItem = 'a_AlertTime03' then Content else NULL end) as AlertTime03, " + Environment.NewLine +
                                "       MAX(case when ControlItem = 'a_AlertTime04' then Content else NULL end) as AlertTime04 " + Environment.NewLine +
                                "  from SysFlag " + Environment.NewLine +
                                " where FormName = 'unSysFlag' " + Environment.NewLine +
                                "   and ControlItem like 'a_AlertTime%'";
            return vSelectStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSelectStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelectStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    eAlertTime01.Text = drTemp["AlertTime01"].ToString().Trim();
                    eAlertTime02.Text = drTemp["AlertTime02"].ToString().Trim();
                    eAlertTime03.Text = drTemp["AlertTime03"].ToString().Trim();
                    eAlertTime04.Text = drTemp["AlertTime04"].ToString().Trim();
                }
            }
        }

        protected void bbModify_Click(object sender, EventArgs e)
        {
            SetEditMode(true);
            eAlertTime01.Focus();
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vUpdateStr = (eAlertTime01.Text.Trim() != "") ?
                                "update SysFlag set Content = '" + eAlertTime01.Text.Trim() + "' where FormName = 'unSysFlag' and ControlItem = 'a_AlertTime01' " :
                                "update SysFlag set Content = NULL where FormName = 'unSysFlag' and ControlItem = 'a_AlertTime01' ";
            PF.ExecSQL(vConnStr, vUpdateStr);
            vUpdateStr = (eAlertTime02.Text.Trim() != "") ?
                         "update SysFlag set Content = '" + eAlertTime02.Text.Trim() + "' where FormName = 'unSysFlag' and ControlItem = 'a_AlertTime02' " :
                         "update SysFlag set Content = NULL where FormName = 'unSysFlag' and ControlItem = 'a_AlertTime02' ";
            PF.ExecSQL(vConnStr, vUpdateStr);
            vUpdateStr = (eAlertTime03.Text.Trim() != "") ?
                         "update SysFlag set Content = '" + eAlertTime03.Text.Trim() + "' where FormName = 'unSysFlag' and ControlItem = 'a_AlertTime03' " :
                         "update SysFlag set Content = NULL where FormName = 'unSysFlag' and ControlItem = 'a_AlertTime03' ";
            PF.ExecSQL(vConnStr, vUpdateStr);
            vUpdateStr = (eAlertTime04.Text.Trim() != "") ?
                         "update SysFlag set Content = '" + eAlertTime04.Text.Trim() + "' where FormName = 'unSysFlag' and ControlItem = 'a_AlertTime04' " :
                         "update SysFlag set Content = NULL where FormName = 'unSysFlag' and ControlItem = 'a_AlertTime04' ";
            PF.ExecSQL(vConnStr, vUpdateStr);
            SetEditMode(false);
            OpenData();
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            SetEditMode(false);
            OpenData();
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}
/*
select MAX(case when ControlItem = 'a_AlertTime01' then Content else NULL end) as AlertTime01,
       MAX(case when ControlItem = 'a_AlertTime02' then Content else NULL end) as AlertTime02,
	   MAX(case when ControlItem = 'a_AlertTime03' then Content else NULL end) as AlertTime03,
	   MAX(case when ControlItem = 'a_AlertTime04' then Content else NULL end) as AlertTime04
  from SysFlag
 where FormName = 'unSysFlag'
   and ControlItem like 'a_AlertTime%'
*/