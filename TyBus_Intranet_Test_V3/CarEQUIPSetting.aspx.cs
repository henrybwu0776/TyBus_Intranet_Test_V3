using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;
using System.Data;
using System.Data.SqlClient;

namespace TyBus_Intranet_Test_V3
{
    public partial class CarEQUIPSetting : System.Web.UI.Page
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
                        GetData();
                    }
                    else
                    {

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

        private string GetControlName(string fControlItem)
        {
            string vResultStr = "";
            int vResultIndex = 0;
            if (Int32.TryParse(fControlItem.Replace("a_ck", ""), out vResultIndex))
            {
                vResultStr = "eCK" + vResultIndex.ToString("D2");
            }
            else if (Int32.TryParse(fControlItem.Replace("a_lb", ""), out vResultIndex))
            {
                vResultStr = "eLB" + vResultIndex.ToString("D2");
            }
            else
            {
                vResultStr = "";
            }
            return vResultStr;
        }

        private void GetData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vControlName = "";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                string vSQLStrTemp = "select ControlItem, Content from SysFlag " + Environment.NewLine +
                                     " where FormName = 'unSysFlag' and (ControlItem like 'a_ck%' or (ControlItem like 'a_lb%')) " + Environment.NewLine +
                                     " order by ControlItem ";
                SqlCommand cmdTemp = new SqlCommand(vSQLStrTemp, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vControlName = GetControlName(drTemp["ControlItem"].ToString().Trim().ToLower());
                    if (vControlName != "")
                    {
                        TextBox eTemp = (TextBox)plData.FindControl(vControlName);
                        if (eTemp != null)
                        {
                            eTemp.Text = drTemp["Content"].ToString().Trim();
                        }
                    }
                }
            }

        }

        protected void bbUpdate_Click(object sender, EventArgs e)
        {
            string vControlItemName = "";
            string vContent = "";
            string vSQLStrTemp = "";
            int vRCount = 0;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            try
            {
                foreach (Control vControlItem in plData.Controls)
                {
                    if (vControlItem is TextBox)
                    {
                        vControlItemName = (vControlItem as TextBox).ToolTip.ToLower().Trim();
                        vContent = (vControlItem as TextBox).Text.Trim();
                        vSQLStrTemp = "select count(Content) RCount from SysFlag where FormName = 'unSysFlag' and ControlItem = '" + vControlItemName + "' ";
                        vRCount = Int32.Parse(PF.GetValue(vConnStr, vSQLStrTemp, "RCount"));
                        if (vRCount == 0)
                        {
                            vSQLStrTemp = "insert into SysFlag(FormName, ControlItem, Content)" + Environment.NewLine +
                                          "values ('unSysFlag', @ControlItem, @Content)";
                        }
                        else
                        {
                            vSQLStrTemp = "update SysFlag set Content = @Content " + Environment.NewLine +
                                          " where FormName = 'unSysFlag' and ControlItem = @ControlItem ";
                        }
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStrTemp, connTemp);
                            cmdTemp.Parameters.Clear();
                            cmdTemp.Parameters.Add(new SqlParameter("ControlItem", vControlItemName));
                            cmdTemp.Parameters.Add(new SqlParameter("Content", vContent));
                            connTemp.Open();
                            cmdTemp.ExecuteNonQuery();
                        }
                    }
                }
                GetData();
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            GetData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}