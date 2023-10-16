using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;

namespace TyBus_Intranet_Test_V3
{
    public partial class ChangePWD : System.Web.UI.Page
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
                    eLoginID.Text = vLoginID + '_' + vLoginName;
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

        protected void bbOK_Click(object sender, EventArgs e)
        {
            string vNewPWD = eChangePWD.Text.Trim();
            if (Page.IsValid)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSQLStr = "update Employee set Password = '" + vNewPWD.Trim() + "' where EmpNo = '" + vLoginID.Trim() + "' ";
                PF.ExecSQL(vConnStr, vSQLStr);
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('您的密碼已經完成變更，請先登出之後再重新登入')");
                Response.Write("</" + "Script>");
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, "變更密碼");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('兩次輸入的密碼不同，請重新輸入')");
                Response.Write("</" + "Script>");
                eChangePWD.Focus();
            }
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}