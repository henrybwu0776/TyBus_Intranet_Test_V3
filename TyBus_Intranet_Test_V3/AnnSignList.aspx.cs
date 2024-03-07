using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;

namespace TyBus_Intranet_Test_V3
{
    public partial class AnnSignList : Page
    {
        PublicFunction PF = new PublicFunction(); //加入公用程式碼參考
        private string vLoginID = "";
        private string vConnStr = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"]!=null)
            {
                vLoginID = Session["LoginID"].ToString().Trim();
                if (vLoginID!="")
                {
                    if (vConnStr=="")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    if (!IsPostBack)
                    {
                        plSearch.Visible = true;
                        plReport.Visible = false;
                    }
                }
                else
                {
                    Response.Redirect("~/Default.aspx");
                }
            }
            else
            {
                Response.Redirect("~/Default.aspx");
            }
        }

        protected void cbEmpType_CheckedChanged(object sender, EventArgs e)
        {
            eEmpType.Text = (cbEmpType.Checked) ? "Y" : "N";
        }

        protected void bbPrint_Click(object sender, EventArgs e)
        {

        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plSearch.Visible = true;
            plReport.Visible = false;
        }
    }
}