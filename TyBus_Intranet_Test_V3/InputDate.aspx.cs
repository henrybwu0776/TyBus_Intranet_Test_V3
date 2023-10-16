using System;

namespace TyBus_Intranet_Test_V3
{
    public partial class InputDate : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            string vScript = "";
            string vTextBoxID = "";

            vTextBoxID = this.Request.QueryString["TextBoxID"];
            string vReturnDateStr = (ChoiseDate.SelectedDate.Year.ToString("D4") != "0001") ?
                                    ChoiseDate.SelectedDate.Date.ToShortDateString() :
                                    DateTime.Today.ToShortDateString();
            vScript = "opener.window.document.getElementById('" + vTextBoxID + "').value='" + vReturnDateStr + "',";
            vScript = vScript + "window.close();";
            this.ClientScript.RegisterStartupScript(GetType(), "_Calender", vScript, true);
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            string vScript = "";
            vScript = "window.close();";
            this.ClientScript.RegisterStartupScript(GetType(), "_Calender", vScript, true);
        }
    }
}