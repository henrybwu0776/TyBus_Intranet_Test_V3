using Amaterasu_Function;
using System;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class SearchCust : System.Web.UI.Page
    {
        PublicFunction PF = new PublicFunction(); //加入公用程式碼參考
        private string vConnStr = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (!IsPostBack)
            {
                bbOK.Enabled = false;
            }
            else
            {
                bbOK.Enabled = ((gridCustData_Search != null) && (gridCustData_Search.Rows.Count > 0));
            }
        }

        protected void eCustName_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Int32 vTempInt = 0;
            string vCustName = eCustName_Search.Text.Trim();
            if (!Int32.TryParse(vCustName, out vTempInt))
            {
                string vSelectStr = "select Code, [Name] from Custom where Types = 'C' and [Name] like @Name ";
                using (SqlDataSource dsTemp=new SqlDataSource(vConnStr, vSelectStr))
                {
                    dsTemp.SelectParameters.Clear();
                    dsTemp.SelectParameters.Add(new Parameter("Name", DbType.String, (vCustName != "") ? "%" + vCustName + "%" : "%%"));
                    dsTemp.Select(new DataSourceSelectArguments());
                    gridCustData_Search.DataSource = dsTemp;
                    gridCustData_Search.DataBind();
                }
            }
            /* 2023.02.23 修改資料取回方式
            dsCustData.SelectCommand = "select Code, [Name] from Custom where Types = 'C' and [Name] like @Name ";
            dsCustData.SelectParameters.Clear();
            dsCustData.SelectParameters.Add(new Parameter("Name", System.Data.DbType.String, (vCustName != "") ? "%" + vCustName + "%" : "%%"));
            dsCustData.Select(new DataSourceSelectArguments());
            gridCustData_Search.DataBind();
            */
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            string vScript = "";
            vScript = "window.close();";
            this.ClientScript.RegisterStartupScript(GetType(), "_Custom", vScript, true);
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            string vScript = "";
            string vTextBoxID = "";
            if (gridCustData_Search.SelectedRow != null)
            {
                int vIndex = gridCustData_Search.SelectedRow.RowIndex;

                vTextBoxID = this.Request.QueryString["TextBoxID"];

                string vReturnCustNo = gridCustData_Search.Rows[vIndex].Cells[1].Text.Trim();

                vScript = "opener.window.document.getElementById('" + vTextBoxID + "').value='" + vReturnCustNo + "', ";
                vScript += "window.close();";
                this.ClientScript.RegisterStartupScript(GetType(), "_Custom", vScript, true);
            }
        }
    }
}