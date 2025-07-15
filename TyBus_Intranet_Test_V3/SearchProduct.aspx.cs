using Amaterasu_Function;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class SearchProduct : System.Web.UI.Page
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
                bbOK.Enabled = ((gridProductData_Search != null) && (gridProductData_Search.Rows.Count > 0));
            }
        }

        protected void eProductName_Search_TextChanged(object sender, EventArgs e)
        {
            string vProductNameStr = "%" + eProductName_Search.Text.Trim() + "%";
            dsProduct.SelectCommand = "select ConsNo as ProductNo, (isnull(CorrespondModel, '') + '-' + isnull(ConsName, '') + '(' + isnull(Spec_Color, '') + ')') as ProductName, Price " + Environment.NewLine +
                                      "  from IAConsumables " + Environment.NewLine +
                                      " where isnull(ConsNo, '') <> '' " + Environment.NewLine +
                                      "   and (CorrespondModel like @CorrespondModel or ConsName like @ConsName) ";
            dsProduct.SelectParameters.Clear();
            dsProduct.SelectParameters.Add(new Parameter("CorrespondModel", System.Data.DbType.String, (eProductName_Search.Text.Trim() != "") ? vProductNameStr : String.Empty));
            dsProduct.SelectParameters.Add(new Parameter("ConsName", System.Data.DbType.String, (eProductName_Search.Text.Trim() != "") ? vProductNameStr : String.Empty));
            dsProduct.Select(new DataSourceSelectArguments());
            gridProductData_Search.DataBind();
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            string vScript = "";
            vScript = "window.close();";
            this.ClientScript.RegisterStartupScript(GetType(), "_Product", vScript, true);
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            string vScript = "";
            //string vTextBoxID = "";
            //string vLabelID = "";
            int vIndex = gridProductData_Search.SelectedRow.RowIndex;

            //vTextBoxID = this.Request.QueryString["TextBoxID"];
            //vLabelID = this.Request.QueryString["LabelID"];
            //2023.01.12 修改
            string vConsNoID = this.Request.QueryString["ConsNoID"];
            //string vConsNameID = this.Request.QueryString["ConsNameID"];
            string vConsPriceID = this.Request.QueryString["ConsPriceID"];
            //=====================================================================================

            string vReturnProductNo = gridProductData_Search.Rows[vIndex].Cells[1].Text.Trim();
            string vReturnProductName = gridProductData_Search.Rows[vIndex].Cells[2].Text.Trim();
            string vReturnProductPrice = gridProductData_Search.Rows[vIndex].Cells[3].Text.Trim();

            //vScript = "opener.window.document.getElementByID('" + vLabelID + "').value='" + vReturnProductName + ", ";
            //vScript += "opener.window.document.getElementById('" + vTextBoxID + "').value='" + vReturnProductNo + "', ";
            vScript += "opener.window.document.getElementById('" + vConsNoID + "').value='" + vReturnProductNo + "', ";
            /*vScript = ((vConsNameID != null) && (vConsNameID.Trim() != "")) ? 
                      vScript + "opener.window.document.getElementById('" + vConsNameID + "').value='" + vReturnProductName + "', " : 
                      vScript; //*/
            vScript = ((vConsPriceID != null) && (vConsPriceID.Trim() != "")) ? 
                      vScript + "opener.window.document.getElementById('" + vConsPriceID + "').value='" + vReturnProductPrice + "', " : 
                      vScript;
            vScript += "window.close();";
            this.ClientScript.RegisterStartupScript(GetType(), "_Product", vScript, true);
        }
    }
}