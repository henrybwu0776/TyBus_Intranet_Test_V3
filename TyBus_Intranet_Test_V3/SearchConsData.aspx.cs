using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class SearchConsData : Page
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
                ddlConsType.Items.Clear();
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    string vSelectStr = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                        " union all " + Environment.NewLine +
                                        "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsType' " + Environment.NewLine +
                                        " order by ClassNo ";
                    SqlCommand cmdTemp = new SqlCommand(vSelectStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                        ddlConsType.Items.Add(liTemp);
                    }
                }
            }
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            string vWStr_ConsType = (eConsType.Text.Trim() != "") ? "   AND ConsType = @ConsType" + Environment.NewLine : "";
            string vConsName = "%" + eConsName.Text.Trim() + "%";
            string vSelectStr = "SELECT ConsNo, ConsName, StockQty, ConsSpec " + Environment.NewLine +
                                "  FROM Consumables " + Environment.NewLine +
                                " WHERE (ISNULL(ConsNo, '') != '') " + Environment.NewLine +
                                vWStr_ConsType +
                                "   AND ConsName like @ConsName ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlDataSource sdsConsList=new SqlDataSource(vConnStr,vSelectStr))
            {
                sdsConsList.SelectParameters.Clear();
                sdsConsList.SelectParameters.Add(new Parameter("ConsName", DbType.String, vConsName));
                if (eConsType.Text.Trim() != "")
                {
                    sdsConsList.SelectParameters.Add(new Parameter("ConsType", DbType.String, eConsType.Text.Trim()));
                }
                sdsConsList.Select(new DataSourceSelectArguments());
                gridConsList.DataSource = sdsConsList;
                gridConsList.DataBind();
            }
        }

        protected void gridConsList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridConsList.PageIndex = e.NewPageIndex;
            gridConsList.DataBind();
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
            int vIndex = gridConsList.SelectedRow.RowIndex;
            string vConsNoID = this.Request.QueryString["ConsNoID"];
            string vConsNameID = this.Request.QueryString["ConsNameID"];

            string vReturnProductNo = gridConsList.Rows[vIndex].Cells[1].Text.Trim();
            string vReturnProductName = gridConsList.Rows[vIndex].Cells[2].Text.Trim();

            vScript += "opener.window.document.getElementById('" + vConsNoID + "').value='" + vReturnProductNo + "', ";
            vScript = ((vConsNameID != null) && (vConsNameID.Trim() != "")) ?
                      vScript + "opener.window.document.getElementById('" + vConsNameID + "').value='" + vReturnProductName + "', " :
                      vScript;
            vScript += "window.close();";
            this.ClientScript.RegisterStartupScript(GetType(), "_Product", vScript, true);
        }

        protected void ddlConsType_SelectedIndexChanged(object sender, EventArgs e)
        {
            eConsType.Text = ddlConsType.SelectedValue.Trim();
        }
    }
}