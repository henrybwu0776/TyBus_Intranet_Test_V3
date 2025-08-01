﻿using Amaterasu_Function;
using System;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class SearchSup : Page
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
                bbOK.Enabled = ((gridSupData_Search != null) && (gridSupData_Search.Rows.Count > 0));
            }
            OpenData();
        }

        protected void eSupName_Search_TextChanged(object sender, EventArgs e)
        {

        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            string vScript = "";
            vScript = "window.close();";
            this.ClientScript.RegisterStartupScript(GetType(), "_SupNo", vScript, true);
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            string vScript = "";
            string vTextBoxID = "";
            if (gridSupData_Search.SelectedRow != null)
            {
                int vIndex = gridSupData_Search.SelectedRow.RowIndex;

                vTextBoxID = this.Request.QueryString["TextBoxID"];
                string vSupNameID = this.Request.QueryString["SupNameID"];

                string vReturnSupNo = gridSupData_Search.Rows[vIndex].Cells[1].Text.Trim();
                string vReturnSupName = gridSupData_Search.Rows[vIndex].Cells[2].Text.Trim();

                vScript = "opener.window.document.getElementById('" + vTextBoxID + "').value='" + vReturnSupNo + "', ";
                vScript = ((vSupNameID != null) && (vSupNameID != "")) ?
                          vScript + "opener.window.document.getElementById('" + vSupNameID + "').value='" + vReturnSupName + "', " :
                          vScript;
                vScript += "window.close();";
                this.ClientScript.RegisterStartupScript(GetType(), "_Product", vScript, true);
            }
        }

        private void OpenData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Int32 vTempInt = 0;
            string vSupName = eSupName_Search.Text.Trim();
            string vSupKind = this.Request.QueryString["SupKind"];
            string vSelectStr = "";
            if (!Int32.TryParse(vSupName, out vTempInt))
            {
                vSelectStr = (vSupKind == "C") ? "select Code, [Name] from Custom where Types = 'S' and Kind in ('U', 'C', 'W') and [Name] like @Name " :
                                                 "select Code, [Name] from Custom where Types = 'S' and Kind in ('U', 'W') and [Name] like @Name ";
                using (SqlDataSource dsTemp = new SqlDataSource(vConnStr, vSelectStr))
                {
                    dsTemp.SelectParameters.Clear();
                    dsTemp.SelectParameters.Add(new Parameter("Name", DbType.String, (vSupName != "") ? "%" + vSupName + "%" : "%%"));
                    dsTemp.Select(new DataSourceSelectArguments());
                    gridSupData_Search.DataSource = dsTemp;
                    gridSupData_Search.DataBind();
                }
            }
            /* 2023.02.23 修改資料取回方式
            dsSupData.SelectCommand = "select Code, [Name] from Custom where Types = 'S' and [Name] like @Name ";
            dsSupData.SelectParameters.Clear();
            dsSupData.SelectParameters.Add(new Parameter("Name", System.Data.DbType.String, (vSupName != "") ? "%" + vSupName + "%" : "%%"));
            dsSupData.Select(new DataSourceSelectArguments());
            gridSupData_Search.DataBind();
            */
        }

        protected void bbShow_Search_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void gridSupData_Search_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridSupData_Search.PageIndex = e.NewPageIndex;
            gridSupData_Search.DataBind();
        }
    }
}