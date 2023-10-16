using Amaterasu_Function;
using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

namespace TyBus_Intranet_Test_V3
{
    public partial class SearchEmp : Page
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
                bbOK.Enabled = ((gridEmpData_Search != null) && (gridEmpData_Search.Rows.Count > 0));
            }
        }

        protected void eEmpName_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Int32 vTempInt = 0;
            string vEmpName = eEmpName_Search.Text.Trim();
            if (!Int32.TryParse(vEmpName,out vTempInt))
            {
                string vSelectStr = "SELECT EMPNO, NAME, worktype, convert(nvarchar(10), Leaveday, 111) Leaveday FROM EMPLOYEE WHERE (NAME like @EmpName) order by Leaveday";
                using (SqlDataSource dsEmp = new SqlDataSource(vConnStr, vSelectStr))
                {
                    dsEmp.SelectParameters.Clear();
                    dsEmp.SelectParameters.Add(new Parameter("EmpName", DbType.String, (vEmpName != "") ? "%" + vEmpName + "%" : "%%"));
                    dsEmp.Select(new DataSourceSelectArguments());
                    gridEmpData_Search.DataSource = dsEmp;
                    gridEmpData_Search.DataBind();
                }
            }
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
            if (gridEmpData_Search.SelectedRow != null)
            {
                int vIndex = gridEmpData_Search.SelectedRow.RowIndex;

                vTextBoxID = this.Request.QueryString["TextBoxID"];

                string vReturnCustNo = gridEmpData_Search.Rows[vIndex].Cells[1].Text.Trim();

                vScript = "opener.window.document.getElementById('" + vTextBoxID + "').value='" + vReturnCustNo + "', ";
                vScript += "window.close();";
                this.ClientScript.RegisterStartupScript(GetType(), "_Custom", vScript, true);
            }
        }
    }
}