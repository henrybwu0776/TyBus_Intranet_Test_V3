using Amaterasu_Function;
using System;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class EmpMultiDep : System.Web.UI.Page
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
        private DateTime vToday = DateTime.Today;

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
                        fvEmpDepDetails.Visible = false;
                    }
                    else
                    {
                        EmpDataBind();
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

        private string GetSelStr()
        {
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ?
                                 "   and e.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine :
                                 "";
            string vWStr_EmpNo = (eEmpNo_Search.Text.Trim() != "") ?
                                 "   and e.EmpNo = '" + eEmpNo_Search.Text.Trim() + "' " + Environment.NewLine :
                                 "";
            string vWStr_WorkType = (ddlWorkType_Search.SelectedValue.Trim() != "") ?
                                    "   and e.WorkType = '" + ddlWorkType_Search.SelectedValue.Trim() + "' " + Environment.NewLine :
                                    "";
            string vSelectStr = "SELECT EMPNO, NAME, DEPNO, (SELECT [NAME] FROM DEPARTMENT WHERE (DEPNO = e.DEPNO)) AS DepName, TYPE, " + Environment.NewLine +
                                "       (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '人事資料檔      EMPLOYEE        type') AND (CLASSNO = e.TYPE)) AS Type_C, " + Environment.NewLine +
                                "       worktype, TITLE, (SELECT CLASSTXT FROM DBDICB AS DBDICB_1 WHERE (FKEY = '員工資料        EMPLOYEE        TITLE') AND (CLASSNO = e.TITLE)) AS Title_C " + Environment.NewLine +
                                "  FROM EMPLOYEE AS e " + Environment.NewLine +
                                " where e.DepNo <> '00' " + Environment.NewLine +
                                "   and e.Type <> '20' " + Environment.NewLine + vWStr_DepNo + vWStr_EmpNo + vWStr_WorkType +
                                " order by e.DepNo, e.Title, e.EmpNo";
            return vSelectStr;
        }

        private void EmpDataBind()
        {
            string vSelStr = GetSelStr();
            sdsEmpList.SelectCommand = "";
            sdsEmpList.SelectCommand = vSelStr;
            sdsEmpList.DataBind();
        }

        protected void eEmpNo_Search_TextChanged(object sender, EventArgs e)
        {
            string vEmpNo = eEmpNo_Search.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo + "' ";
            string vEmpName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName == "")
            {
                vEmpName = vEmpNo;
                vSQLStr = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName + "' order by EmpNo DESC";
                vEmpNo = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eEmpNo_Search.Text = vEmpNo;
            lbEmpName_Search.Text = vEmpName;
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            string vDepNo_Temp = eDepNo_Search.Text.Trim();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp;
                vSQLStr = "select top 1 DepNo from Department where [Name] = '" + vDepName_Temp + "' order by DepNo DESC";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Search.Text = vDepNo_Temp;
            lbDepName_Search.Text = vDepName_Temp;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            EmpDataBind();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void gridEmpList_DataBound(object sender, EventArgs e)
        {
            fvEmpDepDetails.Visible = false;
        }

        protected void gridEmpList_SelectedIndexChanged(object sender, EventArgs e)
        {
            fvEmpDepDetails.Visible = true;
        }

        protected void fvEmpDepDetails_DataBound(object sender, EventArgs e)
        {
            switch (fvEmpDepDetails.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    break;
                case FormViewMode.Insert:
                    string vEmpNo_Temp = "";
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    TextBox eEmpNo_INS = (TextBox)fvEmpDepDetails.FindControl("eEmpNo_INS");
                    Label eEmpNo_C_INS = (Label)fvEmpDepDetails.FindControl("eEmpNo_C_INS");
                    Label lbBuMan_INS = (Label)fvEmpDepDetails.FindControl("eBuMan_INS");
                    Label lbBuMan_C_INS = (Label)fvEmpDepDetails.FindControl("eBuMan_C_INS");
                    Label lbBuDate_INS = (Label)fvEmpDepDetails.FindControl("eBuDate_INS");
                    //vEmpNo_Temp = gridDepDetails.SelectedRow.Cells[2].Text.Trim();
                    vEmpNo_Temp = gridEmpList.SelectedRow.Cells[1].Text.Trim();
                    eEmpNo_INS.Text = vEmpNo_Temp;
                    eEmpNo_C_INS.Text = (vEmpNo_Temp != "") ? PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vEmpNo_Temp + "' ", "Name") : "";
                    lbBuDate_INS.Text = DateTime.Today.ToShortDateString();
                    lbBuMan_INS.Text = vLoginID;
                    lbBuMan_C_INS.Text = PF.GetValue(vConnStr, "select [Name] from EMployee where EmpNo = '" + vLoginID + "' ", "Name");
                    break;
                default:
                    break;
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo_Edit = (TextBox)fvEmpDepDetails.FindControl("eDepNo_Edit");
            Label eDepNo_C_Edit = (Label)fvEmpDepDetails.FindControl("eDepNo_C_Edit");
            string vDepNo_Edit = eDepNo_Edit.Text.Trim();
            string vDepName_Edit = "";
            string vSQLStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_Edit + "' ";
            vDepName_Edit = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_Edit == "")
            {
                vDepName_Edit = vDepNo_Edit;
                vSQLStr = "select top 1 DepNo from Department where [Name] = '" + vDepName_Edit + "' order by DepnO desc";
                vDepNo_Edit = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Edit.Text = vDepNo_Edit;
            eDepNo_C_Edit.Text = vDepName_Edit;
        }

        protected void eEmpNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eEmpNo_INS = (TextBox)fvEmpDepDetails.FindControl("eEmpNo_INS");
            Label eEmpNo_C_INS = (Label)fvEmpDepDetails.FindControl("eEmpNo_C_INS");
            string vEmpNo_INS = eEmpNo_INS.Text.Trim();
            string vEmpName_INS = "";
            string vSQLStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select [Name] from Employee where EmpNo = '" + vEmpNo_INS + "' ";
            vEmpName_INS = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vEmpName_INS == "")
            {
                vEmpName_INS = vEmpNo_INS;
                vSQLStr = "select top 1 EmpNo from Employee where [Name] = '" + vEmpName_INS + "' order by EmpNo DESC";
                vEmpNo_INS = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eEmpNo_INS.Text = vEmpNo_INS;
            eEmpNo_C_INS.Text = vEmpName_INS;
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo_INS = (TextBox)fvEmpDepDetails.FindControl("eDepNo_INS");
            Label eDepNo_C_INS = (Label)fvEmpDepDetails.FindControl("eDepNo_C_INS");
            string vDepNo_INS = eDepNo_INS.Text.Trim();
            string vDepName_INS = "";
            string vSQLStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo_INS + "' ";
            vDepName_INS = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName_INS == "")
            {
                vDepName_INS = vDepNo_INS;
                vSQLStr = "select top 1 DepNo from Department where [Name] = '" + vDepName_INS + "' order by DepnO desc";
                vDepNo_INS = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_INS.Text = vDepNo_INS;
            eDepNo_C_INS.Text = vDepName_INS;
        }

        protected void sdsEmpDepDetails_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridDepDetails.DataBind();
                plSearch.Visible = true;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('刪除完成！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + e.Exception.Message + "')");
                Response.Write("</" + "Script>");
            }
        }

        protected void sdsEmpDepDetails_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridDepDetails.DataBind();
                plSearch.Visible = true;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('新增完成！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + e.Exception.Message + "')");
                Response.Write("</" + "Script>");
            }
        }

        protected void sdsEmpDepDetails_Inserting(object sender, SqlDataSourceCommandEventArgs e)
        {
            TextBox eEmpNo_INS = (TextBox)fvEmpDepDetails.FindControl("eEmpNo_INS");
            Label eBuDate_INS = (Label)fvEmpDepDetails.FindControl("eBuDate_INS");
            Label eBuMan_INS = (Label)fvEmpDepDetails.FindControl("eBuMan_INS");
            string vEmpNo_INS = eEmpNo_INS.Text.Trim();
            DateTime vBuDate_INS = (eBuDate_INS.Text.Trim() != "") ? DateTime.Parse(eBuDate_INS.Text.Trim()) : DateTime.Today;
            string vBuMan_INS = eBuMan_INS.Text.Trim();
            string vItemsStr = "";
            int vItems = 0;
            string vSQLStr = "select isnull(MAX(Items), '0000') MaxItems from EmployeeDepNo where EmpNo = '" + vEmpNo_INS + "' ";
            vItems = Int32.Parse(PF.GetValue(vConnStr, vSQLStr, "MaxItems")) + 1;
            vItemsStr = vItems.ToString("D4");
            string vEmpNoItems = vEmpNo_INS + vItemsStr;
            e.Command.Parameters["@EmpNoItems"].Value = vEmpNoItems;
            e.Command.Parameters["@Items"].Value = vItemsStr;
            e.Command.Parameters["@BuDate"].Value = vBuDate_INS;
            e.Command.Parameters["@BuMan"].Value = vBuMan_INS;
        }

        protected void sdsEmpDepDetails_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                this.gridDepDetails.DataBind();
                plSearch.Visible = true;
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('修改完成！')");
                Response.Write("</" + "Script>");
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + e.Exception.Message + "')");
                Response.Write("</" + "Script>");
            }
        }
    }
}