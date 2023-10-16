using Amaterasu_Function;
using System;
using System.Data.SqlClient;
using System.Web.UI.HtmlControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class MainPage : System.Web.UI.MasterPage
    {
        PublicFunction PF = new PublicFunction();
        private string vLoginID = "";
        private string vLoginName = "";
        private string vLoginDepNo = "";
        private string vLoginDepName = "";
        private string vLoginTitle = "";
        private string vLoginTitleName = "";
        private string vLoginEmpType = "";
        private string vConnStr = "";
        private string vSQLStr = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                vSQLStr = "select Content from SysFlag where FormName = 'unSysflag' and ControlItem = 'a_WebCompanyName'";
                TitleText.Text = PF.GetValue(vConnStr, vSQLStr, "Content");
                if (Session["LoginID"] == null)
                {
                    Session["LoginID"] = "";
                    Session["LoginName"] = "";
                    Session["LoginDepNo"] = "";
                    Session["LoginDepName"] = "";
                    Session["LoginTitle"] = "";
                    Session["LoginTitleName"] = "";
                    Session["LoginEmpType"] = "";
                }
            }
            vLoginID = (Session["LoginID"] != null) ? Session["LoginID"].ToString().Trim() : "";
            vLoginDepNo = (Session["LoginDepNo"] != null) ? Session["LoginDepNo"].ToString().Trim() : "";
            if (vLoginID != "")
            {
                plGuest.Visible = false;
                plMember.Visible = true;
                eLogInID_Log.Text = Session["LoginID"].ToString().Trim();
                eLoginMan_Log.Text = Session["LoginName"].ToString().Trim();
                vLoginName = Session["LoginName"].ToString().Trim();
                vLoginDepNo = Session["LoginDepNo"].ToString().Trim();
                eLoginDepName_Log.Text = Session["LoginDepName"].ToString().Trim();
                vLoginDepName = Session["LoginDepName"].ToString().Trim();
                vLoginTitle = Session["LoginTitle"].ToString().Trim();
                eLoginTitle_Log.Text = Session["LoginTitleName"].ToString().Trim();
                vLoginTitleName = Session["LoginTitleName"].ToString().Trim();
                vLoginEmpType = Session["LoginEmpType"].ToString().Trim();
                CheckPermission(vLoginID);
            }
            else
            {
                plGuest.Visible = true;
                plMember.Visible = false;
                HideAllList();
            }
        }

        protected void bbLogin_Click(object sender, EventArgs e)
        {
            string vTempEmpNo = eAccount_Log.Text.Trim();
            string vEmpDataStr = "";
            string vEmpPassword = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select count(EmpNo) RCount from Employee where EmpNo = '" + vTempEmpNo + "' and isnull(LeaveDay, '') = '' ";
            int vRCount = 0;
            if (Int32.TryParse(PF.GetValue(vConnStr, vSQLStr, "RCount"), out vRCount))
            {
                if (vRCount != 0)
                {
                    vSQLStr = "select top 1 EmpNo + ',' + isnull([Name], '') + ',' + " + Environment.NewLine +
                              "       isnull(DepNo, '') + ',' + isnull((select LTrim(RTrim([Name])) from Department where DepNo = e.DepNo), '') +',' + " + Environment.NewLine +
                              "       isnull(Title, '') + ',' + isnull((select LTrim(RTrim(ClassTxt)) from DBDICB where FKey = '員工資料        EMPLOYEE        TITLE' and ClassNo = e.Title), '') +',' + " + Environment.NewLine +
                              "       isnull([Type], '') + ',' + isnull(Password, '') as EmpData " + Environment.NewLine +
                              "  from Employee e " + Environment.NewLine +
                              " where isnull(LeaveDay, '') = '' and EmpNo = '" + vTempEmpNo + "' " + Environment.NewLine +
                              " order by isnull(LeaveDay, '9999/12/31')";
                    vEmpDataStr = PF.GetValue(vConnStr, vSQLStr, "EmpData");
                    string[] aEmpData = vEmpDataStr.Split(',');
                    //更新變數
                    vLoginID = aEmpData[0].Trim();
                    vLoginName = aEmpData[1].Trim();
                    vLoginDepNo = aEmpData[2].Trim();
                    vLoginDepName = aEmpData[3].Trim();
                    vLoginTitle = aEmpData[4].Trim();
                    vLoginTitleName = aEmpData[5].Trim();
                    vLoginEmpType = aEmpData[6].Trim();
                    vEmpPassword = aEmpData[7].Trim();
                    if (vEmpPassword == ePassword_Log.Text.Trim())
                    {

                        //寫入 Session
                        Session["LoginID"] = aEmpData[0].Trim();
                        Session["LoginName"] = aEmpData[1].Trim();
                        Session["LoginDepNo"] = aEmpData[2].Trim();
                        Session["LoginDepName"] = aEmpData[3].Trim();
                        Session["LoginTitle"] = aEmpData[4].Trim();
                        Session["LoginTitleName"] = aEmpData[5].Trim();
                        Session["LoginEmpType"] = aEmpData[6].Trim();
                        //更新畫面上的欄位
                        eLoginDepName_Log.Text = aEmpData[3].Trim();
                        eLoginTitle_Log.Text = aEmpData[5].Trim();
                        eLogInID_Log.Text = aEmpData[0].Trim();
                        eLoginMan_Log.Text = aEmpData[1].Trim();
                        //讀取選單權限
                        CheckPermission(aEmpData[0].Trim());

                        //導向起始頁面
                        Response.Redirect("~/default.aspx");
                    }
                    else
                    {
                        Session["LoginID"] = "";
                        Session["LoginName"] = "";
                        Session["LoginDepNo"] = "";
                        Session["LoginDepName"] = "";
                        Session["LoginTitle"] = "";
                        Session["LoginTitleName"] = "";
                        Session["LoginEmpType"] = "";
                        eAccount_Log.Text = "";
                        eAccount_Log.Focus();
                        Response.Write("<Script language='Javascript'>");
                        Response.Write("alert('您輸入的帳號密碼錯誤')");
                        Response.Write("</" + "Script>");
                    }
                }
                else
                {
                    Session["LoginID"] = "";
                    Session["LoginName"] = "";
                    Session["LoginDepNo"] = "";
                    Session["LoginDepName"] = "";
                    Session["LoginTitle"] = "";
                    Session["LoginTitleName"] = "";
                    Session["LoginEmpType"] = "";
                    eAccount_Log.Text = "";
                    eAccount_Log.Focus();
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('您輸入的帳號不存在')");
                    Response.Write("</" + "Script>");
                }
            }
            else
            {
                Session["LoginID"] = "";
                Session["LoginName"] = "";
                Session["LoginDepNo"] = "";
                Session["LoginDepName"] = "";
                Session["LoginTitle"] = "";
                Session["LoginTitleName"] = "";
                Session["LoginEmpType"] = "";
                eAccount_Log.Text = "";
                eAccount_Log.Focus();
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('您輸入的帳號不存在')");
                Response.Write("</" + "Script>");
            }
        }

        protected void bbLogOut_Click(object sender, EventArgs e)
        {
            Session["LoginID"] = "";
            Session["LoginName"] = "";
            Session["LoginDepNo"] = "";
            Session["LoginDepName"] = "";
            Session["LoginTitle"] = "";
            Session["LoginTitleName"] = "";
            Session["LoginEmpType"] = "";
            Response.Redirect("~/default.aspx");
        }

        protected void HideAllList()
        {
            string vTempControlName = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            vSQLStr = "select ControlName from WebPermissionA order by OrderIndex";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vTempControlName = drTemp["ControlName"].ToString().Trim();
                    HtmlGenericControl liTemp = (HtmlGenericControl)FindControl(vTempControlName);
                    if (liTemp != null)
                    {
                        liTemp.Attributes.Remove("class");
                        liTemp.Attributes.Add("class", "NoneDisplay");
                    }
                }
                drTemp.Close();
                cmdTemp.Cancel();
                connTemp.Close();

                vSQLStr = "select ClassNo from DBDICB where FKey = '網頁功能權限群組WebPermission   GroupID' order by ClassNo";
                cmdTemp = new SqlCommand(vSQLStr, connTemp);
                connTemp.Open();
                drTemp = cmdTemp.ExecuteReader();
                while (drTemp.Read())
                {
                    vTempControlName = drTemp["ClassNo"].ToString().Trim();
                    HtmlGenericControl liTemp = (HtmlGenericControl)FindControl(vTempControlName);
                    if (liTemp != null)
                    {
                        liTemp.Attributes.Remove("class");
                        liTemp.Attributes.Add("class", "NoneDisplay");
                    }
                }
                drTemp.Close();
                cmdTemp.Cancel();
            }
        }

        protected void CheckPermission(string fLoginID)
        {
            string vTempControlName = "";
            HideAllList(); //先把功能表設為不顯示
            if ((vLoginEmpType != "") && (vLoginEmpType != "20")) //排除行車人員
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    //比對部門權限
                    vSQLStr = "select ControlName from WebPermissionB where DepNo = '" + vLoginDepNo + "' and AllowPermission = 1";
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        vTempControlName = drTemp["ControlName"].ToString().Trim();
                        HtmlGenericControl liTemp = (HtmlGenericControl)FindControl(vTempControlName);
                        if (liTemp != null)
                        {
                            liTemp.Attributes.Remove("class");
                        }
                    }
                    drTemp.Close();
                    cmdTemp.Cancel();
                    connTemp.Close();

                    vSQLStr = "select distinct GroupID " + Environment.NewLine +
                              "  from WebPermissionA " + Environment.NewLine +
                              " where isnull(GroupID, '') <> '' " + Environment.NewLine +
                              "   and ControlName in (select ControlName from WebPermissionB where DepNo = '" + vLoginDepNo + "' and AllowPermission = 1 )";
                    cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        vTempControlName = drTemp["GroupID"].ToString().Trim();
                        HtmlGenericControl liTemp = (HtmlGenericControl)FindControl(vTempControlName);
                        if (liTemp != null)
                        {
                            liTemp.Attributes.Remove("class");
                        }
                    }
                    drTemp.Close();
                    cmdTemp.Cancel();
                    connTemp.Close();

                    //比對單位有權限但是個人沒有權限的部份
                    vSQLStr = "select distinct GroupID " + Environment.NewLine +
                              "  from WebPermissionA " + Environment.NewLine +
                              " where isnull(GroupID, '') <> '' " + Environment.NewLine +
                              "   and ControlName in (select ControlName from WebPermissionB where DepNo = '" + vLoginDepNo + "' and AllowPermission = 1 )" + Environment.NewLine +
                              "   and ControlName in (select ControlName from WebPermissionB where EmpNo = '" + vLoginID + "' and AllowPermission = 0)";
                    cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        vTempControlName = drTemp["GroupID"].ToString().Trim();
                        HtmlGenericControl liTemp = (HtmlGenericControl)FindControl(vTempControlName);
                        if (liTemp != null)
                        {
                            liTemp.Attributes.Remove("class");
                            liTemp.Attributes.Add("class", "NoneDisplay");
                        }
                    }
                    drTemp.Close();
                    cmdTemp.Cancel();
                    connTemp.Close();

                    //比對個人有權限的部份
                    vSQLStr = "select ControlName from WebPermissionB where EmpNo = '" + vLoginID + "' and AllowPermission = 1";
                    cmdTemp.CommandText = vSQLStr;
                    connTemp.Open();
                    drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        vTempControlName = drTemp["ControlName"].ToString().Trim();
                        HtmlGenericControl liTemp = (HtmlGenericControl)FindControl(vTempControlName);
                        if (liTemp != null)
                        {
                            liTemp.Attributes.Remove("class");
                        }
                    }
                    drTemp.Close();
                    cmdTemp.Cancel();
                    connTemp.Close();

                    vSQLStr = "select distinct GroupID " + Environment.NewLine +
                              "  from WebPermissionA " + Environment.NewLine +
                              " where isnull(GroupID, '') <> '' " + Environment.NewLine +
                              "   and ControlName in (select ControlName from WebPermissionB where EmpNo = '" + vLoginID + "' and AllowPermission = 1 )";
                    cmdTemp = new SqlCommand(vSQLStr, connTemp);
                    connTemp.Open();
                    drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        vTempControlName = drTemp["GroupID"].ToString().Trim();
                        HtmlGenericControl liTemp = (HtmlGenericControl)FindControl(vTempControlName);
                        if (liTemp != null)
                        {
                            liTemp.Attributes.Remove("class");
                        }
                    }
                    drTemp.Close();
                    cmdTemp.Cancel();
                    connTemp.Close();

                    //比對個人權限拒絕開放的部份
                    vSQLStr = "select ControlName from WebPermissionB where EmpNo = '" + vLoginID + "' and AllowPermission = 0";
                    cmdTemp.CommandText = vSQLStr;
                    connTemp.Open();
                    drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        vTempControlName = drTemp["ControlName"].ToString().Trim();
                        HtmlGenericControl liTemp = (HtmlGenericControl)FindControl(vTempControlName);
                        if (liTemp != null)
                        {
                            liTemp.Attributes.Remove("class");
                            liTemp.Attributes.Add("class", "NoneDisplay");
                        }
                    }
                    drTemp.Close();
                    cmdTemp.Cancel();
                }
            }
        }
    }
}