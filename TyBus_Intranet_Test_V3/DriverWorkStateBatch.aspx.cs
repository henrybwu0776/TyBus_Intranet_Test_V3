using Amaterasu_Function;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.Threading;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class DriverWorkStateBatch : System.Web.UI.Page
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
        private string vSHReport = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                Thread.CurrentThread.CurrentCulture = Cal;

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
                        vSHReport = PF.GetValue(vConnStr, "select InSHReport from Department where DepNo = '" + vLoginDepNo.Trim() + "' ", "InSHReport");
                        if (vSHReport.Trim() == "V")
                        {
                            eDepNo_Edit.Text = vLoginDepNo;
                            eDepName_Edit.Text = vLoginDepName;
                            eDepNo_Edit.Enabled = false;
                        }
                    }
                    else
                    {

                    }
                    string vBuildDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eSheetDate_Edit.ClientID;
                    string vBuildDateScript = "window.open('" + vBuildDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eSheetDate_Edit.Attributes["onClick"] = vBuildDateScript;
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

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName == "")
            {
                vDepName = eDepNo_Edit.Text.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '" + vDepName + "%' order by DepNo DESC";
                vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Edit.Text = vDepNo.Trim();
            eDepName_Edit.Text = vDepName.Trim();
        }

        protected void bbGetDriverList_Click(object sender, EventArgs e)
        {
            if (eDepNo_Edit.Text.Trim() != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                eDriverList_Edit.Items.Clear();
                eFixedDriver_Edit.Items.Clear();
                string vDepNo = eDepNo_Edit.Text.Trim();
                DateTime vBuDate = DateTime.Parse(eSheetDate_Edit.Text.Trim());
                string vBuDateStr = vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd");
                string vSQLStr_Temp = "select a.Driver, (b.EmpNo + '_' + b.[Name]) as [Name] " + Environment.NewLine +
                                      "  from RunSheetA a left join Employee b on b.EmpNo = a.Driver " + Environment.NewLine +
                                      " where a.DepNo = @DepNo " + Environment.NewLine +
                                      "   and a.BuDate = @BuDate " + Environment.NewLine +
                                      " order by a.Driver ";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                    cmdTemp.Parameters.Clear();
                    cmdTemp.Parameters.Add(new SqlParameter("DepNo", vDepNo));
                    cmdTemp.Parameters.Add(new SqlParameter("BuDate", vBuDate));
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        eDriverList_Edit.Items.Add(new ListItem(drTemp["Name"].ToString().Trim(), drTemp["Driver"].ToString().Trim()));
                    }
                }
            }
        }

        protected void bbGetDriver_Edit_Click(object sender, EventArgs e)
        {
            bool vAlreadyIn = false;
            foreach (ListItem vItem in eDriverList_Edit.Items)
            {
                if (vItem.Selected)
                {
                    vAlreadyIn = false;
                    for (int i = 0; i < eFixedDriver_Edit.Items.Count; i++)
                    {
                        if (vItem.Value == eFixedDriver_Edit.Items[i].Value) vAlreadyIn = true;
                    }
                    if (!vAlreadyIn) eFixedDriver_Edit.Items.Add(vItem);
                }
            }
        }

        protected void bbCancelDriver_Edit_Click(object sender, EventArgs e)
        {
            ListItemCollection vCollection = new ListItemCollection();
            vCollection.Clear();
            foreach (ListItem vItem in eFixedDriver_Edit.Items)
            {
                if (vItem.Selected)
                {
                    vCollection.Add(vItem);
                }
            }
            foreach (ListItem vSelectedItem in vCollection)
            {
                eFixedDriver_Edit.Items.Remove(vSelectedItem);
            }
        }

        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if ((eFixedDriver_Edit.Items.Count > 0) && (eWorkState_Edit.SelectedIndex >= 0))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vDepNo = eDepNo_Edit.Text.Trim();
                string vWorkState = eWorkState_Edit.SelectedValue.Trim();
                DateTime vBuDate = DateTime.Parse(eSheetDate_Edit.Text.Trim());
                string vBuDateStr = vBuDate.Year.ToString("D4") + "/" + vBuDate.ToString("MM/dd");
                string vDriverList = "";
                foreach (ListItem vItem in eFixedDriver_Edit.Items)
                {
                    vDriverList = (vDriverList == "") ? "'" + vItem.Value + "', " : vDriverList += "'" + vItem.Value + "', ";
                }
                if (vDriverList.Substring(vDriverList.Length - 2, 2) == ", ")
                {
                    vDriverList = vDriverList.Substring(0, vDriverList.Length - 2);
                }
                string vSQLStr = "update RunSheetA set WorkState = '" + vWorkState + "' " + Environment.NewLine +
                                 " where BuDate = '" + vBuDateStr + "' " + Environment.NewLine +
                                 "   and DepNo = '" + vDepNo + "' " + Environment.NewLine +
                                 "   and Driver in (" + vDriverList + ") ";
                int vModifyCount = PF.ExecSQL(vConnStr, vSQLStr);
                if (vModifyCount == eFixedDriver_Edit.Items.Count)
                {
                    eFixedDriver_Edit.Items.Clear();
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }
    }
}