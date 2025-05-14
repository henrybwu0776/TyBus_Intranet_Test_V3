using Amaterasu_Function;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ConsBrand : Page
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
        //private DataTable dtShowData;
        private string vDataMode = "";

        //資料繫結用
        public string vdBrandCode;
        public string vdBrandName;
        public string vdBelongGroup;
        public string vdRemark;
        public string vdBuMan;
        public string vdBuMan_C;
        public string vdBuDate;
        public string vdModifyMan;
        public string vdModifyMan_C;
        public string vdModifyDate;

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
                vDataMode = (Session["BrandDataMode"] != null) ? Session["BrandDataMode"].ToString().Trim() : "";

                if (vLoginID != "")
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    OpenData();
                    if (!IsPostBack)
                    {
                        Session["BrandDataMode"] = "LIST";
                        vDataMode = Session["BrandDataMode"].ToString().Trim();
                        SetPanelStatus(vDataMode);
                        lbErrorMessage.Text = "";
                        lbErrorMessage.Visible = false;
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

        /// <summary>
        /// 開啟資料
        /// </summary>
        private void OpenData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vWStr_BrandCode = (eBrandCode_Search.Text.Trim() != "") ? "   and BrandCode like '%" + eBrandCode_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_BrandName = (eBrandName_Search.Text.Trim() != "") ? "   and BrandName like '%" + eBrandName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_BelongGroup = (ddlBelongGroup_Search.SelectedValue.Trim() != "") ? "   and BelongGroup like '%" + ddlBelongGroup_Search.SelectedValue.Trim() + "%' " + Environment.NewLine : "";
            string vSelectStr = "select BrandCode, BrandName, BelongGroup, Remark " + Environment.NewLine +
                                "  from ConsBrand " + Environment.NewLine +
                                " where isnull(BrandCode, '') != '' " + Environment.NewLine +
                                vWStr_BrandCode + vWStr_BrandName + vWStr_BelongGroup +
                                " order by BrandCode ";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelectStr, connTemp);
                connTemp.Open();
                SqlDataAdapter daTemp = new SqlDataAdapter(cmdTemp);
                DataTable dtTemp = new DataTable();
                daTemp.Fill(dtTemp);
                gridBrandList.DataSource = dtTemp;
                gridBrandList.DataBind();
            }
        }

        private void SetPanelStatus(string fDataMode)
        {
            plButton_List.Visible = (fDataMode.ToUpper() == "LIST");
            plButton_Action.Visible = !(fDataMode.ToUpper() == "LIST");
            eBrandCode.Enabled = (fDataMode.ToUpper() == "INS");
            eBrandName.Enabled = ((fDataMode.ToUpper() == "INS") || (fDataMode.ToUpper() == "EDIT"));
            eBelongGroup.Enabled = ((fDataMode.ToUpper() == "INS") || (fDataMode.ToUpper() == "EDIT"));
            eRemark.Enabled = ((fDataMode.ToUpper() == "INS") || (fDataMode.ToUpper() == "EDIT"));
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbInsert_Click(object sender, EventArgs e)
        {
            Session["BrandDataMode"] = "INS";
            vDataMode = Session["BrandDataMode"].ToString().Trim();
            SetPanelStatus(vDataMode);
            eBrandCode.Focus();
        }

        protected void bbEdit_Click(object sender, EventArgs e)
        {
            Session["BrandDataMode"] = "EDIT";
            vDataMode = Session["BrandDataMode"].ToString().Trim();
            SetPanelStatus(vDataMode);
            eBrandName.Focus();
        }

        protected void bbDelete_Click(object sender, EventArgs e)
        {
            lbErrorMessage.Text = "";
            lbErrorMessage.Visible = false;
            if (eBrandCode.Text.Trim() != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    string vTempStr = "delete ConsBrand where BrandCode = '" + eBrandCode.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vTempStr);
                    OpenData();
                }
                catch (Exception eMessage)
                {
                    lbErrorMessage.Text = eMessage.Message.Trim();
                    lbErrorMessage.Visible = true;
                }
            }
            else
            {
                lbErrorMessage.Text = "請選擇廠商！";
                lbErrorMessage.Visible = true;
            }
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            lbErrorMessage.Visible = false;
            lbErrorMessage.Text = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vTempStr = "";
            string vBrandCode_Temp = eBrandCode.Text.Trim();
            if (!string.IsNullOrEmpty(vBrandCode_Temp))
            {
                string vBrandName_Temp = eBrandName.Text.Trim();
                string vBelongGroup_Temp = eBelongGroup.Text.Trim();
                string vRemark_Temp = eRemark.Text.Trim();
                string vBuMan_Temp = vLoginID;
                string vModifyMan_Temp = vLoginID;
                try
                {
                    switch (vDataMode)
                    {
                        case "INS":
                            vTempStr = "insert into ConsBrand (BrandCode, BrandName, BelongGroup, Remark, BuMan, BuDate)" + Environment.NewLine +
                                       "values (@BrandCode, @BrandName, @BelongGroup, @Remark, @BuMan, GetDate())";
                            break;

                        case "EDIT":
                            vTempStr = "update ConsBrand set BrandName = @BrandName, BelongGroup = @BelongGroup, Remark = @Remark, " + Environment.NewLine +
                                       "                     ModifyMan = @ModifyMan, ModifyDate = GetDate() " + Environment.NewLine +
                                       " where BrandCode = @BrandCode";
                            break;
                    }
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                        connTemp.Open();
                        cmdTemp.Parameters.Clear();
                        cmdTemp.Parameters.Add(new SqlParameter("BrandCode", vBrandCode_Temp));
                        cmdTemp.Parameters.Add(new SqlParameter("BrandName", (vBrandName_Temp != "") ? vBrandName_Temp : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("BelongGroup", (vBelongGroup_Temp != "") ? vBelongGroup_Temp : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("Remark", (vRemark_Temp != "") ? vRemark_Temp : String.Empty));
                        if (vDataMode == "INS")
                            cmdTemp.Parameters.Add(new SqlParameter("BuMan", vBuMan_Temp));
                        else if (vDataMode == "EDIT")
                            cmdTemp.Parameters.Add(new SqlParameter("ModifyMan", vModifyMan_Temp));
                        cmdTemp.ExecuteNonQuery();
                        gridBrandList.DataBind();
                    }
                }
                catch (Exception eMessage)
                {
                    lbErrorMessage.Text = eMessage.ToString();
                    lbErrorMessage.Visible = true;
                }
                Session["BrandDataMode"] = "LIST";
                vDataMode = Session["BrandDataMode"].ToString().Trim();
                SetPanelStatus(vDataMode);
            }
            else
            {
                lbErrorMessage.Visible = true;
                lbErrorMessage.Text = "請選擇廠商！";
            }
        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            Session["BrandDataMode"] = "LIST";
            vDataMode = Session["BrandDataMode"].ToString().Trim();
            SetPanelStatus(vDataMode);
        }

        protected void gridBrandList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            lbErrorMessage.Visible = false;
            lbErrorMessage.Text = "";
            string vTempStr = "";
            string vBrandCodeTemp = gridBrandList.SelectedDataKey.ToString().Trim();
            if (!string.IsNullOrEmpty(vBrandCodeTemp))
            {
                vTempStr = "select a.BrandCode, a.BrandName, a.BelongGroup, a.Remark, " + Environment.NewLine +
                           "       convert(nvarchar(10),a.BuDate, 111) BuDate, a.BuMan, (select [Name] from Employee where EmpNo = a.BuMan) BuMan_C, " + Environment.NewLine +
                           "       convert(nvarchar(10),a.ModifyDate, 111) ModifyDate, a.ModifyMan, (select [Name] from Employee where EmpNo = a.ModifyMan) ModifyMan_C " + Environment.NewLine +
                           "  from ConsBrand a " + Environment.NewLine +
                           " where BrandCode = @BrandCode";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                    cmdTemp.Parameters.Clear();
                    cmdTemp.Parameters.Add(new Parameter("BrandCode", DbType.String, vBrandCodeTemp));
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    if (drTemp.HasRows)
                    {
                        while (drTemp.Read())
                        {
                            vdBrandCode = drTemp["BrandCode"].ToString().Trim();
                            vdBrandName = drTemp["BrandName"].ToString().Trim();
                            vdBelongGroup = drTemp["BelongGroup"].ToString().Trim();
                            vdRemark = drTemp["Remark"].ToString().Trim();
                            vdBuMan = drTemp["BuMan"].ToString().Trim();
                            vdBuMan_C = drTemp["BuMan_C"].ToString().Trim();
                            vdBuDate = drTemp["BuDate"].ToString().Trim();
                            vdModifyMan = drTemp["ModifyMan"].ToString().Trim();
                            vdModifyMan_C = drTemp["ModifyMan_C"].ToString().Trim();
                            vdModifyDate = drTemp["ModifyDate"].ToString().Trim();
                        }
                    }
                }
            }
            else
            {
                lbErrorMessage.Visible = true;
                lbErrorMessage.Text = "請選擇廠商！";
            }
            Page.DataBind();      
        }

        protected void eBrandCode_TextChanged(object sender, EventArgs e)
        {
            lbErrorMessage.Visible = false;
            lbErrorMessage.Text = "";
            string vBrandCode_Source = eBrandCode.Text.Trim();
            if (!string.IsNullOrEmpty(vBrandCode_Source))
            {
                string vTempStr = "select BrandName from ConsBrand where BrandCode = '" + vBrandCode_Source + "' ";
                string vBrandName = PF.GetValue(vConnStr, vTempStr, "BrandName");
                if (!string.IsNullOrEmpty(vBrandName))
                {
                    lbErrorMessage.Text = "系統已有相同廠商代碼：【" + vBrandName + "】";
                    lbErrorMessage.Visible = true;
                    eBrandCode.Text = "";
                    eBrandCode.Focus();
                }
            }
            else
            {
                lbErrorMessage.Visible = true;
                lbErrorMessage.Text = "廠商代碼不可空白！";
                eBrandCode.Focus();
            }
        }
    }
}