using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;
using System.Data.SqlClient;
using System.IO;
using System.Data;

namespace TyBus_Intranet_Test_V3
{
    public partial class Consumables : Page
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
        private DataTable dtShowData;

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
                        plReport.Visible = false;
                    }
                    else
                    {

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
        /// 取回下拉選單的內容
        /// </summary>
        private void GetItemListData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            //取回耗材分類
            string vSelectStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                     " union all" + Environment.NewLine +
                                     "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsType' order by ClassNo";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelectStr_Temp, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    ddlConsType_Search.Items.Clear();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                        ddlConsType_Search.Items.Add(liTemp);
                    }
                }
            }
            //取回耗材單位
            vSelectStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                              " union all" + Environment.NewLine +
                              "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsUnit' order by ClassNo";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSelectStr_Temp, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {

                }
            }
        }

        protected void eConsNo_Search_TextChanged(object sender, EventArgs e)
        {
            lbErrorMSG_ConsName.Text = "";
            lbErrorMSG_ConsName.Visible = false;
            string vTempStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vConsNo = eConsNo_Search.Text.Trim();
            vTempStr = "select ConsName from Consumables where ConsNo = '" + vConsNo + "' ";
            string vConsName = PF.GetValue(vConnStr, vTempStr, "ConsName");
            if (vConsName == "")
            {
                eConsNo_Search.Text = "";
                lbErrorMSG_ConsName.Text = "查無 [ " + vConsNo + " ] 料號";
                lbErrorMSG_ConsName.Visible = true;
            }
        }

        private string OpenListStr()
        {
            string vResultStr = "";
            string vWStr_ConsNo = (eConsNo_Search.Text.Trim() != "") ? "   and ConsNo = '" + eConsNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "   and ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_ConsType = (ddlConsType_Search.SelectedValue.Trim() != "") ? "   and ConsType = '" + ddlConsType_Search.SelectedValue.Trim() + "' " + Environment.NewLine : "";
            vResultStr = "select c.ConsNo, c.ConsName, d1.ClassTxt ConsType, c.StockQty, d2.ClassTxt ConsUnit, c.AvgPrice, c.StoreLocation, c.Brand " + Environment.NewLine +
                         "  from Consumables c left join DBDICB d1 on d1.ClassNo = c.ConsType and d1.FKey = '耗材庫存        CONSUMABLES     ConsType' " + Environment.NewLine +
                         "                     left join DBDICB d2 on d2.ClassNo = c.ConsUnit and d2.FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                         " where isnull(c.ConsNo, '') != '' " + Environment.NewLine +
                         vWStr_ConsNo +
                         vWStr_ConsName +
                         vWStr_ConsType +
                         " order by c.ConsNo";
            return vResultStr;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vListStr = OpenListStr();
            using (SqlConnection connShowList = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daShowList = new SqlDataAdapter(vListStr, connShowList);
                connShowList.Open();
                if (dtShowData != null)
                {
                    dtShowData.Clear();
                }
                else
                {
                    dtShowData = new DataTable();
                }
                daShowList.Fill(dtShowData);
            }
            gvShowList.DataSourceID = "";
            gvShowList.DataSource = dtShowData;
            gvShowList.DataBind();
        }

        protected void bbPrintCheckList_Click(object sender, EventArgs e)
        {

        }

        protected void bbUpdateReserve_Search_Click(object sender, EventArgs e)
        {

        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plReport.Visible = false;
            plShowData.Visible = true;
        }

        protected void gvShowList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gvShowList.PageIndex = e.NewPageIndex;
            gvShowList.DataBind();
        }

        protected void bbInsert_Click(object sender, EventArgs e)
        {

        }

        protected void bbEdit_Click(object sender, EventArgs e)
        {

        }

        protected void bbDelete_Click(object sender, EventArgs e)
        {

        }

        protected void bbStopUse_Click(object sender, EventArgs e)
        {

        }

        protected void gvShowList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}