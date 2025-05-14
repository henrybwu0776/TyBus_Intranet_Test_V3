using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;

namespace TyBus_Intranet_Test_V3
{
    public partial class ConsumablesInstore : Page
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
        private string vSheetModeA = "";
        private string vSheetModeB = "";

        //資料繫結用的變數
        //表頭
        public string vdSheetNo = "";
        public string vdSupNo = "";
        public string vdSupName = "";
        public string vdSheetNote = "";
        public string vdSheetStatus = "";
        public string vdSheetStatus_C = "";
        public string vdPayMode = "";
        public string vdPayMode_C = "";
        public string vdTaxType = "";
        public string vdTaxType_C = "";
        public string vdDepNo = "";
        public string vdDepName = "";
        public string vdAssignMan = "";
        public string vdAssignManName = "";
        public string vdRemarkA = "";
        public string vdBuMan = "";
        public string vdBuManName = "";
        public string vdModifyMan = "";
        public string vdModifyManName = "";
        public string vdBuDate = "";
        public string vdStatusDate = "";
        public string vdPayDate = "";
        public string vdModifyDate = "";
        public string vdAmount = "";
        public string vdTaxRate = "";
        public string vdTaxAMT = "";
        public string vdTotalAmount = "";
        //明細

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                DateTime vToday = DateTime.Today;
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

                    //開單日期
                    string vCaseDateURL = "InputDate.aspx?TextboxID=" + eBuDate_S_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_S_Search.Attributes["onClick"] = vCaseDateScript;

                    vCaseDateURL = "InputDate.aspx?TextboxID=" + eBuDate_E_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuDate_E_Search.Attributes["onClick"] = vCaseDateScript;

                    OpenListItems();
                    if (!IsPostBack)
                    {
                        vSheetModeA = (Session["SheetModeA"].ToString().Trim() != "") ? Session["SheetModeA"].ToString().Trim() : "LIST";
                        vSheetModeB = (Session["SheetModeB"].ToString().Trim() != "") ? Session["SheetModeB"].ToString().Trim() : "LIST";
                        SetButtonListA(vSheetModeA);
                        SetButtonListB(vSheetModeB);
                    }
                    else
                    {
                        OpenData();
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
        /// 取回下拉選單選項
        /// </summary>
        private void OpenListItems()
        {
            string vTempStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            //單據狀態
            if (ddlSheetStatus_Search.Items.Count > 0)
            {
                ddlSheetStatus_Search.Items.Clear();
            }
            using (SqlConnection connGetItems = new SqlConnection(vConnStr))
            {
                vTempStr = "SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '電腦課耗材進出單fmIASheetA      SheetStatus')";
                SqlCommand cmdGetItems = new SqlCommand(vTempStr, connGetItems);
                connGetItems.Open();
                SqlDataReader drGetItems = cmdGetItems.ExecuteReader();
                if (drGetItems.HasRows)
                {
                    while (drGetItems.Read())
                    {
                        ListItem iTemp = new ListItem();
                        iTemp.Value = drGetItems["ClassNo"].ToString().Trim();
                        iTemp.Text = drGetItems["ClassTxt"].ToString().Trim();
                        ddlSheetStatus_Search.Items.Add(iTemp);
                    }
                }
            }
            //廠牌
            if (ddlBrand_Search.Items.Count > 0)
            {
                ddlBrand_Search.Items.Clear();
            }
            using (SqlConnection connGetItems = new SqlConnection(vConnStr))
            {
                vTempStr = "SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '總務耗材管理    fmConsumables   Brand') ORDER BY ClassNo";
                SqlCommand cmdGetItems = new SqlCommand(vTempStr, connGetItems);
                connGetItems.Open();
                SqlDataReader drGetItems = cmdGetItems.ExecuteReader();
                if (drGetItems.HasRows)
                {
                    while (drGetItems.Read())
                    {
                        ListItem iTemp = new ListItem();
                        iTemp.Value = drGetItems["ClassNo"].ToString().Trim();
                        iTemp.Text = drGetItems["ClassTxt"].ToString().Trim();
                        ddlBrand_Search.Items.Add(iTemp);
                    }
                }
            }
            //付款方式
            if (ddlPayMode.Items.Count > 0)
            {
                ddlPayMode.Items.Clear();
            }
            using (SqlConnection connGetItems = new SqlConnection(vConnStr))
            {
                vTempStr = "SELECT NULL AS ClassNo, NULL AS ClassTxt UNION ALL SELECT CLASSNO, CLASSTXT FROM DBDICB WHERE (FKEY = '總務耗材進出單  fmConsSheetA    PayMode') ORDER BY ClassNo ";
                SqlCommand cmdGetItems = new SqlCommand(vTempStr, connGetItems);
                connGetItems.Open();
                SqlDataReader drGetItems = cmdGetItems.ExecuteReader();
                if (drGetItems.HasRows)
                {
                    while (drGetItems.Read())
                    {
                        ListItem iTemp = new ListItem();
                        iTemp.Value = drGetItems["ClassNo"].ToString().Trim();
                        iTemp.Text = drGetItems["ClassTxt"].ToString().Trim();
                        ddlPayMode.Items.Add(iTemp);
                    }
                }
            }
        }

        /// <summary>
        /// 取回查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            string vResultStr = "";

            String vBuDate_S = (eBuDate_S_Search.Text.Trim() != "") ? PF.TransDateString(eBuDate_S_Search.Text.Trim(), "B") : DateTime.Now.ToShortDateString();
            string vBuDate_E = (eBuDate_E_Search.Text.Trim() != "") ? PF.TransDateString(eBuDate_E_Search.Text.Trim(), "B") : DateTime.Now.ToShortDateString();

            string vWStr_SheetNo = ((eSheetNo_S_Search.Text.Trim() != "") && (eSheetNo_E_Search.Text.Trim() != "")) ? "   and a.SheetNo between '" + eSheetNo_S_Search.Text.Trim() + "' and '" + eSheetNo_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                   ((eSheetNo_S_Search.Text.Trim() != "") && (eSheetNo_E_Search.Text.Trim() == "")) ? "   and a.SheetNo = '" + eSheetNo_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                   ((eSheetNo_S_Search.Text.Trim() == "") && (eSheetNo_E_Search.Text.Trim() != "")) ? "   and a.SheetNo = '" + eSheetNo_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_SupNo = (eSupNo_Search.Text.Trim() != "") ? "   and a.SupNo = '" + eSupNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_SheetStatus = (eSheetStatus_Search.Text.Trim() != "") ? "   and a.SheetStatus = '" + eSheetStatus_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuMan = (eBuMan_Search.Text.Trim() != "") ? "   and a.BuMan = '" + eBuMan_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuDate = ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() != "")) ? "   and a.BuDate between '" + vBuDate_S + "' and '" + vBuDate_E + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() != "") && (eBuDate_E_Search.Text.Trim() == "")) ? "   and a.BuDate = '" + vBuDate_S + "' " + Environment.NewLine :
                                  ((eBuDate_S_Search.Text.Trim() == "") && (eBuDate_E_Search.Text.Trim() != "")) ? "   and a.BuDate = '" + vBuDate_E + "' " + Environment.NewLine : "";
            string vWStr_SheetNote = (eSheetNote_Search.Text.Trim() != "") ? "   and a.SheetNote like '%" + eSheetNote_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vWStr_ConsNo = (eConsNo_Search.Text.Trim() != "") ? "           and ConsNo = '" + eConsNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ? "           and ConsNo in (select ConsNo from Consumables where ConsName lke '%" + eConsName_Search.Text.Trim() + "%') " + Environment.NewLine : "";
            string vWStr_Brand = (eBrand_Search.Text.Trim() != "") ? "           and ConsNo in (select ConsNo from Consumables where Brand = '" + eBrand_Search.Text.Trim() + "') " + Environment.NewLine : "";

            string vWStr_Substr = ((eConsNo_Search.Text.Trim() == "") && (eConsName_Search.Text.Trim() == "") && (eBrand_Search.Text.Trim() == "")) ? "" :
                                  "   and a.SheetNo in " + Environment.NewLine +
                                  "       (select distinct SheetNo " + Environment.NewLine +
                                  "          from ConsSheetB " + Environment.NewLine +
                                  "         where isnull(SheetNo, '') <> '' " + Environment.NewLine +
                                  vWStr_ConsNo +
                                  vWStr_ConsName +
                                  vWStr_Brand +
                                  "       )" + Environment.NewLine;

            vResultStr = "SELECT SheetNo, SupNo, (SELECT name FROM CUSTOM WHERE (types = 'S') AND (code = a.SupNo)) AS SupName, BuDate, TotalAmount, SheetNote, SheetStatus, " + Environment.NewLine +
                         "       (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '總務耗材進出單  fmConsSheetA    SheetStatus') AND (CLASSNO = a.SheetStatus)) AS SheetStatus_C, RemarkA " + Environment.NewLine +
                         "  FROM ConsSheetA AS a " + Environment.NewLine +
                         " WHERE ISNULL(a.SheetNo, '') <> '' " + Environment.NewLine +
                         "   AND SheetMode = 'SI' " + Environment.NewLine +
                         vWStr_SheetNo +
                         vWStr_SupNo +
                         vWStr_SheetStatus +
                         vWStr_BuMan +
                         vWStr_BuDate +
                         vWStr_SheetNote +
                         vWStr_Substr +
                         " ORDER BY a.SheetNo DESC ";
            return vResultStr;
        }

        /// <summary>
        /// 取回查詢資料
        /// </summary>
        private void OpenData()
        {
            string vSelectStr = GetSelectStr();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(vSelectStr, connTemp);
                connTemp.Open();
                DataTable dtTemp = new DataTable();
                daTemp.Fill(dtTemp);
                gridSheetA.DataSource = dtTemp;
                gridSheetA.DataBind();
                plShowA.Visible = (dtTemp.Rows.Count > 0);
            }
        }

        /// <summary>
        /// 設定A檔按鈕列狀態
        /// </summary>
        /// <param name="SheetMode"></param>
        private void SetButtonListA(string SheetMode)
        {
            plButton_List.Visible = (SheetMode.ToUpper() == "LIST");
            plButton_Action.Visible = ((SheetMode.ToUpper() == "EDIT") || (SheetMode.ToUpper() == "INS"));
            plButton_ChangeMode.Visible = (SheetMode.ToUpper() == "LIST");
        }

        /// <summary>
        /// 設定B檔按鈕列狀態
        /// </summary>
        /// <param name="SheetMode"></param>
        private void SetButtonListB(string SheetMode)
        {

        }

        protected void eSupNo_Search_TextChanged(object sender, EventArgs e)
        {

        }

        protected void ddlSheetStatus_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eSheetStatus_Search.Text = ddlSheetStatus_Search.SelectedValue.ToString().Trim();
        }

        protected void eBuMan_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vBuMan = eBuMan_Search.Text;
            string vTempStr = "select [Name] from Employee where isnull(Leaveday, '') = '' and EmpNo = '" + vBuMan + "' ";
            string vBuManName = PF.GetValue(vConnStr, vTempStr, "Name");
            if (vBuManName == "")
            {
                vBuManName = vBuMan;
                vTempStr = "select top 1 EmpNo from Employee where [Name]] = '" + vBuManName + "' and isnull(Leaveday,'') = '' order by EmpNo DESC";
                vBuMan = PF.GetValue(vConnStr, vTempStr, "EmpNo");
            }
            eBuMan.Text = vBuMan;
            eBuManName.Text = vBuManName;
        }

        protected void ddlBrand_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            eBrand_Search.Text = ddlBrand_Search.SelectedValue.ToString().Trim();
        }

        protected void bbOK_Search_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
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

        protected void bbOK_Click(object sender, EventArgs e)
        {

        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {

        }

        protected void bbGetGoods_Click(object sender, EventArgs e)
        {

        }

        protected void bbIsPaid_Click(object sender, EventArgs e)
        {

        }

        protected void bbIsAbort_Click(object sender, EventArgs e)
        {

        }

        protected void bbCaseClose_Click(object sender, EventArgs e)
        {

        }

        protected void eBuMan_TextChanged(object sender, EventArgs e)
        {

        }

        protected void eSupNo_TextChanged(object sender, EventArgs e)
        {

        }

        protected void ddlPayMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            ePayMode.Text = ddlPayMode.SelectedValue.Trim();
        }
    }
}