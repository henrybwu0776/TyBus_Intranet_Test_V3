using System;
using System.Data;
using System.Data.SqlClient;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using Amaterasu_Function;

namespace TyBus_Intranet_Test_V3
{ 
    public partial class CreateIAAsset : System.Web.UI.Page
    {
        private string vLoginID = "";
        private string vPWD = "";
        private string vDepNo = "";
        private string vTitle = "";
        private string vEmpType = "";
        private string vConnStr = "";

        PublicFunction PF = new PublicFunction();

        protected void Page_Load(object sender, EventArgs e)
        {
            string StartURL = "InputDate.aspx?TextboxID=" + eBuyDate_Search1.ClientID;
            string EndURL = "InputDate.aspx?TextboxID=" + eBuyDate_Search2.ClientID;
            string BuyDateURL = "InputDate.aspx?TextboxID=" + eBuyDate.ClientID;
            string WarrantyURL = "InputDate.aspx?TextboxID=" + eWarrantyDate.ClientID;

            string StartScript = "window.open('" + StartURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
            string EndScript = "window.open('" + EndURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
            string BuyDateScript = "window.open('" + BuyDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
            string WarrantyScript = "window.open('" + WarrantyURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";

            eBuyDate_Search1.Attributes["onClick"] = StartScript;
            eBuyDate_Search2.Attributes["onClick"] = EndScript;
            eBuyDate.Attributes["onClick"] = BuyDateScript;
            eWarrantyDate.Attributes["onClick"] = WarrantyScript;

            DateTime vToday;
            vLoginID = (string)Session["LoginID"] ?? "";
            vDepNo = (string)Session["DepNo"] ?? "";
            if (vLoginID != "")
            {
                vToday = DateTime.Today;
                //CheckPermission(vLoginID);
                lbLoginID.Text = vLoginID;
                string vPath = Request.ApplicationPath;
                string connStr = PF.GetConnectionStr(vPath);
                string vSQLStr = "select Name from Employee where EmpNo ='" + vLoginID + "'";
                lbLoginName.Text = PF.GetValue(connStr, vSQLStr, "Name");
                GetAssetBrandList(eBrand);
                GetAssetBrandList(ddlBrandSearch);
                GetAssetProductUnitList();
                ClearSearch();

                if (IsPostBack)
                {

                }
            }
            else
            {
                Response.Redirect("~/Default.aspx");
            }
        }

        private void GetSystematics(DropDownList ddlTemp, string fAssetType)
        {
            string vPath = Request.ApplicationPath;
            string connTempStr = PF.GetConnectionStr(vPath);
            SqlConnection connTemp = new SqlConnection(connTempStr);
            string vSQLStr = "select ClassTxt, ClassNo from DBDICB where FKey = '電腦資產管理    fmIAAsset       Systematics' and Description = @Description order by ClassNo";
            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
            cmdTemp.Parameters.Add("@Description", SqlDbType.VarChar);
            cmdTemp.Parameters["@Description"].Value = fAssetType;
            connTemp.Open();
            SqlDataReader drTemp = cmdTemp.ExecuteReader();
            while (drTemp.Read())
            {
                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                ddlTemp.Items.Add(liTemp);
            }
            cmdTemp.Cancel();
            drTemp.Close();
            connTemp.Close();
        }

        private void GetAssetBrandList(DropDownList ddlTemp)
        {
            ddlTemp.Items.Clear();
            string vPath = Request.ApplicationPath;
            string connTempStr = PF.GetConnectionStr(vPath);
            string vSQLStr = "select ClassTxt, ClassNo from DBDICB where FKey = '電腦資產管理    fmIAAsset       Brand' order by ClassNo";
            SqlConnection connTemp = new SqlConnection(connTempStr);
            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
            connTemp.Open();
            SqlDataReader drTemp = cmdTemp.ExecuteReader();
            while (drTemp.Read())
            {
                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                ddlTemp.Items.Add(liTemp);
            }
            cmdTemp.Cancel();
            drTemp.Close();
            connTemp.Close();
        }

        private void GetAssetProductUnitList()
        {
            eProductUnit.Items.Clear();
            string vPath = Request.ApplicationPath;
            string connTempStr = PF.GetConnectionStr(vPath);
            string vSQLStr = "select ClassTxt, ClassNo from DBDICB where FKey = '電腦資產管理    fmIAAsset       ProductUnit' order by ClassNo";
            SqlConnection connTemp = new SqlConnection(connTempStr);
            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
            connTemp.Open();
            SqlDataReader drTemp = cmdTemp.ExecuteReader();
            while (drTemp.Read())
            {
                ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                eProductUnit.Items.Add(liTemp);
            }
            cmdTemp.Cancel();
            drTemp.Close();
            connTemp.Close();
        }

        private void ClearSearch()
        {
            cbAssetType.Checked = false;
            ddlAssetTypeSearch.SelectedIndex = 0;
            cbSystematics.Checked = false;
            ddlSystematicsSearch.SelectedIndex = 0;
            cbBrand.Checked = false;
            ddlBrandSearch.SelectedIndex = 0;
            eProductName_Search.Text = "";
            eAssetNo_Search1.Text = "";
            eAssetNo_Search2.Text = "";
            eBuyDate_Search1.Text = "";
            eBuyDate_Search2.Text = "";
            eSupplier_Search1.Text = "";
            eSupplier_Search2.Text = "";
            eOldAssetNo_Search1.Text = "";
            eOldAssetNo_Search2.Text = "";
        }

        private void SQLDataBaseInit(string fSQLStr)
        {
            dsAssetList.SelectCommand = fSQLStr;
        }

        private void AssetDataInit(string fAssetNo)
        {
            string vSQLStr = "select AssetNo, AssetType, Systematics, Brand, OriModeNumber, ProductName, OriSerialNumber, convert(varchar, BuyDate, 111) BuyDate, Warranty, Supplier, " +
                             "       convert(varchar, WarrantyDate, 111) WarrantyDate, OldAssetNo, ProductUnit, Quantity, NewPosition, OldPosition, InstallIn, Remark, ComputerName, " +
                             "       ComputerIPv4, ComputerIPv6, ComputerMACAdd, " +
                             "       (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       AssetType' and ClassNo = a.AssetType) AssetType_C, " +
                             "       (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       Systematics' and ClassNo = a.Systematics) Systematics_C, " +
                             "       (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       Brand' and ClassNo = a.Brand) Brand_C, " +
                             "       (select[Name] from[Custom] where Code = a.Supplier) Supplier_C " +
                             "  from IAAssetMain a " +
                             " where isnull(AssetNo,'') = @AssetNo";
            string vPath = Request.ApplicationPath;
            string connStr = PF.GetConnectionStr(vPath);

            SqlConnection connTemp = new SqlConnection(connStr);
            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
            cmdTemp.Parameters.Add("@AssetNo", SqlDbType.VarChar);
            cmdTemp.Parameters["@AssetNo"].Value = fAssetNo;
            connTemp.Open();
            SqlDataAdapter daTemp = new SqlDataAdapter(cmdTemp);
            DataSet dsTemp = new DataSet();
            daTemp.Fill(dsTemp, "tableTemp");
            DataTable tTemp = dsTemp.Tables["tableTemp"];

            //將資料顯示到畫面上
            eAssetNo.Text = tTemp.Rows[0]["AssetNo"].ToString().Trim();
            string vAssetType = tTemp.Rows[0]["AssetType"].ToString();
            eAssetType.SelectedValue = vAssetType;
            GetSystematics(eSystematics, vAssetType);
            string vSystematics = tTemp.Rows[0]["Systematics"].ToString().Trim();
            eSystematics.SelectedValue = vSystematics;
            GetAssetBrandList(eBrand);
            eBrand.SelectedValue = tTemp.Rows[0]["Brand"].ToString();
            eOriModeNumber.Text = tTemp.Rows[0]["OriModeNumber"].ToString();
            eOriSerialNumber.Text = tTemp.Rows[0]["OriSerialNumber"].ToString();
            eBuyDate.Text = tTemp.Rows[0]["BuyDate"].ToString();
            eWarranty.Text = tTemp.Rows[0]["Warranty"].ToString();
            eProductName.Text = tTemp.Rows[0]["ProductName"].ToString();
            eQuantity.Text = tTemp.Rows[0]["Quantity"].ToString();
            GetAssetProductUnitList();
            string vProductUnit = tTemp.Rows[0]["ProductUnit"].ToString().Trim();
            eProductUnit.SelectedValue = vProductUnit;
            eSupplier.Text = tTemp.Rows[0]["Supplier"].ToString();
            eSupplierName.Text = tTemp.Rows[0]["Supplier_C"].ToString();
            eWarrantyDate.Text = tTemp.Rows[0]["WarrantyDate"].ToString();
            eOldAssetNo.Text = tTemp.Rows[0]["OldAssetNo"].ToString();
            eComputerName.Text = tTemp.Rows[0]["ComputerName"].ToString();
            eComputerIPv4.Text = tTemp.Rows[0]["ComputerIPv4"].ToString();
            eComputerIPv6.Text = tTemp.Rows[0]["ComputerIPV6"].ToString();
            eComputerMACAdd.Text = tTemp.Rows[0]["ComputerMACAdd"].ToString();
            eNewPosition.Text = tTemp.Rows[0]["NewPosition"].ToString();
            eOldPosition.Text = tTemp.Rows[0]["OldPosition"].ToString();
            eInstallIn.Text = tTemp.Rows[0]["InstallIn"].ToString();
            eRemark.Text = tTemp.Rows[0]["Remark"].ToString();
            tTemp.Clear();
            cmdTemp.Cancel();
            connTemp.Close();
        }

        /*
        protected void CheckPermission(string fLoginID)
        {
            vPWD = "";
            vDepNo = "";
            vTitle = "";
            vEmpType = "";
            string vTempControlName = "";
            //宣告一個 SQLConnection 
            //到 webconfig 取回連線字串
            vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            using (SqlConnection ConGetEmpData = new SqlConnection(vConnStr))
            {
                //開啟 SQLConnection
                ConGetEmpData.Open();
                //產生 SQL 查詢字串
                string vSQLStr = "select EmpNo, DepNo, Title, Password, Type from Employee where EmpNo = @EmpNo and LeaveDay is null";
                //宣告一個 SQLCommand 
                SqlCommand cmdLogin = new SqlCommand(vSQLStr, ConGetEmpData);
                cmdLogin.Parameters.Add("@EmpNo", SqlDbType.VarChar);
                cmdLogin.Parameters["@EmpNo"].Value = fLoginID;
                //宣告一個 SQLDataReader
                SqlDataReader drLogin = cmdLogin.ExecuteReader();
                //實際讀取資料
                while (drLogin.Read())
                {
                    //取回登入人員的真實密碼
                    vPWD = drLogin["Password"].ToString();
                    //部門別
                    vDepNo = drLogin["DepNo"].ToString();
                    //職稱
                    vTitle = drLogin["Title"].ToString();
                    //員工類別 00--管理人員  10--修護人員  20--行車人員  30--其他人員
                    vEmpType = drLogin["Type"].ToString();
                    if (vEmpType != "20")
                    {
                        //不是行車人員才執行
                        //把登入者編號寫進 Session
                        Session["LoginID"] = fLoginID;
                        //回資料庫取回權限的設定
                        //先取回所有的功能並且設為隱藏
                        vSQLStr = "select ControlName from WebPermissionA order by OrderIndex";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                vTempControlName = drTemp["ControlName"].ToString();
                                HtmlGenericControl liTemp = (HtmlGenericControl)Master.FindControl(vTempControlName);
                                liTemp.Attributes.Remove("style");
                                liTemp.Attributes.Add("style", "display: none");
                            }
                            drTemp.Close();
                            cmdTemp.Cancel();
                            //connTemp.Close();
                            //先比對部門權限
                            vSQLStr = "select ControlName from WebPermissionB where DepNo = '" + vDepNo + "' and AllowPermission = 1";
                            cmdTemp.CommandText = vSQLStr;
                            //connTemp.Open();
                            drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                vTempControlName = drTemp["ControlName"].ToString();
                                HtmlGenericControl liTemp = (HtmlGenericControl)Master.FindControl(vTempControlName);
                                if (liTemp != null)
                                {
                                    liTemp.Attributes.Remove("style");
                                    liTemp.Attributes.Add("style", "display: normal");
                                }
                            }
                            drTemp.Close();
                            cmdTemp.Cancel();
                            //connTemp.Close();
                            //再比對個人有授予權限的
                            vSQLStr = "select ControlName from WebPermissionB where EmpNo = '" + vLoginID + "' and AllowPermission = 1";
                            cmdTemp.CommandText = vSQLStr;
                            //connTemp.Open();
                            drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                vTempControlName = drTemp["ControlName"].ToString();
                                HtmlGenericControl liTemp = (HtmlGenericControl)Master.FindControl(vTempControlName);
                                if (liTemp != null)
                                {
                                    liTemp.Attributes.Remove("style");
                                    liTemp.Attributes.Add("style", "display: normal");
                                }
                            }
                            drTemp.Close();
                            cmdTemp.Cancel();
                            //connTemp.Close();
                            //再比對個人沒有授予權限的
                            vSQLStr = "select ControlName from WebPermissionB where EmpNo = '" + vLoginID + "' and AllowPermission = 0";
                            cmdTemp.CommandText = vSQLStr;
                            //connTemp.Open();
                            drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                vTempControlName = drTemp["ControlName"].ToString();
                                HtmlGenericControl liTemp = (HtmlGenericControl)Master.FindControl(vTempControlName);
                                if (liTemp != null)
                                {
                                    liTemp.Attributes.Remove("style");
                                    liTemp.Attributes.Add("style", "display: none");
                                }
                            }
                            drTemp.Close();
                            cmdTemp.Cancel();
                        }
                    }
                    else
                    {
                        Session["LoginID"] = "";
                    }
                }
                cmdLogin.Cancel();
                drLogin.Close();
            }
        } //*/

        protected void ddlAssetTypeSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            ddlSystematicsSearch.Items.Clear();
            string vDescription = (ddlAssetTypeSearch.SelectedValue.ToString().Trim() != "") ? ddlAssetTypeSearch.SelectedValue : "H";
            GetSystematics(ddlSystematicsSearch, vDescription);
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            string vSearchStr = "", vWStr_AssetType = "", vWStr_Systematics = "", vWStr_Brand = "", vWStr_ProductName = "", vWStr_AssetNo = "", vWStr_BuyDate = "", vWStr_Supplier = "", vWStr_OldAssetNo = "";
            //類別
            vWStr_AssetType = (cbAssetType.Checked) ? " and AssetType = '" + ddlAssetTypeSearch.SelectedValue.ToString() + "' " : "";
            //品項
            vWStr_Systematics = (cbSystematics.Checked) ? " and Systemaitcs = '" + ddlSystematicsSearch.SelectedValue.ToString() + "' " : "";
            //廠牌
            vWStr_Brand = (cbBrand.Checked) ? " and Brand = '" + ddlBrandSearch.SelectedValue.ToString() + "' " : "";
            //品名 (模糊比對)
            vWStr_ProductName = (eProductName_Search.Text.Trim() != "") ? " and ProductName like '%" + eProductName_Search.Text.Trim() + "%'" : "";
            //資產編號
            if ((eAssetNo_Search1.Text.Trim() != "") && (eAssetNo_Search2.Text.Trim() != ""))
            {
                vWStr_AssetNo = " and (AssetNo >= '" + eAssetNo_Search1.Text.Trim() + "' and AssetNo <= '" + eAssetNo_Search2.Text.Trim() + "' )";
            }
            else if ((eAssetNo_Search1.Text.Trim() != "") && (eAssetNo_Search2.Text.Trim() == ""))
            {
                vWStr_AssetNo = " and AssetNo like '" + eAssetNo_Search1.Text.Trim() + "%' ";
            }
            else if ((eAssetNo_Search1.Text.Trim() == "") && (eAssetNo_Search2.Text.Trim() != ""))
            {
                vWStr_AssetNo = " and AssetNo like '" + eAssetNo_Search2.Text.Trim() + "%' ";
            }
            //購置日期
            if ((eBuyDate_Search1.Text.Trim() != "") && (eBuyDate_Search2.Text.Trim() != ""))
            {
                vWStr_BuyDate = " and (BuyDate >= '" + eBuyDate_Search1.Text.Trim() + "' and BuyDate <= '" + eBuyDate_Search2.Text.Trim() + "' )";
            }
            else if ((eBuyDate_Search1.Text.Trim() != "") && (eBuyDate_Search2.Text.Trim() == ""))
            {
                vWStr_BuyDate = " and BuyDate like '" + eBuyDate_Search1.Text.Trim() + "%' ";
            }
            else if ((eBuyDate_Search1.Text.Trim() == "") && (eBuyDate_Search2.Text.Trim() != ""))
            {
                vWStr_BuyDate = " and BuyDate like '" + eBuyDate_Search2.Text.Trim() + "%' ";
            }
            //供應商
            if ((eSupplier_Search1.Text.Trim() != "") && (eSupplier_Search2.Text.Trim() != ""))
            {
                vWStr_Supplier = " and (Supplier >= '" + eSupplier_Search1.Text.Trim() + "' and Supplier <= '" + eSupplier_Search2.Text.Trim() + "' )";
            }
            else if ((eSupplier_Search1.Text.Trim() != "") && (eSupplier_Search2.Text.Trim() == ""))
            {
                vWStr_Supplier = " and Supplier like '" + eSupplier_Search1.Text.Trim() + "%' ";
            }
            else if ((eSupplier_Search1.Text.Trim() == "") && (eSupplier_Search2.Text.Trim() != ""))
            {
                vWStr_Supplier = " and Supplier like '" + eSupplier_Search2.Text.Trim() + "%' ";
            }
            //舊編號
            if ((eOldAssetNo_Search1.Text.Trim() != "") && (eOldAssetNo_Search2.Text.Trim() != ""))
            {
                vWStr_OldAssetNo = " and (OldAssetNo >= '" + eOldAssetNo_Search1.Text.Trim() + "' and OldAssetNo <= '" + eOldAssetNo_Search2.Text.Trim() + "' )";
            }
            else if ((eOldAssetNo_Search1.Text.Trim() != "") && (eOldAssetNo_Search2.Text.Trim() == ""))
            {
                vWStr_OldAssetNo = " and OldAssetNo like '" + eAssetNo_Search1.Text.Trim() + "%' ";
            }
            else if ((eOldAssetNo_Search1.Text.Trim() == "") && (eOldAssetNo_Search2.Text.Trim() != ""))
            {
                vWStr_OldAssetNo = " and OldAssetNo like '" + eOldAssetNo_Search2.Text.Trim() + "%' ";
            }

            vSearchStr = "select AssetNo, AssetType, Systematics, Brand, OriModeNumber, ProductName, OriSerialNumber, convert(varchar, BuyDate, 111) BuyDate, Warranty, Supplier, " +
                         "       convert(varchar, WarrantyDate, 111) WarrantyDate, OldAssetNo, ProductUnit, Quantity, NewPosition, OldPosition, InstallIn, Remark, ComputerName, " +
                         "       ComputerIPv4, ComputerIPv6, ComputerMACAdd, " +
                         "       (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       AssetType' and ClassNo = a.AssetType) AssetType_C, " +
                         "       (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       Systematics' and ClassNo = a.Systematics) Systematics_C, " +
                         "       (select ClassTxt from DBDICB where FKey = '電腦資產管理    fmIAAsset       Brand' and ClassNo = a.Brand) Brand_C, " +
                         "       (select[Name] from[Custom] where Code = a.Supplier) Supplier_C " +
                         "  from IAAssetMain a " +
                         " where isnull(AssetNo,'') <> ''" +
                         vWStr_AssetType + vWStr_Systematics + vWStr_Brand + vWStr_ProductName + vWStr_AssetNo + vWStr_BuyDate + vWStr_Supplier + vWStr_OldAssetNo;
            lbSQLStr.Text = vSearchStr;
            SQLDataBaseInit(lbSQLStr.Text.Trim());
        }

        protected void bbClearSearch_Click(object sender, EventArgs e)
        {
            ClearSearch();
        }

        protected void gvAssetData_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vAssetNo = gvAssetList.SelectedDataKey.Value.ToString();
            if (vAssetNo != "")
            {
                AssetDataInit(vAssetNo);
            }
        }

        protected void eAssetType_SelectedIndexChanged(object sender, EventArgs e)
        {
            eSystematics.Items.Clear();
            string vPath = Request.ApplicationPath;
            string connGetDataStr = PF.GetConnectionStr(vPath);
            string vDescription = (eAssetType.SelectedValue.ToString().Trim() != "") ? eAssetType.SelectedValue : "H";
            GetSystematics(eSystematics, vDescription);
        }

        protected void bbCreate_Click(object sender, EventArgs e)
        {
            eAssetNo.Text = "";
            eAssetNo.Enabled = false;
            eAssetType.Enabled = true;
            eSystematics.Enabled = true;
            eAssetType.Focus();
            //把新增修改和刪除的按鈕 Disable 掉
            bbCreate.Enabled = false;
            bbModify.Enabled = false;
            bbDel.Enabled = false;
            //打開確定和取消的按鈕
            bbOK.Enabled = true;
            bbCancel.Enabled = true;
            Session["EditMode"] = "Insert";
        }

        protected void bbModify_Click(object sender, EventArgs e)
        {
            if (eAssetNo.Text.Trim() == "")
            {
                Response.Write("<script>alert('必須先挑選出要修改的資產才能進行修改！');</script>");
            }
            else
            {
                //資產的類別和品項不可修改
                eAssetType.Enabled = false;
                eSystematics.Enabled = false;
                //游標定位到原廠型號
                eOriModeNumber.Focus();
                //把新增修改和刪除的按鈕 Disable 掉
                bbCreate.Enabled = false;
                bbModify.Enabled = false;
                bbDel.Enabled = false;
                //打開確定和取消的按鈕
                bbOK.Enabled = true;
                bbCancel.Enabled = true;
                Session["EditMode"] = "Modify";
            }
        }

        protected void bbDel_Click(object sender, EventArgs e)
        {
            string vAssetNo = eAssetNo.Text.Trim();
            if (vAssetNo != "")
            {
                string vSQLStr = "delete IAAssetMain where AssetNo = '" + vAssetNo + "'";
                string vPath = Request.ApplicationPath;
                PF.ExecSQL(PF.GetConnectionStr(vPath), vSQLStr);
                SQLDataBaseInit(lbSQLStr.Text.Trim());
                gvAssetList.SelectedIndex = 0;
                //vAssetNo = gvAssetList.SelectedDataKey.Value.ToString();
                //AssetDataInit(vAssetNo);
            }
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            string vSQLStr = "";
            string vAssetNo = "";
            string vAssetType = eAssetType.SelectedValue;
            string vSystematics = eSystematics.SelectedValue;
            string vBrand = eBrand.SelectedValue;
            DateTime vBuyDate = DateTime.Parse(eBuyDate.Text.Trim());
            string vOriModeNumber = eOriModeNumber.Text.Trim();
            string vProductName = eProductName.Text.Trim();
            string vOriSerialNumber = eOriSerialNumber.Text.Trim();
            int vWarranty = int.Parse(eWarranty.Text.Trim());
            string vSupplier = eSupplier.Text.Trim();
            DateTime vWarrantyDate = DateTime.Parse(eWarrantyDate.Text.Trim());
            string vOldAssetNo = eOldAssetNo.Text.Trim();
            string vProductUnit = eProductUnit.Text.Trim();
            float vQuantity = float.Parse(eQuantity.Text.Trim());
            string vNewPosition = eNewPosition.Text.Trim();
            string vOldPosition = eOldPosition.Text.Trim();
            string vInstallIn = eInstallIn.Text.Trim();
            string vRemark = eRemark.Text.Trim();
            string vPath = Request.ApplicationPath;

            if (Session["EditMode"].ToString() == "Insert")
            {
                //新增資料
                string vGetSerialStr = "select top 1 AssetNo from IAAssetMain where AssetType = '" + vAssetType + "' and Systematics = '" + vSystematics + "' order by AssetNo DESC";
                string vTempassetNo = PF.GetValue(PF.GetConnectionStr(vPath), vGetSerialStr, "AssetNo");
                if (vTempassetNo == "")
                {
                    vAssetNo = vAssetType.Trim() + vSystematics.Trim() + "0001"; //如果沒有找到相符的資料表示這個類別/品項是第一筆資料，就直接給流水號 "0001"
                }
                else
                {
                    vAssetNo = vAssetType.Trim() + vSystematics.Trim() + string.Format("{0:D4}", (int.Parse(vTempassetNo.Substring(vTempassetNo.Length - 4)) + 1));
                }
                vSQLStr = "insert into IAAssetMain (AssetNo, Brand, BuyDate, OriModeNumber, OriSerialNumber, ProductName, Warranty, Supplier, WarrantyDate, OldAssetNo, Quantity, " +
                                                   "ProductUnit, NewPosition, OldPosition, InstallIn, Remark, AssetType, Systematics)" +
                          "values ('" + vAssetNo + "', '" + vBrand + "', '" + vBuyDate.ToString("yyyy/MM/dd") + "', '" + vOriModeNumber + "', '" + vOriSerialNumber + "', '" + vProductName + "', " + vWarranty.ToString() +
                                ", '" + vSupplier + "', '" + vWarrantyDate.ToString("yyyy/MM/dd") + "', '" + vOldAssetNo + "', " + vQuantity.ToString() + ", '" + vProductUnit + "', '" + vNewPosition + "', '" + vOldPosition +
                                "', '" + vInstallIn + "', '" + vRemark + "', '" + vAssetType + "', '" + vSystematics + "')";
            }
            else if (Session["EditMode"].ToString() == "Modify")
            {
                //修改資料
                vAssetNo = eAssetNo.Text.Trim();
                vSQLStr = "update IAAssetMain set Brand = '" + vBrand + "', BuyDate = '" + vBuyDate.ToString("yyyy/MM/dd") + "', OriModeNumber = '" + vOriModeNumber + "', OriSerialNumber = '" + vOriSerialNumber +
                                              "', ProductName = '" + vProductName + "', Warranty = " + vWarranty.ToString() + ", Supplier = '" + vSupplier + "', WarrantyDate = '" + vWarrantyDate.ToString("yyyy/MM/dd") +
                                              "', OldAssetNo = '" + vOldAssetNo + "', Quantity = " + vQuantity.ToString() + ", ProductUnit = '" + vProductUnit + "', NewPosition = '" + vNewPosition +
                                              "', OldPosition = '" + vOldPosition + "', InstallIn = '" + vInstallIn + "', Remark = '" + vRemark + "' " +
                          " where AssetNo = '" + vAssetNo + "' ";
            }
            try
            {
                PF.ExecSQL(PF.GetConnectionStr(vPath), vSQLStr);
            }
            catch (Exception)
            {
                Response.Write("<script>alert('" + vSQLStr + "');</script>");
                throw;
            }

            ///打開新增修改和刪除的按鈕
            bbCreate.Enabled = true;
            bbModify.Enabled = true;
            bbDel.Enabled = true;
            //把確定和取消的按鈕 Disable 掉
            bbOK.Enabled = false;
            bbCancel.Enabled = false;
            Session["EditMode"] = "";
            SQLDataBaseInit(lbSQLStr.Text.Trim());
            gvAssetList.SelectedIndex = 0;
        }
    }
}