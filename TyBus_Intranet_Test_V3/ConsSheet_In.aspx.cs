using Amaterasu_Function;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ConsSheet_In : Page
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
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                System.Threading.Thread.CurrentThread.CurrentCulture = Cal;

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
                    //動態掛載日期輸入視窗
                    string vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDateS_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBuDateS_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDateE_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBuDateE_Search.Attributes["onClick"] = vDateScript;

                    if (!IsPostBack)
                    {
                        plSearch.Visible = true;
                        plShowDataA.Visible = true;
                        plShowDataB.Visible = true;
                        plPickupDetail.Visible = false;
                        //plPrint.Visible = false;
                    }
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    OpenData();
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

        private string GetSelectStr()
        {
            DateTime vTempDateS;
            DateTime vTempDateE;
            string vResultStr;
            string vWStr_BuDate = (DateTime.TryParse(eBuDateS_Search.Text.Trim(), out vTempDateS) && DateTime.TryParse(eBuDateE_Search.Text.Trim(), out vTempDateE)) ?
                                  "   and a.BuDate between '" + PF.GetAD(vTempDateS.ToShortDateString()) + "' and '" + PF.GetAD(vTempDateE.ToShortDateString()) + "' " + Environment.NewLine :
                                  (DateTime.TryParse(eBuDateS_Search.Text.Trim(), out vTempDateS) && !DateTime.TryParse(eBuDateE_Search.Text.Trim(), out vTempDateE)) ?
                                  "   and a.BuDate = '" + PF.GetAD(vTempDateS.ToShortDateString()) + "' " + Environment.NewLine :
                                  (!DateTime.TryParse(eBuDateS_Search.Text.Trim(), out vTempDateS) && DateTime.TryParse(eBuDateE_Search.Text.Trim(), out vTempDateE)) ?
                                  "   and a.BuDate = '" + PF.GetAD(vTempDateE.ToShortDateString()) + "' " + Environment.NewLine : "";
            string vWStr_SupName = (eSupNo_Search.Text.Trim() != "") ? "   and c.[Name] like '%" + eSupNo_Search + "%' " + Environment.NewLine : "";
            string vWStr_ConsName = (eConsName_Search.Text.Trim() != "") ?
                                    "   and a.SheetNo in (select SheetNo from ConsSheetB where ConsNo in (select ConsNo from Consumables where ConsName like '%" + eConsName_Search.Text.Trim() + "%' " + Environment.NewLine :
                                    "";
            vResultStr = "select a.SheetNo, a.BuDate, e.[Name] BuMan_C, c.[Name] SupName, a.TotalAmount, d.ClassTxt SheetStatus_C " + Environment.NewLine +
                         "  from ConsSheetA a left join [Custom] c on c.Code = a.SupNo and c.[Types] = 'S' " + Environment.NewLine +
                         "                    left join DBDICB d on d.ClassNo = a.SheetStatus and d.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus' " + Environment.NewLine +
                         "                    left join Employee e on e.EmpNo = a.BuMan " + Environment.NewLine +
                         " where a.SheetMode = 'IS' " + Environment.NewLine +
                         vWStr_BuDate +
                         vWStr_ConsName +
                         vWStr_SupName +
                         " order by a.SheetNo ";
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSelectStr();
            sdsConsSheetA_List.SelectCommand = vSelectStr;
            sdsConsSheetA_List.Select(new DataSourceSelectArguments());
            gridConsSheetA_List.DataBind();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvConsSheetA_Detail_DataBound(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            string vSQLStr_Temp;
            string vSupDataURL;
            string vSupDataScript;
            Label eSheetNo;
            TextBox eSupNo;
            TextBox eSupName;
            switch (fvConsSheetA_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
                    if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
                    {
                        string vSheetNo = eSheetNo.Text.Trim();
                        //只要有不是 "已開單" 或 "已作廢" 的明細就不開放整單作廢
                        vSQLStr_Temp = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '000') not in ('000', '999') ";
                        string vRCount_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount");
                        int vRCount = Int32.Parse(vRCount_Str);
                        Button bbAbortA = (Button)fvConsSheetA_Detail.FindControl("bbAbortA_List");
                        bbAbortA.Enabled = (vRCount <= 0);
                        //如果沒有 "已開單" 的明細就不開放 "整單入庫"
                        vSQLStr_Temp = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' and isnull(ItemStatus, '000') = '000' ";
                        vRCount_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount");
                        vRCount = Int32.Parse(vRCount_Str);
                        Button bbInstore = (Button)fvConsSheetA_Detail.FindControl("bbInStoreA_List");
                        bbInstore.Enabled = (vRCount > 0);
                        //如果單據已作廢，整個明細都不能操作
                        Label eSheetStatus = (Label)fvConsSheetA_Detail.FindControl("eSheetStatus_List");
                        plShowDataB.Enabled = !((eSheetStatus.Text.Trim() == "999") || (eSheetStatus.Text.Trim() == "998"));
                        //單據已完成入庫，就不允許整單作廢、入庫及刪除
                        if (eSheetStatus.Text.Trim()=="030")
                        {
                            bbAbortA.Enabled = false;
                            Button bbDeleteA = (Button)fvConsSheetA_Detail.FindControl("bbDeleteA_List");
                            bbDeleteA.Enabled = false;
                            bbInstore.Enabled = false;
                            Button bbEditA = (Button)fvConsSheetA_Detail.FindControl("bbEditA_List");
                            bbEditA.Enabled = false;
                        }
                    }
                    break;
                case FormViewMode.Edit:
                    eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_Edit");
                    if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
                    {
                        eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_Edit");
                        eSupName = (TextBox)fvConsSheetA_Detail.FindControl("eSupName_Edit");
                        vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo.ClientID + "&SupNameID=" + eSupName.ClientID + "&SupKind=C";
                        vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eSupNo.Attributes["onClick"] = vSupDataScript;
                    }
                    break;
                case FormViewMode.Insert:
                    eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_INS");
                    if (eSheetNo != null)
                    {
                        eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_INS");
                        eSupName = (TextBox)fvConsSheetA_Detail.FindControl("eSupName_INS");
                        vSupDataURL = "SearchSup.aspx?TextBoxID=" + eSupNo.ClientID + "&SupNameID=" + eSupName.ClientID + "&SupKind=C";
                        vSupDataScript = "window.open('" + vSupDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eSupNo.Attributes["onClick"] = vSupDataScript;
                    }
                    break;
            }
        }

        protected void bbOKA_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_Edit");
            if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                int vTempINT;
                string vSheetNo = eSheetNo.Text.Trim();
                TextBox eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_Edit");
                TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
                TextBox eSheetNote = (TextBox)fvConsSheetA_Detail.FindControl("eSheetNote_Edit");
                TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_Edit");
                string vSQLStr_Temp;
                string vOptionString;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vSupNo = eSupNo.Text.Trim();
                string vDepNo = eDepNo.Text.Trim();
                string vSheetNote = eSheetNote.Text.Trim();
                string vRemarkA = eRemarkA.Text.Trim();
                string vMaxIndex;
                string vNewIndex;
                //寫入異動檔
                vOptionString = "修改總務耗材入庫單" + Environment.NewLine +
                                "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                "異動種類：修改" + Environment.NewLine +
                                "ConsSheet_In.aspx";
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                //備份舊資料
                vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                               "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                               "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                               "       ModifyMan, ModifyDate, DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                               "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                               "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                               "       ModifyMan, ModifyDate, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                               "  from ConsSheetA " + Environment.NewLine +
                               " where SheetNo = '" + vSheetNo + "' ";
                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                //寫入修改
                vSQLStr_Temp = "update ConsSheetA set SupNo = @SupNo, DepNo = @DepNo, SheetNote = @SheetNote, " + Environment.NewLine +
                               "       RemarkA = @RemarkA, ModifyDate = GetDate(), ModifyMan = @ModifyMan " + Environment.NewLine +
                               " where SheetNo = @SheetNo ";
                sdsConsSheetA_Detail.UpdateCommand = vSQLStr_Temp;
                sdsConsSheetA_Detail.UpdateParameters.Clear();
                sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SupNo", DbType.String, (vSupNo != "") ? vSupNo : String.Empty));
                sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (vDepNo != "") ? vDepNo : String.Empty));
                sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNote", DbType.String, (vSheetNote != "") ? vSheetNote : String.Empty));
                sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("RemarkA", DbType.String, (vRemarkA != "") ? vRemarkA : String.Empty));
                sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                sdsConsSheetA_Detail.Update();
                gridConsSheetA_List.DataBind();
                fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                fvConsSheetA_Detail.DataBind();
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
            if ((eDepNo != null) && (eDepNo.Text.Trim() != ""))
            {
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_Edit");
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo DESC";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo.Text = vDepNo;
                eDepName.Text = vDepName;
            }
        }

        protected void bbOKA_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_INS");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    string vSheetNo = PF.GetConsSheetNo(vConnStr, "IS");
                    TextBox eSupNo = (TextBox)fvConsSheetA_Detail.FindControl("eSupNo_INS");
                    TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
                    TextBox eSheetNote = (TextBox)fvConsSheetA_Detail.FindControl("eSheetNote_INS");
                    TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_INS");
                    string vSQLStr_Temp = "insert into ConsSheetA " + Environment.NewLine +
                                          "      (SheetNo, SheetMode, DepNo, SupNo, BuDate, BuMan, SheetNote, SheetStatus, StatusDate, RemarkA) " + Environment.NewLine +
                                          "values(@SheetNo, 'IS', @DepNo, @SupNo, GetDate(), @BuMan, @SheetNote, '000', GetDate(), @RemarkA) ";
                    sdsConsSheetA_Detail.InsertCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.InsertParameters.Clear();
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SupNo", DbType.String, (eSupNo.Text.Trim() != "") ? eSupNo.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNote", DbType.String, (eSheetNote.Text.Trim() != "") ? eSheetNote.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA.Text.Trim() != "") ? eRemarkA.Text.Trim() : String.Empty));
                    sdsConsSheetA_Detail.Insert();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
            if ((eDepNo != null) && (eDepNo.Text.Trim() != ""))
            {
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_INS");
                string vDepNo = eDepNo.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo.Text.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo DESC";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo.Text = vDepNo;
                eDepName.Text = vDepName;
            }
        }

        /// <summary>
        /// 主檔整單入庫
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbInStoreA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
            {
                int vTempINT;
                string vSheetNo = eSheetNo.Text.Trim();
                string vSheetNoItems;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vMaxIndex;
                string vNewIndex;
                //先檢查明細有沒有已到貨或已開單的項目
                string vSQLStr_Temp = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' and ItemStatus in ('000', '001') ";
                int vRCount;
                if (Int32.TryParse(PF.GetValue(vConnStr, vSQLStr_Temp, "RCount"), out vRCount) && vRCount > 0)
                {
                    vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = '" + vSheetNo + "' and ItemStatus in ('000', '001') order by Items";
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        while (drTemp.Read())
                        {
                            vSheetNoItems = drTemp["SheetNoItems"].ToString().Trim();
                            //備份舊資料
                            vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                            vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                            vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                            vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                           "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                           "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                           "       ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                           "select '" + vH_FirstCode + vNewIndex + "', 'SINA', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                           "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                           "       GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                           "  from ConsSheetB " + Environment.NewLine +
                                           " where SheetNoItems = '" + vSheetNoItems + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            //明細設定入庫
                            vSQLStr_Temp = "update ConsSheetB set ItemStatus = '103' where SheetNoItems = '" + vSheetNoItems + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        }
                    }
                    //寫入異動檔
                    string vOptionString = "修改總務耗材入庫單" + Environment.NewLine +
                                           "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：整單入庫" + Environment.NewLine +
                                           "ConsSheet_In.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份舊資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                                   "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                                   "       ModifyMan, ModifyDate, DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'SINA', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                                   "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                                   "       ModifyMan, ModifyDate, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //修改主檔資料
                    vSQLStr_Temp = "update ConsSheetA set SheetStatus = '030', StatusDate = GetDate() where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                }
                else
                {
                    eErrorMSG_Main.Text = "找不到已到貨明細，請重新確認！";
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 主檔整單作廢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbAbortA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
            {
                string vSheetNo = eSheetNo.Text.Trim();
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                int vTempINT;
                int vRCount;
                string vSQLStr_Temp;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vOptionString;
                string vMaxIndex;
                string vNewIndex;
                string vRCount_Str;
                string vSheetNoItems;
                //先搜尋明細有沒有狀態不是 '000' 或 '999' 的項次，有的話就不執行整單作廢
                vSQLStr_Temp = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' and ItemStatus NOT IN ('000', '999') ";
                vRCount_Str = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount");
                vRCount = Int32.TryParse(vRCount_Str, out vTempINT) ? vTempINT : 0;
                if (vRCount == 0)
                {
                    //所有項次狀態都是 '000' 或 '999'
                    vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = '" + vSheetNo + "' and ItemStatus = '000' order by Items";
                    //先把明細逐筆備份後作廢
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        while (drTemp.Read())
                        {
                            try
                            {
                                //備份舊資料
                                vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                                vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                                vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                                vSheetNoItems = drTemp["SheetNoItems"].ToString().Trim();
                                vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                               "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                               "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                               "       ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                               "select '" + vH_FirstCode + vNewIndex + "', 'ABRA', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                               "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                               "       GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                               "  from ConsSheetB " + Environment.NewLine +
                                               " where SheetNoItems = '" + vSheetNoItems + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                                //設定作廢
                                vSQLStr_Temp = "update ConsSheetB set ItemStatus = '999' where SheetNoItems = '" + vSheetNoItems + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            }
                            catch (Exception eMessage)
                            {
                                eErrorMSG_Main.Text = eMessage.Message;
                                eErrorMSG_Main.Visible = true;
                            }
                        }
                    }
                    //寫入異動檔
                    vOptionString = "修改總務耗材入庫單" + Environment.NewLine +
                                    "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：作廢" + Environment.NewLine +
                                    "ConsSheet_In.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    try
                    {
                        //備份舊資料
                        vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                        vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                        vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                        vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                       "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                                       "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                                       "       ModifyMan, ModifyDate, DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                       "select '" + vH_FirstCode + vNewIndex + "', 'ABRA', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                                       "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                                       "       ModifyMan, ModifyDate, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                       "  from ConsSheetA " + Environment.NewLine +
                                       " where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        //主檔設定作廢
                        vSQLStr_Temp = "update ConsSheetA set SheetStatus = '999', StatusDate = GetDate() where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    }
                    catch (Exception eMessage)
                    {
                        eErrorMSG_Main.Text = eMessage.Message;
                        eErrorMSG_Main.Visible = true;
                    }
                }
                else
                {
                    eErrorMSG_Main.Text = "部份明細已到貨或入庫，不可整單作廢！";
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 主檔刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDeleteA_List_Click(object sender, EventArgs e)
        {
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if ((eSheetNo != null) && (eSheetNo.Text.Trim() != ""))
            {
                int vTempINT;
                int vRCount;
                string vOptionString;
                string vSQLStr_Temp;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vMaxIndex;
                string vNewIndex;
                string vSheetNo = eSheetNo.Text.Trim();
                string vRCountStr;
                //檢查有沒有已經到貨或入庫的明細
                vSQLStr_Temp = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' and ItemStatus in ('001', '103') ";
                vRCountStr = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount");
                vRCount = Int32.TryParse(vRCountStr, out vTempINT) ? vTempINT : 0;
                if (vRCount == 0)
                {
                    //寫入異動
                    vOptionString = "刪除總務耗材入庫單" + Environment.NewLine +
                                    "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：刪除" + Environment.NewLine +
                                    "ConsSheet_In.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份舊資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                                   "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                                   "       ModifyMan, ModifyDate, DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'DELA', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxRate, " + Environment.NewLine +
                                   "       TaxType, TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, " + Environment.NewLine +
                                   "       ModifyMan, ModifyDate, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //刪除資料
                    vSQLStr_Temp = "delete ConsSheetA where SheetNo = @SheetNo";
                    sdsConsSheetA_Detail.DeleteCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.DeleteParameters.Clear();
                    sdsConsSheetA_Detail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetA_Detail.Delete();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.DataBind();
                }
                else
                {
                    eErrorMSG_Main.Text = "明細已部份到貨 / 入庫，不允許整單刪除";
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        protected void fvConsSheetB_Detail_DataBound(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            string vSQLStr_Temp;
            Label eSheetNoItems;
            string vConsDataURL;
            string vConsDataScript;
            string vConsUnit;
            TextBox eConsNo;
            TextBox eConsName;
            DropDownList ddlConsUnit;
            Label eConsUnit;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvConsSheetB_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_List");
                    if ((eSheetNoItems != null) && (eSheetNoItems.Text.Trim() != ""))
                    {
                        Label eItemStatus = (Label)fvConsSheetB_Detail.FindControl("eItemStatus_List");
                        string vItemStatus = eItemStatus.Text.Trim();
                        Button bbEditB = (Button)fvConsSheetB_Detail.FindControl("bbEditB_List");
                        Button bbInstoreB = (Button)fvConsSheetB_Detail.FindControl("bbInstoreB_List");
                        Button bbAbortB = (Button)fvConsSheetB_Detail.FindControl("bbAbortB_List");
                        Button bbDeleteB = (Button)fvConsSheetB_Detail.FindControl("bbDeleteB_List");
                        bbEditB.Enabled = (vItemStatus == "000");
                        bbInstoreB.Enabled = (vItemStatus == "000");
                        bbAbortB.Enabled = (vItemStatus == "000");
                        bbDeleteB.Enabled = (vItemStatus == "000");
                    }
                    break;
                case FormViewMode.Edit:
                    eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_Edit");
                    if ((eSheetNoItems != null) && (eSheetNoItems.Text.Trim() != ""))
                    {
                        eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
                        eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_Edit");
                        vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsNameID=" + eConsName.ClientID;
                        vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eConsNo.Attributes["onClick"] = vConsDataScript;
                        ddlConsUnit = (DropDownList)fvConsSheetB_Detail.FindControl("ddlConsUnit_C_Edit");
                        eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_Edit");
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                                       " order by ClassNo ";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            if (drTemp.HasRows)
                            {
                                ddlConsUnit.Items.Clear();
                                while (drTemp.Read())
                                {
                                    ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                    ddlConsUnit.Items.Add(liTemp);
                                }
                                vConsUnit = eConsUnit.Text.Trim();
                                ddlConsUnit.SelectedIndex = ddlConsUnit.Items.IndexOf(ddlConsUnit.Items.FindByValue(vConsUnit));
                            }
                        }
                    }
                    break;
                case FormViewMode.Insert:
                    eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_INS");
                    if (eSheetNoItems != null)
                    {
                        eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                        eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_INS");
                        vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsNameID=" + eConsName.ClientID;
                        vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eConsNo.Attributes["onClick"] = vConsDataScript;
                        ddlConsUnit = (DropDownList)fvConsSheetB_Detail.FindControl("ddlConsUnit_C_INS");
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB where FKey = '耗材庫存        CONSUMABLES     ConsUnit' " + Environment.NewLine +
                                       " order by ClassNo ";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            if (drTemp.HasRows)
                            {
                                ddlConsUnit.Items.Clear();
                                while (drTemp.Read())
                                {
                                    ListItem liTemp = new ListItem(drTemp["ClassTxt"].ToString().Trim(), drTemp["ClassNo"].ToString().Trim());
                                    ddlConsUnit.Items.Add(liTemp);
                                }
                            }
                        }
                    }
                    break;
            }
        }

        protected void bbOKB_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            int vTempINT;
            double vTempFloat;
            string vOptionString;
            string vSQLStr_Temp;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vMaxIndex;
            string vNewIndex;
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_Edit");
            if ((eSheetNoItems != null) && (eSheetNoItems.Text.Trim() != ""))
            {
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_Edit");
                string vSheetNo = eSheetNo.Text.Trim();
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_Edit");
                string vItems = eItems.Text.Trim();
                try
                {
                    //寫入異動記錄
                    vOptionString = "修改總務耗材入庫單明細" + Environment.NewLine +
                                    "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                    "項次：[ " + vItems + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：修改" + Environment.NewLine +
                                    "ConsSheet_In.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份明細資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //更新明細資料
                    TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_Edit");
                    TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
                    TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_Edit");
                    Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_Edit");
                    string vRemarkB = eRemarkB.Text.Trim();
                    string vConsNo = eConsNo.Text.Trim();
                    string vConsUnit = eConsUnit.Text.Trim();
                    string vQTYStr = eQuantity.Text.Trim();
                    vSQLStr_Temp = "update ConsSheetB set ConsNo = @ConsNo, Qunantity = @Quantity, RemarkB = @RemarkB, " + Environment.NewLine +
                                 "                        ConsUnit = @ConsUnit, ModifyDate = GetDate(), ModifyMan = @ModifyMan " + Environment.NewLine +
                                 " where SheetNoItems = @SheetNoItems ";
                    sdsConsSheetB_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.UpdateParameters.Clear();
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ConsNo", DbType.String, (vConsNo != "") ? vConsNo : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Quantity", DbType.Double, double.TryParse(vQTYStr, out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("RemarkB", DbType.String, (vRemarkB != "") ? vRemarkB : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ConsUnit", DbType.String, (vConsUnit != "") ? vConsUnit : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.Update();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 單位下拉選單變更選項
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlConsUnit_C_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlConsUnit = (DropDownList)fvConsSheetB_Detail.FindControl("ddlConsUnit_C_Edit");
            if (ddlConsUnit != null)
            {
                Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_Edit");
                eConsUnit.Text = ddlConsUnit.SelectedValue.Trim();
            }
        }

        protected void bbOKB_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            int vTempINT;
            double vTempFloat;
            string vSQLStr_Temp;
            string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
            string vMaxIndex;
            string vItems;
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_INS");
            if (eSheetNoItems != null)
            {
                try
                {
                    Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
                    string vSheetNo = eSheetNo.Text.Trim();
                    TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                    string vConsNo = eConsNo.Text.Trim();
                    Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                    string vConsUnit = eConsUnit.Text.Trim();
                    TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_INS");
                    string vRemarkB = eRemarkB.Text.Trim();
                    TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_INS");
                    string vQTYStr = eQuantity.Text.Trim();
                    vSQLStr_Temp = "select max(Items) MaxIndex from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vItems = Int32.TryParse(vMaxIndex, out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    string vSheetNoItems = vSheetNo + vItems;
                    vSQLStr_Temp = "insert into ConsSheetB " + Environment.NewLine +
                                   "      (SheetNoItems, SheetNo, Items, ConsNo, Quantity, ConsUnit, RemarkB, BuDate, BuMan, ItemStatus, QtyMode) " + Environment.NewLine +
                                   "values(@SheetNoItems, @SheetNo, @Items, @ConsNo, @Quantity, @ConsUnit, @RemarkB, GetDate(), @BuMan, '000', 1) ";
                    sdsConsSheetB_Detail.InsertCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.InsertParameters.Clear();
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ConsNo", DbType.String, (vConsNo != "") ? vConsNo : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Quantity", DbType.Double, double.TryParse(vQTYStr, out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ConsUnit", DbType.String, (vConsUnit != "") ? vConsUnit : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("RemarkB", DbType.String, (vRemarkB != "") ? vRemarkB : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsConsSheetB_Detail.Insert();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 單位下拉選單變更選項
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlConsUnit_C_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlConsUnit = (DropDownList)fvConsSheetB_Detail.FindControl("ddlConsUnit_C_INS");
            if (ddlConsUnit != null)
            {
                Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                eConsUnit.Text = ddlConsUnit.SelectedValue.Trim();
            }
        }

        private void PickupPurchase()
        {
            string vSelectStr = "select b.SheetNoItems, a.SheetNo, b.Items, b.ConsNo, c.ConsName, b.Quantity, c.StockQty, b.RemarkB " + Environment.NewLine +
                                "  from ConsSheetA a left join ConsSheetB b on b.SheetNo = a.SheetNo " + Environment.NewLine +
                                "                    left join Consumables c on c.ConsNo = b.ConsNo " + Environment.NewLine +
                                " where a.SheetMode = 'PS' " + Environment.NewLine +
                                "   and b.ItemStatus = '000' ";
            sdsPickup_List.SelectCommand = vSelectStr;
            sdsPickup_List.Select(new DataSourceSelectArguments());
            gridPickup_List.DataBind();
        }

        /// <summary>
        /// 明細挑單 (沒有明細時)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPickup_Empty_Click(object sender, EventArgs e)
        {
            PickupPurchase();
            plPickupDetail.Visible = true;
            plShowDataB.Visible = false;
        }

        /// <summary>
        /// 明細挑單 (已有明細時)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPickup_List_Click(object sender, EventArgs e)
        {
            PickupPurchase();
            plPickupDetail.Visible = true;
            plShowDataB.Visible = false;
        }

        /// <summary>
        /// 明細入庫
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbInstoreB_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_List");
            if ((eSheetNoItems != null) && (eSheetNoItems.Text.Trim() != ""))
            {
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_List");
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_List");
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                string vSheetNo = eSheetNo.Text.Trim();
                string vItems = eItems.Text.Trim();
                string vOptionString;
                string vSQLStr_Temp;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vMaxIndex;
                string vNewIndex;
                int vTempINT;
                try
                {
                    //寫入異動記錄
                    vOptionString = "修改總務耗材入庫單明細" + Environment.NewLine +
                                    "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                    "項次：[ " + vItems + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：明細入庫" + Environment.NewLine +
                                    "ConsSheet_In.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份明細資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'SINB', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //明細設定入庫
                    vSQLStr_Temp = "update ConsSheetB set ItemStatus = '103', ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "' where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 明細作廢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbAbortB_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_List");
            if ((eSheetNoItems != null) && (eSheetNoItems.Text.Trim() != ""))
            {
                Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_List");
                Label eItems = (Label)fvConsSheetB_Detail.FindControl("Item");
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                string vSheetNo = eSheetNo.Text.Trim();
                string vItems = eItems.Text.Trim();
                int vTempINT;
                string vOptionString;
                string vSQLStr_Temp;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vMaxIndex;
                string vNewIndex;
                try
                {
                    //寫入異動檔
                    vOptionString = "修改總務耗材入庫單明細" + Environment.NewLine +
                                    "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                    "項次：[ " + vItems + " ]" + Environment.NewLine +
                                    "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                    "異動種類：明細作廢" + Environment.NewLine +
                                    "ConsSheet_In.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份舊資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'ABRB', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                   "       GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //明細設定作廢
                    vSQLStr_Temp = "update ConsSheetB set ItemStatus = '999', ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "' where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 明細刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDeleteB_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_List");
            if ((eSheetNoItems!=null)&&(eSheetNoItems.Text.Trim()!=""))
            {
                Label eItemStatus = (Label)fvConsSheetB_Detail.FindControl("eItemStatus");
                string vItemStatus = eItemStatus.Text.Trim();
                if ((vItemStatus == "000") || (vItemStatus == "999") || (vItemStatus == "998"))
                {
                    Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_List");
                    Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_List");
                    string vSheetNoItems = eSheetNoItems.Text.Trim();
                    string vSheetNo = eSheetNo.Text.Trim();
                    string vItems = eItems.Text.Trim();
                    int vTempINT;
                    string vOptionString;
                    string vSQLStr_Temp;
                    string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    string vMaxIndex;
                    string vNewIndex;
                    try
                    {
                        //寫入異動檔
                        vOptionString = "刪除總務耗材入庫單明細" + Environment.NewLine +
                                        "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                        "項次：[ " + vItems + " ]" + Environment.NewLine +
                                        "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                        "異動種類：明細刪除" + Environment.NewLine +
                                        "ConsSheet_In.aspx";
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                        //備份舊資料
                        vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                        vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                        vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                        vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                       "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                       "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                       "       ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                       "select '" + vH_FirstCode + vNewIndex + "', 'DELB', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                       "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                       "       GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                       "  from ConsSheetB " + Environment.NewLine +
                                       " where SheetNoItems = '" + vSheetNoItems + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        //刪除明細
                        vSQLStr_Temp = "delete ConsSheetB where SheetNoItems = @SheetNoItems";
                        sdsConsSheetB_Detail.DeleteCommand = vSQLStr_Temp;
                        sdsConsSheetB_Detail.DeleteParameters.Clear();
                        sdsConsSheetB_Detail.DeleteParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                        sdsConsSheetB_Detail.Delete();
                        sdsConsSheetB_Detail.DataBind();
                        fvConsSheetB_Detail.DataBind();
                    }
                    catch (Exception eMessage)
                    {
                        eErrorMSG_Main.Text = eMessage.Message;
                        eErrorMSG_Main.Visible = true;
                    }
                }
                else
                {
                    eErrorMSG_Main.Text = "明細已到貨 / 入庫，不可刪除";
                    eErrorMSG_Main.Visible = true;
                }
            }
        }

        /// <summary>
        /// 全選
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSelectAll_P_Click(object sender, EventArgs e)
        {
            foreach (GridViewRow vCurrentRows in gridPickup_List.Rows)
            {
                CheckBox cbTemp = (CheckBox)vCurrentRows.FindControl("cbChoise");
                if (cbTemp != null) cbTemp.Checked = true;
            }
        }

        /// <summary>
        /// 取消全選
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbUnselectAll_P_Click(object sender, EventArgs e)
        {
            foreach (GridViewRow vCurrentRows in gridPickup_List.Rows)
            {
                CheckBox cbTemp = (CheckBox)vCurrentRows.FindControl("cbChoise");
                if (cbTemp != null) cbTemp.Checked = false;
            }
        }

        protected void bbOK_P_Click(object sender, EventArgs e)
        {
            eErrorMSG_Main.Text = "";
            eErrorMSG_Main.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                int vTempINT;
                string vSQLStr_Temp;
                string vSheetNoItems;
                string vSourceNo;
                string vMaxIndex;
                string vNewIndex;
                string vSheetNo = eSheetNo.Text.Trim();
                try
                {
                    foreach (GridViewRow vCurrentRows in gridPickup_List.Rows)
                    {
                        CheckBox cbTemp = (CheckBox)vCurrentRows.FindControl("cbChoise");
                        if (cbTemp != null && cbTemp.Checked)
                        {
                            vSQLStr_Temp = "select max(Items) MaxIndex from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                            vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                            vNewIndex = Int32.TryParse(vMaxIndex, out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                            vSheetNoItems = vSheetNo + vNewIndex;
                            vSourceNo = gridPickup_List.DataKeys[vCurrentRows.RowIndex].Value.ToString().Trim();
                            vSQLStr_Temp = "insert into ConsSheetB " + Environment.NewLine +
                                           "      (SheetNoItems, SheetNo, Items, ConsNo, ItemStatus, Quantity, ConsUnit, RemarkB, BuDate, BuMan, SourceNo) " + Environment.NewLine +
                                           "select '" + vSheetNoItems + "', '" + vSheetNo + "', '" + vNewIndex + "', ConsNo, '000', Quantity, ConsUnit, RemarkB, GetDate(), " + Environment.NewLine +
                                           "       '" + vLoginID + "', SheetNoItems " + Environment.NewLine +
                                           "  from ConsSheetB " + Environment.NewLine +
                                           " where SheetNoItems = '" + vSourceNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            vSQLStr_Temp = "update ConsSheetB set ItemStatus = '103' where SheetNoItems = '" + vSourceNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        }
                    }
                    gridConsSheetB_List.DataBind();
                    plShowDataB.Visible = true;
                    plPickupDetail.Visible = false;
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_Main.Text = eMessage.Message;
                    eErrorMSG_Main.Visible = true;
                    throw;
                }
            }
        }

        protected void bbCancel_P_Click(object sender, EventArgs e)
        {
            plShowDataB.Visible = true;
            plPickupDetail.Visible = false;
        }
    }
}