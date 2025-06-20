using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class ConsSheet_Out : Page
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
                    string vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDate_S_Search.ClientID;
                    string vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBuDate_S_Search.Attributes["onClick"] = vDateScript;

                    vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDate_E_Search.ClientID;
                    vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                    eBuDate_E_Search.Attributes["onClick"] = vDateScript;

                    if (!IsPostBack)
                    {
                        eBuDate_S_Search.Text = "";
                        eBuDate_E_Search.Text = "";
                        plSearch.Visible = true;
                        plShowData_A.Visible = true;
                        plShowData_B.Visible = true;
                        plShowData_C.Visible = false;
                        plPrint.Visible = false;
                    }
                    else
                    {
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

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        private string GetSearchStr()
        {
            DateTime vDate_S;
            DateTime vDate_E;
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   AND a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_AssignMan = (eAssignMan_Search.Text.Trim() != "") ? "   AND a.AssignMan = '" + eAssignMan_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuDate = (DateTime.TryParse(eBuDate_S_Search.Text.Trim(), out vDate_S) && DateTime.TryParse(eBuDate_E_Search.Text.Trim(), out vDate_E)) ?
                                  "   AND a.BuDate between '" + PF.TransDateString(vDate_S, "B") + "' and '" + PF.TransDateString(vDate_E, "B") + "' " + Environment.NewLine :
                                  (!DateTime.TryParse(eBuDate_S_Search.Text.Trim(), out vDate_S) && DateTime.TryParse(eBuDate_E_Search.Text.Trim(), out vDate_E)) ?
                                  "   AND a.BuDate = '" + PF.TransDateString(vDate_E, "B") + "' " + Environment.NewLine :
                                  DateTime.TryParse(eBuDate_S_Search.Text.Trim(), out vDate_S) && !DateTime.TryParse(eBuDate_E_Search.Text.Trim(), out vDate_E) ?
                                  "   AND a.BuDate = '" + PF.TransDateString(vDate_S, "B") + "' " + Environment.NewLine : "";
            string vResultStr = "SELECT a.SheetNo, a.SheetNote, a.DepNo, d.NAME AS DepName, a.AssignMan, e.NAME AS AssignManName, " + Environment.NewLine +
                                "       CONVERT (varchar(10), a.BuDate, 111) AS BuDate, a.TotalAmount, f.ClassTxt AS SheetStatus_C " + Environment.NewLine +
                                "  FROM ConsSheetA AS a LEFT OUTER JOIN DEPARTMENT AS d ON d.DEPNO = a.DepNo " + Environment.NewLine +
                                "                       LEFT OUTER JOIN EMPLOYEE AS e ON e.EMPNO = a.AssignMan " + Environment.NewLine +
                                "                       LEFT OUTER JOIN DBDICB AS f on f.ClassNo = a.SheetStatus and f.FKey = '總務耗材進出單  fmConsSheetA    SheetStatus' " + Environment.NewLine +
                                " WHERE isnull(a.SheetMode, '') = 'OS' " + Environment.NewLine +
                                vWStr_DepNo +
                                vWStr_AssignMan +
                                vWStr_BuDate +
                                " ORDER BY SheetNo ";
            return vResultStr;
        }

        private void OpenData()
        {
            string vSelectStr = GetSearchStr();
            gridConsSheetA_List.DataSourceID = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connSearch = new SqlConnection(vConnStr))
            {
                DataTable dtSearch = new DataTable();
                SqlDataAdapter daSearch = new SqlDataAdapter(vSelectStr, connSearch);
                connSearch.Open();
                daSearch.Fill(dtSearch);
                gridConsSheetA_List.DataSource = dtSearch;
                gridConsSheetA_List.DataBind();
            }
        }

        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo_Search.Text.Trim();
            string vSQLStr = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            string vDepName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vDepName.Trim() == "")
            {
                vDepName = eDepNo_Search.Text.Trim();
                vSQLStr = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo ";
                vDepNo = PF.GetValue(vConnStr, vSQLStr, "DepNo");
            }
            eDepNo_Search.Text = vDepNo.Trim();
            eDepName_Search.Text = vDepName.Trim();
        }

        protected void eAssignMan_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vAssignMan = eAssignMan_Search.Text.Trim();
            string vSQLStr = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
            string vAssignManName = PF.GetValue(vConnStr, vSQLStr, "Name");
            if (vAssignManName == "")
            {
                vAssignManName = eAssignMan_Search.Text.Trim();
                vSQLStr = "select top 1 EmpNo from Employee where [Name] like '%" + vAssignManName + "%' order by EmpNo DESC";
                vAssignMan = PF.GetValue(vConnStr, vSQLStr, "EmpNo");
            }
            eAssignMan_Search.Text = vAssignMan.Trim();
            eAssignManName_Search.Text = vAssignManName.Trim();
        }

        protected void fvConsSheetA_Detail_DataBound(object sender, EventArgs e)
        {
            string vDateURL;
            string vDateScript;
            Label eSheetNo;
            TextBox eBuDate;
            switch (fvConsSheetA_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_Edit");
                    if (eSheetNo != null)
                    {
                        eBuDate = (TextBox)fvConsSheetA_Detail.FindControl("eBuDateA_Edit");
                        vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDate.ClientID;
                        vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                        eBuDate.Attributes["onClick"] = vDateScript;
                    }
                    break;
                case FormViewMode.Insert:
                    eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_INS");
                    if (eSheetNo != null)
                    {
                        Label eBuMan = (Label)fvConsSheetA_Detail.FindControl("eBuManA_INS");
                        eBuMan.Text = vLoginID;
                        Label eBuManName = (Label)fvConsSheetA_Detail.FindControl("eBuManNameA_INS");
                        eBuManName.Text = vLoginName;
                        eBuDate = (TextBox)fvConsSheetA_Detail.FindControl("eBuDateA_INS");
                        eBuDate.Text = DateTime.Today.ToShortDateString();
                        vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDate.ClientID;
                        vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                        eBuDate.Attributes["onClick"] = vDateScript;
                    }
                    break;
            }
        }

        protected void bbOKA_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_Edit");
            if (eSheetNo != null)
            {
                string vSheetNo = eSheetNo.Text.Trim();
                //寫入異動記錄
                string vOptionString = "修改總務耗材請購發料單" + Environment.NewLine +
                                       "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                       "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                       "異動種類：修改" + Environment.NewLine +
                                       "ConsSheet_Order.aspx";
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                //備份舊資料
                int vTempINT;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                string vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                               "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                               "       SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                               "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                               "       SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                               "  from ConsSheetA " + Environment.NewLine +
                               " where SheetNo = '" + vSheetNo + "' ";
                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                try
                {
                    //寫入異動
                    TextBox eSheetNote = (TextBox)fvConsSheetA_Detail.FindControl("eSheetNote_Edit");
                    string vSheetNote = eSheetNote.Text.Trim();
                    TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
                    string vDepNo = eDepNo.Text.Trim();
                    TextBox eBuDate = (TextBox)fvConsSheetA_Detail.FindControl("eBuDateA_Edit");
                    string vBuDateStr = eBuDate.Text.Trim();
                    TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_Edit");
                    string vRemarkA = eRemarkA.Text.Trim();
                    DateTime vTempDate;
                    vSQLStr_Temp = "update ConsSheetA " + Environment.NewLine +
                                   "   set SheetNote = @SheetNote, DepNo = @DepNo, BuDate = @BuDate, RemarkA = @RemarkA, ModifyDate = GetDate(), ModifyMan = @ModifyMan " + Environment.NewLine +
                                   " where SheetNo = @SheetNo ";
                    sdsConsSheetA_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.UpdateParameters.Clear();
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNote", DbType.String, (vSheetNote != "") ? vSheetNote : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (vDepNo != "") ? vDepNo : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("BuDate", DbType.String, DateTime.TryParse(vBuDateStr, out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("RemarkA", DbType.String, (vRemarkA != "") ? vRemarkA : String.Empty));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetA_Detail.Update();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = false;
                }
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_Edit");
            if (eDepNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
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
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_INS");
            if (eSheetNo != null)
            {
                DateTime vTempDate;
                TextBox eSheetNote = (TextBox)fvConsSheetA_Detail.FindControl("eSheetNote_INS");
                string vSheetNote = eSheetNote.Text.Trim();
                TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
                string vDepNo = eDepNo.Text.Trim();
                TextBox eBuDate = (TextBox)fvConsSheetA_Detail.FindControl("eBuDateA_INS");
                string vBuDateStr = eBuDate.Text.Trim();
                TextBox eRemarkA = (TextBox)fvConsSheetA_Detail.FindControl("eRemarkA_INS");
                string vRemarkA = eRemarkA.Text.Trim();
                string vSheetMode = "OS";
                string vSheetNo = PF.GetConsSheetNo(vConnStr, vSheetMode);
                try
                {
                    string vSQLStr_Termp = "insert into ConsSheetA " + Environment.NewLine +
                                           "      (SheetNo, SheetMode, BuDate, BuMan, SheetStatus, StatusDate, DepNo, RemarkA, SheetNote) " + Environment.NewLine +
                                           "values(@SheetNo, @SheetMode, @BuDate, @BuMan, '000', GetDate(), @DepNo, @RemarkA, @SheetNote) ";
                    sdsConsSheetA_Detail.InsertCommand = vSQLStr_Termp;
                    sdsConsSheetA_Detail.InsertParameters.Clear();
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetMode", DbType.String, vSheetMode));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.TryParse(vBuDateStr, out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (vDepNo != "") ? vDepNo : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("RemarkA", DbType.String, (vRemarkA != "") ? vRemarkA : String.Empty));
                    sdsConsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNote", DbType.String, (vSheetNote != "") ? vSheetNo : String.Empty));
                    sdsConsSheetA_Detail.Insert();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvConsSheetA_Detail.FindControl("eDepNo_INS");
            if (eDepNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
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

        protected void bbAbortA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNOA_List");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSheetNo = eSheetNo.Text.Trim();
                string vSheetNoItems;
                string vSourceNo;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vOptionString;
                string vMaxIndex;
                string vNewIndex;
                int vTempINT;
                //先檢查有沒有已經撥付或是簽收了的明細
                string vSQLStr_Temp = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' and ItemStatus in ('102', '103') ";
                string vRCountStr = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount");
                if (Int32.Parse(vRCountStr) == 0)
                {
                    try
                    {
                        vSQLStr_Temp = "select SheetNoItems, SourceNo from ConsSheetB where SheetNo = '" + vSheetNo + "' order by Items";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                vSheetNoItems = drTemp["SheetNoItems"].ToString().Trim();
                                vSourceNo = drTemp["SourceNo"].ToString().Trim();
                                //把明細備份到歷史檔
                                vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                                vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                                vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                                vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                               "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                               "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                               "select '" + vH_FirstCode + vNewIndex + "', 'ABR', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                               "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                               "  from ConsSheetB " + Environment.NewLine +
                                               " where SheetNoItems = '" + vSheetNoItems + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                                //把明細作廢
                                vSQLStr_Temp = "update ConsSheetB set ItemStatus = '999' where SheetNoItems = '" + vSheetNoItems + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            }
                        }
                        //寫入異動檔
                        vOptionString = "修改總務耗材請購發料單" + Environment.NewLine +
                                        "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                        "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                        "異動種類：作廢" + Environment.NewLine +
                                        "ConsSheet_Order.aspx";
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                        //備份主檔資料
                        vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                        vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                        vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                        vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                       "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, " + Environment.NewLine +
                                       "       TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, " + Environment.NewLine +
                                       "       AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                       "select '" + vH_FirstCode + vNewIndex + "', 'ABR', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, " + Environment.NewLine +
                                       "       TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, " + Environment.NewLine +
                                       "       AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                       "  from ConsSheetA " + Environment.NewLine +
                                       " where SheetNo = '" + vSheetNo + "' ";
                        PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        //作廢主檔
                        vSQLStr_Temp = "update ConsSheetA set SheetStatus = '999', StatusDate = GetDate() where SheetNo = @SheetNo";
                        sdsConsSheetA_Detail.UpdateCommand = vSQLStr_Temp;
                        sdsConsSheetA_Detail.UpdateParameters.Clear();
                        sdsConsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                        sdsConsSheetA_Detail.Update();
                        gridConsSheetA_List.DataBind();
                        fvConsSheetA_Detail.DataBind();
                    }
                    catch (Exception eMessage)
                    {
                        eErrorMSG_A.Text = eMessage.Message;
                        eErrorMSG_A.Visible = true;
                    }
                }
                else
                {
                    eErrorMSG_A.Text = "部份明細已撥付，本單不可作廢";
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        /// <summary>
        /// 列印簽收單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_List_Click(object sender, EventArgs e)
        {
            DataTable dtPrint;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (eSheetNo != null)
            {
                Label eDepName = (Label)fvConsSheetA_Detail.FindControl("eDepName_List");
                string vSheetNo = eSheetNo.Text.Trim();
                string vSQLStr_Print = "SELECT b.Items, b.ConsNo, c.ConsName, b.Quantity, b.RemarkB " + Environment.NewLine +
                                       "  FROM ConsSheetB AS b LEFT OUTER JOIN Consumables AS c ON c.ConsNo = b.ConsNo " + Environment.NewLine +
                                       " WHERE b.SheetNo = '" + vSheetNo + "' ";
                using (SqlConnection connPrint = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daPrint = new SqlDataAdapter(vSQLStr_Print, connPrint);
                    connPrint.Open();
                    dtPrint = new DataTable();
                    daPrint.Fill(dtPrint);
                }

                string vCompanyName = PF.GetValue(vConnStr, "select [Name] from [Custom] where [Code] = 'A000' and Types = 'O' ", "Name");
                string vReportName = "請購發料單";
                string vReportPath = @"Report\ConsSheet_OutP.rdlc";
                string vAssignDep = eDepName.Text.Trim();

                ReportDataSource rdsPrint = new ReportDataSource("ConsSheet_OutP", dtPrint);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = vReportPath;
                rvPrint.LocalReport.DataSources.Add(rdsPrint);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("AssignDep", vAssignDep));
                rvPrint.LocalReport.Refresh();
                plPrint.Visible = true;
                plSearch.Visible = false;
                plShowData_A.Visible = false;
                plShowData_B.Visible = false;
                plShowData_C.Visible = false;
            }
        }

        /// <summary>
        /// 單據結案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCloseA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    string vSheetNo = eSheetNo.Text.Trim();
                    //寫入異動記錄
                    string vOptionString = "修改總務耗材請購發料單" + Environment.NewLine +
                                           "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：結案" + Environment.NewLine +
                                           "ConsSheet_Order.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //因為只是結案，不用特別再複製到歷史資料檔
                    string vSQLStr_Temp = "update ConsSheetA set SheetStatus = '998', StatusDate = GetDate() where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        /// <summary>
        /// 刪除主檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDeleteA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                int vTempINT;
                int vRCount;
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vSheetNo = eSheetNo.Text.Trim();
                string vSQLStr_Temp;
                string vMaxIndex;
                string vNewIndex;
                string vRCountStr;
                string vSheetNoItems;
                try
                {
                    //寫入異動記錄
                    string vOptionString = "刪除總務耗材請購發料單" + Environment.NewLine +
                                           "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：刪除" + Environment.NewLine +
                                           "ConsSheet_Order.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //檢查有沒有明細，有明細的話要一併刪除
                    vSQLStr_Temp = "select count(Items) RCount from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                    vRCountStr = PF.GetValue(vConnStr, vSQLStr_Temp, "RCount");
                    vRCount = Int32.Parse(vRCountStr);
                    if (vRCount > 0)
                    {
                        //有明細資料，先備份再刪除
                        vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = '" + vSheetNo + "' order by Items";
                        using (SqlConnection connTemp = new SqlConnection(vConnStr))
                        {
                            SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                            connTemp.Open();
                            SqlDataReader drTemp = cmdTemp.ExecuteReader();
                            while (drTemp.Read())
                            {
                                vSheetNoItems = drTemp["SheetNoItems"].ToString().Trim();
                                vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                                vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                                vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                                vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                               "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                               "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                               "       ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                               "select '" + vH_FirstCode + vNewIndex + "', 'DELA', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                               "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, " + Environment.NewLine +
                                               "       GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                               "  from ConsSheetB " + Environment.NewLine +
                                               " where SheetNoItems = '" + vSheetNoItems + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp); //備份明細舊資料
                                vSQLStr_Temp = "delete ConsSheetB where SheetNoItems = '" + vSheetNoItems + "' ";
                                PF.ExecSQL(vConnStr, vSQLStr_Temp); //刪除明細資料
                            }
                        }
                    }
                    //備份主檔資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'DELA', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //刪除主檔資料
                    vSQLStr_Temp = "delete ConsSheetA where SheetNo = @SheetNo";
                    sdsConsSheetA_Detail.DeleteCommand = vSQLStr_Temp;
                    sdsConsSheetA_Detail.DeleteParameters.Clear();
                    sdsConsSheetA_Detail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetA_Detail.Delete();
                    gridConsSheetA_List.DataBind();
                    fvConsSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        protected void fvConsSheetB_Detail_DataBound(object sender, EventArgs e)
        {
            TextBox eConsNo;
            TextBox eConsName;
            string vConsDataURL;
            string vConsDataScript;
            Label eSheetNoA;
            Label eSheetNoB;

            switch (fvConsSheetB_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    Label eSheetStatus = (Label)fvConsSheetA_Detail.FindControl("eSheetStatus_List");
                    if (eSheetStatus != null)
                    {
                        Button bbEdit = (Button)fvConsSheetB_Detail.FindControl("bbEditB_List");
                        Button bbPatchNew = (Button)fvConsSheetB_Detail.FindControl("bbPatchNew_List");
                        Button bbNew = (Button)fvConsSheetB_Detail.FindControl("bbNewB_List");
                        Button bbAbort = (Button)fvConsSheetB_Detail.FindControl("bbAbortB_List");
                        Button bbDelete = (Button)fvConsSheetB_Detail.FindControl("bbDeleteB_List");
                        switch (eSheetStatus.Text.Trim())
                        {
                            case "998":
                            case "999":
                                bbEdit.Enabled = false;
                                bbPatchNew.Enabled = false;
                                bbNew.Enabled = false;
                                bbAbort.Enabled = false;
                                bbDelete.Enabled = false;
                                break;

                            default:
                                bbEdit.Enabled = true;
                                bbPatchNew.Enabled = true;
                                bbNew.Enabled = true;
                                bbAbort.Enabled = true;
                                bbDelete.Enabled = true;
                                break;
                        }
                    }
                    break;
                case FormViewMode.Edit:
                    eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
                    eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_Edit");
                    vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsPriceID=" + eConsName.ClientID;
                    vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                    eConsNo.Attributes["onClick"] = vConsDataScript;
                    break;
                case FormViewMode.Insert:
                    eSheetNoA = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
                    eSheetNoB = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_INS");
                    if (eSheetNoA != null && eSheetNoB != null)
                    {
                        eSheetNoB.Text = eSheetNoA.Text.Trim();
                        eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                        eConsName = (TextBox)fvConsSheetB_Detail.FindControl("eConsName_INS");
                        vConsDataURL = "SearchConsData.aspx?ConsNoID=" + eConsNo.ClientID + "&ConsPriceID=" + eConsName.ClientID;
                        vConsDataScript = "window.open('" + vConsDataURL + "','','height=600, width=650,status=no,toolbar=no,menubar=no,location=no','')";
                        eConsNo.Attributes["onClick"] = vConsDataScript;
                        Label eItemStatus = (Label)fvConsSheetB_Detail.FindControl("eItemStatus_INS");
                        eItemStatus.Text = "000";
                        Label eItemStatus_C = (Label)fvConsSheetB_Detail.FindControl("eItemStatus_C_INS");
                        eItemStatus_C.Text = "已開單";
                    }
                    break;
            }
        }

        protected void bbOKB_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNoItems = (Label)fvConsSheetB_Detail.FindControl("eSheetNoItems_Edit");
            if (eSheetNoItems != null)
            {
                try
                {
                    string vSheetNoItems = eSheetNoItems.Text.Trim();
                    Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_Edit");
                    string vSheetNo = eSheetNo.Text.Trim();
                    Label eItems = (Label)fvConsSheetB_Detail.FindControl("eItems_Edit");
                    string vItems = eItems.Text.Trim();
                    //寫入異動記錄
                    string vOptionString = "修改總務耗材請購發料單明細" + Environment.NewLine +
                                           "單號：[ " + vSheetNo + " ]" + Environment.NewLine +
                                           "項次：[ " + vItems + " ]" + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：修改" + Environment.NewLine +
                                           "ConsSheet_Order.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    int vTempINT;
                    string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    string vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    string vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    //複製舊資料到歷史檔
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionDate, ActionMan, SourceNo) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, GetDate(), '" + vLoginID + "', SourceNo " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //異動資料
                    TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_Edit");
                    TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_Edit");
                    TextBox eConsNo = (TextBox)fvConsSheetA_Detail.FindControl("eConsNo_Edit");
                    double vTempFloat;
                    vSQLStr_Temp = "update ConsSheetB set ConsNo = @ConsNo, Quantity = @Quantity, RemarkB = @RemarkB, ModifyDate = GetDate(), ModifyMan = @ModifyMan where SheetNoItems = @SheetNoItems ";
                    sdsConsSheetB_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.UpdateParameters.Clear();
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB.Text.Trim() != "") ? eRemarkB.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ConsNo", DbType.String, (eConsNo.Text.Trim() != "") ? eConsNo.Text.Trim() : String.Empty));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("Quantity", DbType.Double, double.TryParse(eQuantity.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsConsSheetB_Detail.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.Update();
                    gridConsSheetB_List.DataBind();
                    fvConsSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvConsSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = false;
                }
            }
        }

        protected void eConsName_Edit_TextChanged(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_Edit");
            if (eConsNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    Label eConsUnit1 = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_C1_Edit");
                    Label eConsUnit2 = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_C2_Edit");
                    string vConsNo = eConsNo.Text.Trim();
                    string vSQLStr_Temp = "select ConsUnit from Consumables where ConsNo = '" + vConsNo + "' ";
                    string vConsUnit = PF.GetValue(vConnStr, vSQLStr_Temp, "ConsUnit");
                    eConsUnit1.Text = vConsUnit;
                    eConsUnit2.Text = vConsUnit;
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        /// <summary>
        /// 批次匯入明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPatchNew_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eDepNo = (Label)fvConsSheetA_Detail.FindControl("eDepNo_List");
            if (eDepNo != null)
            {
                try
                {
                    string vDepNo = eDepNo.Text.Trim();
                    string vWStr_DepNo = (vDepNo != "") ? "   and a.DepNo = '" + vDepNo + "' " + Environment.NewLine : "";
                    string vSQLStr_Temp = "select a.SheetNo, b.Items, b.SheetNoItems, b.ConsNo, c.ConsName, b.Price, b.Quantity " + Environment.NewLine +
                                          "  from ConsSheetA as a left join ConsSheetB as b on b.SheetNo = a.SheetNo " + Environment.NewLine +
                                          "                       left join Consumables as c on c.ConsNo = b.ConsNo " + Environment.NewLine +
                                          " where a.SheetMode = 'BS' " + Environment.NewLine +
                                          "   and b.ItemStatus = '000' " + Environment.NewLine +
                                          vWStr_DepNo +
                                          " order by a.SheetNo, b.Items";
                    sdsOrderItems_List.SelectCommand = vSQLStr_Temp;
                    sdsOrderItems_List.Select(new DataSourceSelectArguments());
                    gridOrderItems_List.DataBind();
                    plShowData_B.Visible = false;
                    plShowData_C.Visible = true;
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        protected void eConsName_INS_TextChanged(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
            if (eConsNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                try
                {
                    Label eConsUnit1 = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_C1_INS");
                    Label eConsUnit2 = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_C2_INS");
                    string vConsNo = eConsNo.Text.Trim();
                    string vSQLStr_Temp = "select ConsUnit from Consumables where ConsNo = '" + vConsNo + "' ";
                    string vConsUnit = PF.GetValue(vConnStr, vSQLStr_Temp, "ConsUnit");
                    eConsUnit1.Text = vConsUnit;
                    eConsUnit2.Text = vConsUnit;
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        protected void bbOKC_Order_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                int vTempINT;
                string vSheetNo = eSheetNo.Text.Trim();
                string vSQLStr_Temp;
                string vMaxIndex;
                string vSheetNoItems;
                string vItems;
                string vSourceNo;
                try
                {
                    foreach (GridViewRow vCurrentRows in gridOrderItems_List.Rows)
                    {
                        CheckBox cbTemp = (CheckBox)vCurrentRows.FindControl("cbChoise");
                        if (cbTemp != null && cbTemp.Checked)
                        {
                            vSourceNo = gridOrderItems_List.DataKeys[vCurrentRows.RowIndex].Value.ToString().Trim();
                            vSQLStr_Temp = "select max(Items) MaxIndex from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                            vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                            vItems = Int32.TryParse(vMaxIndex, out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                            vSheetNoItems = vSheetNo.Trim() + vItems.Trim();
                            vSQLStr_Temp = "insert into ConsSheetB " + Environment.NewLine +
                                           "      (SheetNoItems, SheetNo, Items, ConsNo, ItemStatus, Quantity, ConsUnit, RemarkB, BuDate, BuMan, SourceNo) " + Environment.NewLine +
                                           "select '" + vSheetNoItems + "', '" + vSheetNo + "', '" + vItems + "', ConsNo, '000', Quantity, ConsUnit, RemarkB, GetDate(), " + Environment.NewLine +
                                           "       '" + vLoginID + "', SheetNoItems " + Environment.NewLine +
                                           "  from ConsSheetB " + Environment.NewLine +
                                           " where SheetNoItems = '" + vSourceNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            vSQLStr_Temp = "update ConsSheetB set ItemStatus = '102' where SheetNoItems = '" + vSourceNo + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        }
                    }
                    gridConsSheetB_List.DataBind();
                    plShowData_B.Visible = true;
                    plShowData_C.Visible = false;
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        protected void bbCancelC_Order_Click(object sender, EventArgs e)
        {
            plShowData_B.Visible = true;
            plShowData_C.Visible = false;
        }

        /// <summary>
        /// 全選
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSelectAll_Order_Click(object sender, EventArgs e)
        {
            foreach (GridViewRow vCurrentRows in gridOrderItems_List.Rows)
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
        protected void bbUnselAll_Order_Click(object sender, EventArgs e)
        {
            foreach (GridViewRow vCurrentRows in gridOrderItems_List.Rows)
            {
                CheckBox cbTemp = (CheckBox)vCurrentRows.FindControl("cbChoise");
                if (cbTemp != null) cbTemp.Checked = false;
            }
        }

        protected void bbOKB_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvConsSheetB_Detail.FindControl("eSheetNoB_INS");
            if (eSheetNo != null)
            {
                try
                {
                    int vTempINT;
                    string vSheetNo = eSheetNo.Text.Trim();
                    string vSQLStr_Temp = "select max(Items) MaxIndex from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                    string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    string vItems = Int32.TryParse(vMaxIndex, out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    string vSheetNoItems = vSheetNo.Trim() + vItems.Trim();
                    double vTempFloat;
                    TextBox eConsNo = (TextBox)fvConsSheetB_Detail.FindControl("eConsNo_INS");
                    string vConsNo = eConsNo.Text.Trim();
                    TextBox eQuantity = (TextBox)fvConsSheetB_Detail.FindControl("eQuantity_INS");
                    string vQuantityStr = eQuantity.Text.Trim();
                    TextBox eRemarkB = (TextBox)fvConsSheetB_Detail.FindControl("eRemarkB_INS");
                    string vRemarkB = eRemarkB.Text.Trim();
                    Label eConsUnit = (Label)fvConsSheetB_Detail.FindControl("eConsUnit_INS");
                    string vConsUnit = eConsUnit.Text.Trim();
                    vSQLStr_Temp = "insert into ConsSheetB " + Environment.NewLine +
                                   "      (SheetNoItems, SheetNo, Items, ConsNo, ItemStatus, Quantity, ConsUnit, RemarkB, BuDate, BuMan, QtyMode) " + Environment.NewLine +
                                   "values(@SheetNoItems, @SheetNo, @Items, @ConsNo, '000', @Quantity, @ConsUnit, @RemarkB, GetDate(), @BuMan, -1) ";
                    sdsConsSheetB_Detail.InsertCommand = vSQLStr_Temp;
                    sdsConsSheetB_Detail.InsertParameters.Clear();
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("ConsNo", DbType.String, (vConsNo != "") ? vConsNo : String.Empty));
                    sdsConsSheetB_Detail.InsertParameters.Add(new Parameter("Quantity", DbType.Double, double.TryParse(vQuantityStr, out vTempFloat) ? vTempFloat.ToString() : "0.0"));
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
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plSearch.Visible = true;
            plShowData_A.Visible = true;
            plShowData_B.Visible = true;
            plShowData_C.Visible = false;
        }
    }
}