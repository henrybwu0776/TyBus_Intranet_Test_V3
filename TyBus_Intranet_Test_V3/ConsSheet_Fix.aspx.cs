using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System.Data.SqlClient;
using System.IO;
using System.Globalization;
using System.Data;

namespace TyBus_Intranet_Test_V3
{
    public partial class ConsSheet_Fix : Page
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
                        eErrorMSG_Main.Text = "";
                        eErrorMSG_Main.Visible = false;
                        plSearch.Visible = true;
                        plShowData_A.Visible = true;
                        plShowData_B.Visible = true;
                        plShowData_C.Visible = true;
                        rvPrint.LocalReport.DataSources.Clear();
                        plReport.Visible = false;
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
        /// 取得主檔查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            string vResultStr;
            DateTime vTempDateS;
            DateTime vTempDateE;
            string vWStr_DepNo = (eDepNo_Search.Text.Trim() != "") ? "   and a.DepNo = '" + eDepNo_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_AssignMan = (eAssignMan_Search.Text.Trim() != "") ? "   and a.AssignMan = '" + eAssignMan_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_BuDate = ((DateTime.TryParse(eBuDateS_Search.Text.Trim(), out vTempDateS) == true) && (DateTime.TryParse(eBuDateE_Search.Text.Trim(), out vTempDateE) == true)) ?
                                  "   and a.BuDate between '" + vTempDateS.ToShortDateString() + "' and '" + vTempDateE.ToShortDateString() + Environment.NewLine :
                                  ((DateTime.TryParse(eBuDateS_Search.Text.Trim(), out vTempDateS) == true) && (DateTime.TryParse(eBuDateE_Search.Text.Trim(), out vTempDateE) != true)) ?
                                  "   and a.BuDate = '" + vTempDateS.ToShortDateString() + "' " + Environment.NewLine :
                                  ((DateTime.TryParse(eBuDateS_Search.Text.Trim(), out vTempDateS) != true) && (DateTime.TryParse(eBuDateE_Search.Text.Trim(), out vTempDateE) == true)) ?
                                  "   and a.BuDate = '" + vTempDateE.ToShortDateString() + "' " + Environment.NewLine : "";
            vResultStr = "select a.SheetNo, a.DepNo, d.[Name] DepName, a.AssignMan, e.[Name] AssignManName, a.BuDate " + Environment.NewLine +
                         "  from ConsSheetA a left join Department d on d.DepNo = a.DepNo " + Environment.NewLine +
                         "                    left join Employee e on e.EmpNo = a.AssignMan " + Environment.NewLine +
                         " where a.SheetMode = 'FS' " + Environment.NewLine +
                         vWStr_DepNo +
                         vWStr_AssignMan +
                         vWStr_BuDate +
                         " order by a.BuDate DESC, a.DepNo";
            return vResultStr;
        }

        /// <summary>
        /// 查詢主檔資料
        /// </summary>
        private void OpenData()
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            string vSelectStr = GetSelectStr();
            try
            {
                sdsSheetA_List.SelectCommand = vSelectStr;
                sdsSheetA_List.Select(new DataSourceSelectArguments());
                gridSheetA.DataBind();
            }
            catch (Exception eMessage)
            {
                eErrorMSG_A.Text = eMessage.Message;
                eErrorMSG_A.Visible = true;
            }
        }

        /// <summary>
        /// 單位代碼欄位內容異動
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eDepNo_Search_TextChanged(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo = eDepNo_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo + "' ";
            try
            {
                string vDepName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vDepName == "")
                {
                    vDepName = eDepNo_Search.Text.Trim();
                    vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '%" + vDepName + "%' order by DepNo DESC";
                    vDepNo = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
                }
                eDepNo_Search.Text = vDepNo;
                eDepName_Search.Text = vDepName;
            }
            catch (Exception eMessage)
            {
                eErrorMSG_A.Text = eMessage.Message;
                eErrorMSG_A.Visible = true;
            }
        }

        /// <summary>
        /// 申請人工號欄位內容異動
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eAssignMan_Search_TextChanged(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vAssignMan = eAssignMan_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
            try
            {
                string vAssignManName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vAssignManName == "")
                {
                    vAssignManName = eAssignMan_Search.Text.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] like '" + vAssignManName + "' order by Assumeday DESC";
                    vAssignMan = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eAssignMan_Search.Text = vAssignMan;
                eAssignManName_Search.Text = vAssignManName;
            }
            catch (Exception eMessage)
            {
                eErrorMSG_A.Text = eMessage.Message;
                eErrorMSG_A.Visible = true;
                throw;
            }
        }

        /// <summary>
        /// 查詢主檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        /// <summary>
        /// 離開功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvSheetA_Detail_DataBound(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDateURL;
            string vDateScript;
            Label eSheetNoA;
            TextBox eBuDate;
            TextBox eAssignMan;
            TextBox eDepNo;
            Label eAssignManName;
            Label eDepName;
            switch (fvSheetA_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    eSheetNoA = (Label)fvSheetA_Detail.FindControl("eSheetNoA_Edit");
                    if (eSheetNoA != null)
                    {
                        eBuDate = (TextBox)fvSheetA_Detail.FindControl("eBuDateA_Edit");
                        vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDate.ClientID;
                        vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                        eBuDate.Attributes["onClick"] = vDateScript;
                    }
                    break;
                case FormViewMode.Insert:
                    eSheetNoA = (Label)fvSheetA_Detail.FindControl("eSheetNoA_INS");
                    if (eSheetNoA != null)
                    {
                        eBuDate = (TextBox)fvSheetA_Detail.FindControl("eBuDateA_INS");
                        eBuDate.Text = vToday.ToShortDateString();
                        vDateURL = "InputDate_ChineseYears.aspx?TextBoxID=" + eBuDate.ClientID;
                        vDateScript = "window.open('" + vDateURL + "', '', 'height=315, width=350, Status=no, toolbar=no, menubar=no, location=no','')";
                        eBuDate.Attributes["onClick"] = vDateScript;
                        eAssignMan = (TextBox)fvSheetA_Detail.FindControl("eAssignMan_INS");
                        eAssignMan.Text = vLoginID;
                        eAssignManName = (Label)fvSheetA_Detail.FindControl("eAssignManName_INS");
                        eAssignManName.Text = vLoginName;
                        eDepNo = (TextBox)fvSheetA_Detail.FindControl("eDepNo_INS");
                        eDepNo.Text = vLoginDepNo;
                        eDepName = (Label)fvSheetA_Detail.FindControl("eDepName_INS");
                        eDepName.Text = vLoginDepName;
                    }
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 主檔確定修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKA_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvSheetA_Detail.FindControl("eSheetNoA_Edit");
            if (eSheetNo != null)
            {
                TextBox eBuDate = (TextBox)fvSheetA_Detail.FindControl("eBuDateA_Edit");
                TextBox eDepNo = (TextBox)fvSheetA_Detail.FindControl("eDepNo_Edit");
                TextBox eAssignMan = (TextBox)fvSheetA_Detail.FindControl("eAssignMan_Edit");
                TextBox eRemarkA = (TextBox)fvSheetA_Detail.FindControl("eRemarkA_Edit");
                DateTime vTempDate;
                string vSQLStr_Temp = "";
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vMaxIndex;
                string vNewIndex;
                int vTempINT;
                try
                {
                    string vOptionString = "修改總務請修單" + Environment.NewLine +
                                           "單號：[ " + eSheetNo.Text.Trim() + " ]" + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：修改" + Environment.NewLine +
                                           "ConsSheet_Fix.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份主檔資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                                   "       SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyDate, ModifyMan, DepNo, AssignMan, ActiveDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                                   "       SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyDate, ModifyMan, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + eSheetNo.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入新資料
                    vSQLStr_Temp = "update ConsSheetA " + Environment.NewLine +
                                   "   set BuDate = @BuDate, DepNo = @DepNo, AssignMan = @AssignMan, RemarkA = @RemarkA " + Environment.NewLine +
                                   " where SheetNo = @SheetNo";
                    sdsSheetA_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsSheetA_Detail.UpdateParameters.Clear();
                    sdsSheetA_Detail.UpdateParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.TryParse(eBuDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                    sdsSheetA_Detail.UpdateParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                    sdsSheetA_Detail.UpdateParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan.Text.Trim() != "") ? eAssignMan.Text.Trim() : String.Empty));
                    sdsSheetA_Detail.UpdateParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA.Text.Trim() != "") ? eRemarkA.Text.Trim() : String.Empty));
                    sdsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, eSheetNo.Text.Trim()));
                    sdsSheetA_Detail.Update();
                    gridSheetA.DataBind();
                    fvSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        protected void eDepNo_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvSheetA_Detail.FindControl("eDepNo_Edit");
            if (eDepNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eDepName = (Label)fvSheetA_Detail.FindControl("eDepName_Edit");
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

        protected void eAssignMan_Edit_TextChanged(object sender, EventArgs e)
        {
            TextBox eAssignMan = (TextBox)fvSheetA_Detail.FindControl("eAssignMan_Edit");
            if (eAssignMan != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eAssignManName = (Label)fvSheetA_Detail.FindControl("eAssignManName_Edit");
                string vAssignMan = eAssignMan.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
                string vAssignManName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vAssignManName == "")
                {
                    vAssignManName = eAssignMan.Text.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] like '%" + vAssignManName + "%' order by Assumeday DESC";
                    vAssignMan = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eAssignMan.Text = vAssignMan;
                eAssignManName.Text = vAssignManName;
            }
        }

        /// <summary>
        /// 主檔確定新增
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOKA_INS_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            DateTime vTempDate;
            string vFirstCode = (vToday.Year - 1911).ToString("D3") + vToday.Month.ToString("D2") + "FS";
            string vSQLStr_Temp;
            string vMaxIndex;
            string vNewIndex;
            int vTempINT;

            TextBox eDepNo = (TextBox)fvSheetA_Detail.FindControl("eDepNo_INS");
            if (eDepNo != null)
            {
                TextBox eBuDate = (TextBox)fvSheetA_Detail.FindControl("eBuDateA_INS");
                TextBox eAssignMan = (TextBox)fvSheetA_Detail.FindControl("eAssignMan_INS");
                TextBox eRemarkA = (TextBox)fvSheetA_Detail.FindControl("eRemarkA_INS");
                vSQLStr_Temp = "select max(SheetNo) MaxNo from ConsSheetA where SheetNo like '" + vFirstCode + "%' ";
                vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                vNewIndex = Int32.TryParse(vMaxIndex.Replace(vFirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                vSQLStr_Temp = "insert into ConsSheetA " + Environment.NewLine +
                               "      (SheetNo, SheetMode, BuDate, BuMan, AssignMan, RemarkA, DepNo) " + Environment.NewLine +
                               "values(@SheetNo, 'FS', @BuDate, @BuMan, @AssignMan, @RemarkA, @DepNo) ";
                sdsSheetA_Detail.InsertCommand = vSQLStr_Temp;
                sdsSheetA_Detail.InsertParameters.Clear();
                sdsSheetA_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vFirstCode + vNewIndex));
                sdsSheetA_Detail.InsertParameters.Add(new Parameter("BuDate", DbType.Date, DateTime.TryParse(eBuDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                sdsSheetA_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                sdsSheetA_Detail.InsertParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan.Text.Trim() != "") ? eAssignMan.Text.Trim() : String.Empty));
                sdsSheetA_Detail.InsertParameters.Add(new Parameter("RemarkA", DbType.String, (eRemarkA.Text.Trim() != "") ? eRemarkA.Text.Trim() : String.Empty));
                sdsSheetA_Detail.InsertParameters.Add(new Parameter("DepNo", DbType.String, (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                sdsSheetA_Detail.Insert();
                gridSheetA.DataBind();
                fvSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                fvSheetA_Detail.DataBind();
            }
        }

        protected void eDepNo_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eDepNo = (TextBox)fvSheetA_Detail.FindControl("eDepNo_INS");
            if (eDepNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eDepName = (Label)fvSheetA_Detail.FindControl("eDepName_INS");
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

        protected void eAssignMan_INS_TextChanged(object sender, EventArgs e)
        {
            TextBox eAssignMan = (TextBox)fvSheetA_Detail.FindControl("eAssignMan_INS");
            if (eAssignMan != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eAssignManName = (Label)fvSheetA_Detail.FindControl("eAssignManName_INS");
                string vAssignMan = eAssignMan.Text.Trim();
                string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan + "' ";
                string vAssignManName = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
                if (vAssignManName == "")
                {
                    vAssignManName = eAssignMan.Text.Trim();
                    vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] like '%" + vAssignManName + "%' order by Assumeday DESC";
                    vAssignMan = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
                }
                eAssignMan.Text = vAssignMan;
                eAssignManName.Text = vAssignManName;
            }
        }

        /// <summary>
        /// 列印請修單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrint_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            // 列印請修單
            DataTable dtPrintA;
            Label eSheetNo = (Label)fvSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSheetNo = eSheetNo.Text.Trim();
                string vSelectStr = "select a.SheetNo, a.DepNo, d.[Name] DepName, a.AssignMan, e.[Name] AssignManName, " + Environment.NewLine +
                                    "       convert(varchar(10), a.BuDate, 111) BuDate " + Environment.NewLine +
                                    "  from ConsSheetA a left join Department d on d.DepNo = a.DepNo " + Environment.NewLine +
                                    "                    left join Employee e on e.EmpNo = a.AssignMan " + Environment.NewLine +
                                    " where SheetNo = '" + vSheetNo + "' ";
                using (SqlConnection connSheetA = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daSheetA = new SqlDataAdapter(vSelectStr, connSheetA);
                    connSheetA.Open();
                    dtPrintA = new DataTable();
                    daSheetA.Fill(dtPrintA);
                }

                string vCompanyName = PF.GetValue(vConnStr, "select [Name] from [Custom] where [Code] = 'A000' and Types = 'O' ", "Name");
                string vReportName = "請購 (修) 單";
                string vReportPath = @"Report\ConsSheet_FixP.rdlc";
                ReportDataSource rdsPrintA = new ReportDataSource("ConsSheetA_FixP", dtPrintA);

                rvPrint.LocalReport.DataSources.Clear();
                rvPrint.LocalReport.ReportPath = vReportPath;
                rvPrint.LocalReport.DataSources.Add(rdsPrintA);
                rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", vCompanyName));
                rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                rvPrint.LocalReport.Refresh();
                plReport.Visible = true;
                plSearch.Visible = false;
                plShowData_A.Visible = false;
                plShowData_B.Visible = false;
                plShowData_C.Visible = false;
            }
        }

        protected void Unnamed_SubreportProcessing(object sender, SubreportProcessingEventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvSheetA_Detail.FindControl("eSheetNoA_List");
            string vSheetNo = eSheetNo.Text.Trim();
            DataTable dtPrintB;
            string vSelectStr = "select b.SheetNo, b.Items, b.ConsNo, c.ConsName, b.ConsUnit, d.ClassTxt ConsUnit_C, " + Environment.NewLine +
                                "       b.Quantity, b.RemarkB " + Environment.NewLine +
                                "  from ConsSheetB b left join Consumables c on c.ConsNo = b.ConsNo " + Environment.NewLine +
                                "                    left join DBDICB d on d.ClassNo = b.ConsUnit and d.FKey = '耗材庫存        CONSUMABLES     ConsUnit'" + Environment.NewLine +
                                " where b.SheetNo = '" + vSheetNo + "' order by Items ";
            using (SqlConnection connSheetB = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintB = new SqlDataAdapter(vSelectStr, connSheetB);
                connSheetB.Open();
                dtPrintB = new DataTable();
                daPrintB.Fill(dtPrintB);
            }
            ReportDataSource rdsPrintB = new ReportDataSource("ConsSheetB_FixP", dtPrintB);
            e.DataSources.Add(rdsPrintB);
        }

        /// <summary>
        /// 請修單作廢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbAbortA_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr=="")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo!=null)
            {
                try
                {
                    string vSheetNo = eSheetNo.Text.Trim();
                    int vTempINT;
                    //寫入異動記錄
                    string vOptionString = "修改總務請修單" + Environment.NewLine +
                                           "單號：[ " + eSheetNo.Text.Trim() + " ]" + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：作廢" + Environment.NewLine +
                                           "ConsSheet_Fix.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //複製資料到歷史檔
                    string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    string vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    string vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                                   "       SheetStatus, StatudDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNo, SheetMode, SupNo, BuDate, BuMan, Amount, TaxType, TaxRate, TaxAMT, TotalAmount, " + Environment.NewLine +
                                   "       SheetStatus, StatudDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, DepNo, AssignMan, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + vSheetNo + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    vSQLStr_Temp = "update ConsSheetA set SheetStatus = '999', StatusDate = GetDate() where SheetNo = @SheetNo";
                    sdsSheetA_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsSheetA_Detail.UpdateParameters.Clear();
                    sdsSheetA_Detail.UpdateParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsSheetA_Detail.Update();
                    gridSheetA.DataBind();
                    fvSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = false;
                    throw;
                }
            }
        }

        /// <summary>
        /// 刪除主檔資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eSheetNo = (Label)fvSheetA_Detail.FindControl("eSheetNoA_List");
            if (eSheetNo != null)
            {
                try
                {
                    string vSQLStr_Temp = "";
                    string vMaxIndex;
                    string vNewIndex;
                    string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    int vTempINT;
                    //先記錄異動
                    string vOptionString = "刪除總務請修單" + Environment.NewLine +
                                           "單號：[ " + eSheetNo.Text.Trim() + " ]" + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：刪除" + Environment.NewLine +
                                           "ConsSheet_Fix.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //檢查有沒有明細資料
                    vSQLStr_Temp = "select SheetNoItems from ConsSheetB where SheetNo = '" + eSheetNo.Text.Trim() + "' order by Items";
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        connTemp.Open();
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        while (drTemp.Read())
                        {
                            string vSheetNoItems = drTemp["SheetNoItems"].ToString().Trim();
                            //備份明細資料
                            vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                            vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                            vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                            vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                           "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                           "       QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionDate, ActionMan) " + Environment.NewLine +
                                           "select '" + vH_FirstCode + vNewIndex + "', 'DELA', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, RemarkB, " + Environment.NewLine +
                                           "       QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, GetDate(), '" + vLoginID + "' " + Environment.NewLine +
                                           "  from ConsSheetB " + Environment.NewLine +
                                           " where SheetNoItems = '" + vSheetNoItems + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                            //刪除明細
                            vSQLStr_Temp = "delete ConsSheetB where SheetNoItems = '" + vSheetNoItems + "' ";
                            PF.ExecSQL(vConnStr, vSQLStr_Temp);
                        }
                    }
                    //備份主檔資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetA_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetA_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNo, SheetMode, SupNo, BuDate, BuMan, AssignMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, ActionDate, ActionMan, SheetNote) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'DELA', SheetNo, SheetMode, SupNo, BuDate, BuMan, AssignMan, Amount, TaxType, TaxRate, " + Environment.NewLine +
                                   "       TaxAMT, TotalAmount, SheetStatus, StatusDate, PayDate, PayMode, RemarkA, ModifyMan, ModifyDate, " + Environment.NewLine +
                                   "       DepNo, GetDate(), '" + vLoginID + "', SheetNote " + Environment.NewLine +
                                   "  from ConsSheetA " + Environment.NewLine +
                                   " where SheetNo = '" + eSheetNo.Text.Trim() + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //刪除主檔資料
                    sdsSheetA_Detail.DeleteCommand = "delete from ConsSheetA where SheetNo = @SheetNo";
                    sdsSheetA_Detail.DeleteParameters.Clear();
                    sdsSheetA_Detail.DeleteParameters.Add(new Parameter("SheetNo", DbType.String, eSheetNo.Text.Trim()));
                    sdsSheetA_Detail.Delete();
                    gridSheetA.DataBind();
                    fvSheetA_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvSheetA_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message;
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        protected void fvSheetB_Detail_DataBound(object sender, EventArgs e)
        {
            Label eSheetNoItems;
            switch (fvSheetB_Detail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    break;
                case FormViewMode.Edit:
                    eSheetNoItems = (Label)fvSheetB_Detail.FindControl("eSheetNoItems_Edit");
                    if (eSheetNoItems != null)
                    {

                    }
                    break;
                case FormViewMode.Insert:
                    eSheetNoItems = (Label)fvSheetB_Detail.FindControl("eSheetNoItems_INS");
                    if (eSheetNoItems != null)
                    {
                        Label eSheetNoA = (Label)fvSheetA_Detail.FindControl("eSheetNoA_List");
                        Label eSheetNoB = (Label)fvSheetB_Detail.FindControl("eSheetNoB_INS");
                        eSheetNoB.Text = eSheetNoA.Text;
                    }
                    break;
                default:
                    break;
            }
        }

        protected void bbOKB_Edit_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            Label eSheetNoItems = (Label)fvSheetB_Detail.FindControl("eSheetNoItems_Edit");
            if (eSheetNoItems != null)
            {
                Label eSheetNo = (Label)fvSheetB_Detail.FindControl("eSheetNoB_Edit");
                Label eItems = (Label)fvSheetB_Detail.FindControl("eItems_Edit");
                TextBox eRemarkB = (TextBox)fvSheetB_Detail.FindControl("eRemarkB_Edit");
                string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                string vSQLStr_Temp;
                string vMaxIndex;
                string vNewIndex;
                int vTempINT;
                try
                {
                    //寫入異動記錄
                    string vSheetNo = eSheetNo.Text.Trim();
                    string vItems = eItems.Text.Trim();
                    string vSheetNoItems = eSheetNoItems.Text.Trim();
                    string vOptionString = "修改總務請修單明細" + Environment.NewLine +
                                           "單號：[ " + vSheetNo + " ] " + Environment.NewLine +
                                           "項次：[ " + vItems + " ] " + Environment.NewLine +
                                           "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                           "異動種類：修改" + Environment.NewLine +
                                           "ConsSheet_Fix.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                    //備份明細資料
                    vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionMan, ActionDate) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'EDIT', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, '" + vLoginID + "', GetDate() " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入修改資料
                    vSQLStr_Temp = "update ConsSheetB set RemarkB = @RemarkB, ModifyMan = @ModifyMan, ModifyDate = GetDate() where SheetNoItems = @SheetNoItems ";
                    sdsSheetB_Detail.UpdateCommand = vSQLStr_Temp;
                    sdsSheetB_Detail.UpdateParameters.Clear();
                    sdsSheetB_Detail.UpdateParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB.Text.Trim() != "") ? eRemarkB.Text.Trim() : String.Empty));
                    sdsSheetB_Detail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, vLoginID));
                    sdsSheetB_Detail.UpdateParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsSheetB_Detail.Update();
                    gridSheetB_List.DataBind();
                    fvSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = false;
                }
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
            TextBox eRemarkB = (TextBox)fvSheetB_Detail.FindControl("eRemarkB_INS");
            if (eRemarkB != null)
            {
                int vTempINT;
                Label eSheetNoB = (Label)fvSheetB_Detail.FindControl("eSheetNoB_INS");
                string vSheetNo = eSheetNoB.Text.Trim();
                string vSQLStr_Temp = "select max(Items) MaxIndex from ConsSheetB where SheetNo = '" + vSheetNo + "' ";
                string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                string vItems = Int32.TryParse(vMaxIndex, out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                string vSheetNoItems = vSheetNo + vItems;
                try
                {
                    vSQLStr_Temp = "insert into ConsSheetB " + Environment.NewLine +
                                   "       (SheetNoItems, SheetNo, Items, ConsNo, Quantity, RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ConsUnit) " + Environment.NewLine +
                                   "values (@SheetNoItems, @SheetNo, @Items, 'F-02-0007', 1, @RemarkB, 1, '000', @BuMan, GetDate(), '0')"; //因為是報修，料號統一使用 "其他" 分項
                    sdsSheetB_Detail.InsertCommand = vSQLStr_Temp;
                    sdsSheetB_Detail.InsertParameters.Clear();
                    sdsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsSheetB_Detail.InsertParameters.Add(new Parameter("SheetNo", DbType.String, vSheetNo));
                    sdsSheetB_Detail.InsertParameters.Add(new Parameter("Items", DbType.String, vItems));
                    sdsSheetB_Detail.InsertParameters.Add(new Parameter("RemarkB", DbType.String, (eRemarkB.Text.Trim() != "") ? eRemarkB.Text.Trim() : String.Empty));
                    sdsSheetB_Detail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                    sdsSheetB_Detail.Insert();
                    gridSheetB_List.DataBind();
                    fvSheetB_Detail.ChangeMode(FormViewMode.ReadOnly);
                    fvSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        protected void bbAbortB_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            Label eSheetNoItems = (Label)fvSheetB_Detail.FindControl("eSheetNoItems_List");
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (eSheetNoItems != null)
            {
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                Label eSheetNo = (Label)fvSheetB_Detail.FindControl("eSheetNoB_List");
                Label eItems = (Label)fvSheetB_Detail.FindControl("eItems_List");
                string vSheetNo = eSheetNo.Text.Trim();
                string vItems = eItems.Text.Trim();
                int vTempINT;
                //寫入異動記錄
                string vOptionString = "異動總務請修單明細" + Environment.NewLine +
                                       "單號：[ " + vSheetNo + " ] " + Environment.NewLine +
                                       "項次：[ " + vItems + " ] " + Environment.NewLine +
                                       "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                       "異動種類：明細作廢" + Environment.NewLine +
                                       "ConsSheet_Fix.aspx";
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                try
                {
                    //複製舊內容到歷史檔
                    string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    string vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History ";
                    string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    string vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionDate, ActionMan) " + Environment.NewLine +
                                   "select '" + vNewIndex + "', 'EDIT', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionDate, ActionMan " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //執行明細作廢
                    vSQLStr_Temp = "update ConsSheetB set ItemStatus = '999' where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    gridSheetB_List.DataBind();
                    fvSheetB_Detail.DataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        protected void bbDeleteB_List_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            Label eSheetNoItems = (Label)fvSheetB_Detail.FindControl("eSheetNoItems_List");
            if (eSheetNoItems != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                Label eSheetNo = (Label)fvSheetB_Detail.FindControl("eSheetNoB_List");
                Label eItems = (Label)fvSheetB_Detail.FindControl("eItems_List");
                string vSheetNoItems = eSheetNoItems.Text.Trim();
                string vSheetNo = eSheetNo.Text.Trim();
                string vItems = eItems.Text.Trim();
                //寫入異動檔
                string vOptionString = "刪除總務請修單明細" + Environment.NewLine +
                                       "單號：[ " + vSheetNo + " ] " + Environment.NewLine +
                                       "項次：[ " + vItems + " ] " + Environment.NewLine +
                                       "異動日期：" + vToday.ToShortDateString() + Environment.NewLine +
                                       "異動種類：刪除" + Environment.NewLine +
                                       "ConsSheet_Fix.aspx";
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vOptionString);
                //複製明細資料到備份檔
                try
                {
                    string vH_FirstCode = vToday.Year.ToString("D4") + vToday.Month.ToString("D2");
                    string vSQLStr_Temp = "select max(H_Index) MaxIndex from ConsSheetB_History where H_Index like '" + vH_FirstCode + "%' ";
                    string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxIndex");
                    int vTempINT;
                    string vNewIndex = Int32.TryParse(vMaxIndex.Replace(vH_FirstCode, ""), out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                    vSQLStr_Temp = "insert into ConsSheetB_History " + Environment.NewLine +
                                   "      (H_Index, ActionMode, SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionDate, ActionMan) " + Environment.NewLine +
                                   "select '" + vH_FirstCode + vNewIndex + "', 'DELB', SheetNoItems, SheetNo, Items, ConsNo, Price, Quantity, Amount, " + Environment.NewLine +
                                   "       RemarkB, QtyMode, ItemStatus, BuMan, BuDate, ModifyMan, ModifyDate, ConsUnit, ActionDate, ActionMan " + Environment.NewLine +
                                   "  from ConsSheetB " + Environment.NewLine +
                                   " where SheetNoItems = '" + vSheetNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //刪除明細資料
                    vSQLStr_Temp = "delete ConsSheetB where SheetNoItems = @SheetNoItems ";
                    sdsSheetB_Detail.DeleteCommand = vSQLStr_Temp;
                    sdsSheetB_Detail.DeleteParameters.Clear();
                    sdsSheetB_Detail.DeleteParameters.Add(new Parameter("SheetNoItems", DbType.String, vSheetNoItems));
                    sdsSheetB_Detail.Delete();
                    gridSheetB_List.DataBind();
                    fvSheetB_Detail.DataBind();
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
            plSearch.Visible = true;
            plShowData_A.Visible = true;
            plShowData_B.Visible = true;
            plShowData_C.Visible = true;
            rvPrint.LocalReport.DataSources.Clear();
            plReport.Visible = false;
        }
    }
}