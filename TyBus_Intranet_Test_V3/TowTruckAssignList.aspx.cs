using Amaterasu_Function;
using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Threading;
using System.Web.UI.WebControls;

namespace TyBus_Intranet_Test_V3
{
    public partial class TowTruckAssignList : System.Web.UI.Page
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
                    string vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_S_Search.ClientID;
                    string vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_S_Search.Attributes["onClick"] = vDateScript_Temp;

                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_E_Search.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_E_Search.Attributes["onClick"] = vDateScript_Temp;

                    if (!IsPostBack)
                    {
                        plPrint.Visible = false;
                        plShowData.Visible = true;
                        plSearch.Visible = true;
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
        /// 取回顯示清單用的查詢語法
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            DateTime eCaseDate_S = (eCaseDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_S_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_S_Str = eCaseDate_S.Year.ToString() + "/" + eCaseDate_S.ToString("MM/dd");
            DateTime eCaseDate_E = (eCaseDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_E_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_E_Str = eCaseDate_E.Year.ToString() + "/" + eCaseDate_E.ToString("MM/dd");
            string vWStr_CaseDate = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate between '" + vCaseDate_S_Str + "' and '" + vCaseDate_E_Str + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? "   and t.CaseDate = '" + vCaseDate_S_Str + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate = '" + vCaseDate_E_Str + "' " + Environment.NewLine : "";
            string vWStr_DepNo_Car = ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? "   and t.DepNo_Car between '" + eDepNo_Car_S_Search.Text.Trim() + "' and '" + eDepNo_Car_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() == "")) ? "   and t.DepNo_Car = '" + eDepNo_Car_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Car_S_Search.Text.Trim() == "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? "   and t.DepNo_Car = '" + eDepNo_Car_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo_Man = (eDepNo_Man_Search.Text.Trim() != "") ? "   and t.DepNo_Man = '" + eDepNo_Man_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_AssignMan = (eAssignMan_Search.Text.Trim() != "") ? "   and (t.AssignMan = '" + eAssignMan_Search.Text.Trim() + "' or t.FirstSupportMan = '" + eAssignMan_Search.Text.Trim() + "' or '" + eAssignMan_Search.Text.Trim() + "') " + Environment.NewLine : "";
            string vWStr_CarID = (eCarID_Search.Text.Trim() != "") ? "   and t.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? "   and t.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vSelStr = "SELECT CaseType, (SELECT CLASSTXT FROM DBDICB WHERE (FKEY = '拖吊車叫用記錄表TowTruckAssign  CaseType') AND (CLASSNO = t.CaseType)) AS CaseType_C, CaseNo, CaseDate, Car_ID, " + Environment.NewLine +
                             "       t.DepNo_Car, (SELECT [NAME] FROM DEPARTMENT WHERE (DEPNO = t.DepNo_Car)) AS DepName_Car, " + Environment.NewLine +
                             "       CasePosition, ParkingPosition, CaseTime, " + Environment.NewLine +
                             "       t.AssignMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.AssignMan)) AS AssignManName, " + Environment.NewLine +
                             "       t.FirstSupportMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.FirstSupportMan)) AS FirstSupportManName, " + Environment.NewLine +
                             "       t.SecondSupportMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.SecondSupportMan)) AS SecondSupportManName, " + Environment.NewLine +
                             //"       Determination " + Environment.NewLine +
                             "       Determination, IsChecked " + Environment.NewLine +
                             "  FROM TowTruckAssignList AS t " + Environment.NewLine +
                             " WHERE isnull(CaseNo, '') <> '' " + Environment.NewLine +
                             vWStr_AssignMan + vWStr_CarID + vWStr_CaseDate + vWStr_DepNo_Car + vWStr_DepNo_Man + vWStr_Driver +
                             " order by CaseDate DESC ";
            return vSelStr;
        }

        /// <summary>
        /// 取回列印用的查詢語法 (只列出有確認的車單)
        /// </summary>
        /// <returns></returns>
        private string GetPrintSelectStr()
        {
            DateTime eCaseDate_S = (eCaseDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_S_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_S_Str = eCaseDate_S.Year.ToString() + "/" + eCaseDate_S.ToString("MM/dd");
            DateTime eCaseDate_E = (eCaseDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_E_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_E_Str = eCaseDate_E.Year.ToString() + "/" + eCaseDate_E.ToString("MM/dd");
            string vWStr_CaseDate = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate between '" + vCaseDate_S_Str + "' and '" + vCaseDate_E_Str + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? "   and t.CaseDate = '" + vCaseDate_S_Str + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate = '" + vCaseDate_E_Str + "' " + Environment.NewLine : "";
            string vWStr_DepNo_Car = ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? "   and t.DepNo_Car between '" + eDepNo_Car_S_Search.Text.Trim() + "' and '" + eDepNo_Car_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() == "")) ? "   and t.DepNo_Car = '" + eDepNo_Car_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Car_S_Search.Text.Trim() == "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? "   and t.DepNo_Car = '" + eDepNo_Car_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo_Man = (eDepNo_Man_Search.Text.Trim() != "") ? "   and t.DepNo_Man = '" + eDepNo_Man_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_AssignMan = (eAssignMan_Search.Text.Trim() != "") ? "   and (t.AssignMan = '" + eAssignMan_Search.Text.Trim() + "' or t.FirstSupportMan = '" + eAssignMan_Search.Text.Trim() + "' or '" + eAssignMan_Search.Text.Trim() + "') " + Environment.NewLine : "";
            string vWStr_CarID = (eCarID_Search.Text.Trim() != "") ? "   and t.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? "   and t.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vSelStr = "SELECT t.CaseNo, t.CaseDate, t.Car_ID, t.Driver, (select [Name] from Employee where EmpNo = t.Driver) DriverName, " + Environment.NewLine +
                             "       t.DepNo_Car, (SELECT [NAME] FROM DEPARTMENT WHERE (DEPNO = t.DepNo_Car)) AS DepName_Car, " + Environment.NewLine +
                             "       c.Car_Class, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = c.Car_Class) AND (FKEY = '車輛資料作業    Car_infoA       CAR_CLASS')) AS Car_ClassName, " + Environment.NewLine +
                             "       c.point, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = c.point) AND (FKEY = '車輛資料作業    Car_infoA       POINT')) AS PointName, " + Environment.NewLine +
                             "       DATEDIFF(Month, c.getlicdate, GETDATE()) / 12 AS CarAge_Year, DATEDIFF(Month, c.getlicdate, GETDATE()) % 12 AS CarAge_Month, " + Environment.NewLine +
                             "       c.getlicdate, t.CasePosition, t.ParkingPosition, t.CaseTime, " + Environment.NewLine +
                             "       t.DepNo_Man, (select [Name] from Department where DepNo = t.DepNo_Man) DepName_Man, " + Environment.NewLine +
                             "       t.AssignMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.AssignMan)) AS AssignManName, " + Environment.NewLine +
                             "       t.FirstSupportMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.FirstSupportMan)) AS FirstSupportManName, " + Environment.NewLine +
                             "       t.SecondSupportMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.SecondSupportMan)) AS SecondSupportManName, " + Environment.NewLine +
                             "       t.Determination, t.FaultParts, t.FaultReason, t.Dispose, t.FollowUp, t.Improvements, t.Remark, " + Environment.NewLine +
                             "       t.BuDate, t.BuMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.BuMan)) AS BuManName, " + Environment.NewLine +
                             "       t.ModifyDate, t.ModifyMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.ModifyMan)) AS ModifyManName, " + Environment.NewLine +
                             //"       t.CarKMs, t.TowingCost, t.FixFee, t.LastServiceDate, t.DamageAnalysis, (select ClassTxt from DBDICB where FKey = '拖吊車叫用記錄表TowTruckAssign  DamageAnalysis' and ClassNo = t.DamageAnalysis) DamageAnalysis_C " + Environment.NewLine +
                             "       t.CarKMs, t.TowingCost, t.FixFee, t.LastServiceDate, t.IsChecked, " + Environment.NewLine +
                             "       t.DamageAnalysis, (select ClassTxt from DBDICB where FKey = '拖吊車叫用記錄表TowTruckAssign  DamageAnalysis' and ClassNo = t.DamageAnalysis) DamageAnalysis_C " + Environment.NewLine +
                             "  FROM TowTruckAssignList AS t LEFT OUTER JOIN Car_infoA AS c ON c.Car_ID = t.Car_ID " + Environment.NewLine +
                             " WHERE isnull(CaseNo, '') <> '' " + Environment.NewLine +
                             "   AND t.IsChecked = 'V' " + Environment.NewLine +
                             vWStr_AssignMan + vWStr_CarID + vWStr_CaseDate + vWStr_DepNo_Car + vWStr_DepNo_Man + vWStr_Driver +
                             " order by CaseDate ";
            return vSelStr;
        }

        /// <summary>
        /// 取回列印用的查詢語法 (列出所有車單)
        /// </summary>
        /// <returns></returns>
        private string GetPrintSelectStr_2()
        {
            DateTime eCaseDate_S = (eCaseDate_S_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_S_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_S_Str = eCaseDate_S.Year.ToString() + "/" + eCaseDate_S.ToString("MM/dd");
            DateTime eCaseDate_E = (eCaseDate_E_Search.Text.Trim() != "") ? DateTime.Parse(eCaseDate_E_Search.Text.Trim()) : DateTime.Today;
            string vCaseDate_E_Str = eCaseDate_E.Year.ToString() + "/" + eCaseDate_E.ToString("MM/dd");
            string vWStr_CaseDate = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate between '" + vCaseDate_S_Str + "' and '" + vCaseDate_E_Str + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? "   and t.CaseDate = '" + vCaseDate_S_Str + "' " + Environment.NewLine :
                                    ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "   and t.CaseDate = '" + vCaseDate_E_Str + "' " + Environment.NewLine : "";
            string vWStr_DepNo_Car = ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? "   and t.DepNo_Car between '" + eDepNo_Car_S_Search.Text.Trim() + "' and '" + eDepNo_Car_E_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() == "")) ? "   and t.DepNo_Car = '" + eDepNo_Car_S_Search.Text.Trim() + "' " + Environment.NewLine :
                                     ((eDepNo_Car_S_Search.Text.Trim() == "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? "   and t.DepNo_Car = '" + eDepNo_Car_E_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_DepNo_Man = (eDepNo_Man_Search.Text.Trim() != "") ? "   and t.DepNo_Man = '" + eDepNo_Man_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_AssignMan = (eAssignMan_Search.Text.Trim() != "") ? "   and (t.AssignMan = '" + eAssignMan_Search.Text.Trim() + "' or t.FirstSupportMan = '" + eAssignMan_Search.Text.Trim() + "' or '" + eAssignMan_Search.Text.Trim() + "') " + Environment.NewLine : "";
            string vWStr_CarID = (eCarID_Search.Text.Trim() != "") ? "   and t.Car_ID = '" + eCarID_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vWStr_Driver = (eDriver_Search.Text.Trim() != "") ? "   and t.Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";

            string vSelStr = "SELECT t.CaseNo, t.CaseDate, t.Car_ID, t.Driver, (select [Name] from Employee where EmpNo = t.Driver) DriverName, " + Environment.NewLine +
                             "       t.DepNo_Car, (SELECT [NAME] FROM DEPARTMENT WHERE (DEPNO = t.DepNo_Car)) AS DepName_Car, " + Environment.NewLine +
                             "       c.Car_Class, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = c.Car_Class) AND (FKEY = '車輛資料作業    Car_infoA       CAR_CLASS')) AS Car_ClassName, " + Environment.NewLine +
                             "       c.point, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = c.point) AND (FKEY = '車輛資料作業    Car_infoA       POINT')) AS PointName, " + Environment.NewLine +
                             "       DATEDIFF(Month, c.getlicdate, GETDATE()) / 12 AS CarAge_Year, DATEDIFF(Month, c.getlicdate, GETDATE()) % 12 AS CarAge_Month, " + Environment.NewLine +
                             "       c.getlicdate, t.CasePosition, t.ParkingPosition, t.CaseTime, " + Environment.NewLine +
                             "       t.DepNo_Man, (select [Name] from Department where DepNo = t.DepNo_Man) DepName_Man, " + Environment.NewLine +
                             "       t.AssignMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.AssignMan)) AS AssignManName, " + Environment.NewLine +
                             "       t.FirstSupportMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.FirstSupportMan)) AS FirstSupportManName, " + Environment.NewLine +
                             "       t.SecondSupportMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.SecondSupportMan)) AS SecondSupportManName, " + Environment.NewLine +
                             "       t.Determination, t.FaultParts, t.FaultReason, t.Dispose, t.FollowUp, t.Improvements, t.Remark, " + Environment.NewLine +
                             "       t.BuDate, t.BuMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.BuMan)) AS BuManName, " + Environment.NewLine +
                             "       t.ModifyDate, t.ModifyMan, (SELECT [NAME] FROM EMPLOYEE WHERE (EMPNO = t.ModifyMan)) AS ModifyManName, " + Environment.NewLine +
                             //"       t.CarKMs, t.TowingCost, t.FixFee, t.LastServiceDate, t.DamageAnalysis, (select ClassTxt from DBDICB where FKey = '拖吊車叫用記錄表TowTruckAssign  DamageAnalysis' and ClassNo = t.DamageAnalysis) DamageAnalysis_C " + Environment.NewLine +
                             "       t.CarKMs, t.TowingCost, t.FixFee, t.LastServiceDate, t.IsChecked, " + Environment.NewLine +
                             "       t.DamageAnalysis, (select ClassTxt from DBDICB where FKey = '拖吊車叫用記錄表TowTruckAssign  DamageAnalysis' and ClassNo = t.DamageAnalysis) DamageAnalysis_C " + Environment.NewLine +
                             "  FROM TowTruckAssignList AS t LEFT OUTER JOIN Car_infoA AS c ON c.Car_ID = t.Car_ID " + Environment.NewLine +
                             " WHERE isnull(CaseNo, '') <> '' " + Environment.NewLine +
                             vWStr_AssignMan + vWStr_CarID + vWStr_CaseDate + vWStr_DepNo_Car + vWStr_DepNo_Man + vWStr_Driver +
                             " order by CaseDate ";
            return vSelStr;
        }

        /// <summary>
        /// 開啟資料庫取回資料
        /// </summary>
        private void OpenData()
        {
            string vSelStr = GetSelectStr();
            sdsTowTruckAssignList.SelectCommand = vSelStr;
            gridTowTruckAssignList.DataBind();
        }

        protected void eDepNo_Car_S_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_Car_S_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '" + vDepName_Temp + "%' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Car_S_Search.Text = vDepNo_Temp.Trim();
            eDepName_Car_S_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDepNo_Car_E_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_Car_E_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '" + vDepName_Temp + "%' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Car_E_Search.Text = vDepNo_Temp.Trim();
            eDepName_Car_E_Search.Text = vDepName_Temp.Trim();
        }

        protected void eDepNo_Man_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDepNo_Temp = eDepNo_Man_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp.Trim() + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp.Trim() == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select top 1 DepNo from Department where [Name] like '" + vDepName_Temp + "%' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Man_Search.Text = vDepNo_Temp.Trim();
            eDepName_Man_Search.Text = vDepName_Temp.Trim();
        }

        protected void eAssignMan_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vAssignMan_Temp = eAssignMan_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan_Temp.Trim() + "' ";
            string vAssignName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vAssignName_Temp.Trim() == "")
            {
                vAssignName_Temp = vAssignMan_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vAssignName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC ";
                vAssignMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eAssignMan_Search.Text = vAssignMan_Temp.Trim();
            eAssignName_Search.Text = vAssignName_Temp.Trim();
        }

        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDriver_Temp = eDriver_Search.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp.Trim() + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp.Trim() == "")
            {
                vDriverName_Temp = vDriver_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC ";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Search.Text = vDriver_Temp.Trim();
            eDriverName_Search.Text = vDriverName_Temp.Trim();
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            OpenData();
        }

        /// <summary>
        /// 預覽拖吊車記錄清冊
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintList_Click(object sender, EventArgs e)
        {
            if ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != ""))
            {
                //2023.07.14 改用取回所有資料的語法
                //string vSelStr = GetPrintSelectStr();
                string vSelStr = GetPrintSelectStr_2();
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daPrintPoint = new SqlDataAdapter(vSelStr, connTempPrint);
                    connTempPrint.Open();
                    DataTable dtPrintPoint = new DataTable("TowTruckAssignListP");
                    daPrintPoint.Fill(dtPrintPoint);
                    if (dtPrintPoint.Rows.Count > 0)
                    {
                        string vReportName = "桃園汽車客運  " +
                                             (DateTime.Parse(eCaseDate_S_Search.Text.Trim()).Year - 1911).ToString() + " 年 " +
                                             (DateTime.Parse(eCaseDate_S_Search.Text.Trim()).Month).ToString() + " 月 ";
                        ReportDataSource rdsPrint = new ReportDataSource("TowTruckAssignListP", dtPrintPoint);

                        rvPrint.LocalReport.DataSources.Clear();
                        rvPrint.LocalReport.ReportPath = @"Report\TowTruckAssignListP.rdlc";
                        rvPrint.LocalReport.DataSources.Add(rdsPrint);
                        rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName_Start", vReportName));
                        rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName_End", "　拖吊車紀錄"));
                        rvPrint.LocalReport.Refresh();
                        plPrint.Visible = true;
                        plShowData.Visible = false;
                        plSearch.Visible = false;

                        string vCaseDateStr = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "從 " + eCaseDate_S_Search.Text.Trim() + " 起至 " + eCaseDate_E_Search.Text.Trim() + " 止" :
                                              ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? eCaseDate_S_Search.Text.Trim() :
                                              ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? eCaseDate_E_Search.Text.Trim() : "";
                        string vDepNoCarStr = ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? "從 " + eDepNo_Car_S_Search.Text.Trim() + " 起至 " + eDepNo_Car_E_Search.Text.Trim() + " 止" :
                                              ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() == "")) ? eDepNo_Car_S_Search.Text.Trim() :
                                              ((eDepNo_Car_S_Search.Text.Trim() == "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? eDepNo_Car_E_Search.Text.Trim() : "全部";
                        string vDepNoManStr = (eDepNo_Man_Search.Text.Trim() != "") ? eDepNo_Man_Search.Text.Trim() : "全部";
                        string vAssignManStr = (eAssignMan_Search.Text.Trim() != "") ? eAssignMan_Search.Text.Trim() : "全部";
                        string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "";
                        string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "";
                        string vRecordNote = "匯出檔案_拖吊車紀錄清冊" + Environment.NewLine +
                                             "TowTruckAssignList.aspx" + Environment.NewLine +
                                             "拖吊日期：" + vCaseDateStr + Environment.NewLine +
                                             "車輛所屬站別：" + vDepNoCarStr + Environment.NewLine +
                                             "檢修人員單位：" + vDepNoManStr + Environment.NewLine +
                                             "檢修人員：" + vAssignManStr + Environment.NewLine +
                                             "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                             "駕駛員：" + vDriverStr;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇查詢起迄日期！')");
                Response.Write("</" + "Script>");
                eCaseDate_S_Search.Focus();
            }
        }

        /// <summary>
        /// 預覽拖吊內容記錄表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintDetail_Click(object sender, EventArgs e)
        {
            string vSelStr = GetPrintSelectStr_2();
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daPrintPoint = new SqlDataAdapter(vSelStr, connTempPrint);
                connTempPrint.Open();
                DataTable dtPrintPoint = new DataTable("TowTruckAssignListP");
                daPrintPoint.Fill(dtPrintPoint);
                if (dtPrintPoint.Rows.Count > 0)
                {
                    ReportDataSource rdsPrint = new ReportDataSource("TowTruckAssignListP", dtPrintPoint);

                    rvPrint.LocalReport.DataSources.Clear();
                    rvPrint.LocalReport.ReportPath = @"Report\TowTruckAssignRecordP.rdlc";
                    rvPrint.LocalReport.DataSources.Add(rdsPrint);
                    rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                    rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", "拖吊車記錄表"));
                    rvPrint.LocalReport.Refresh();
                    plPrint.Visible = true;
                    plShowData.Visible = false;
                    plSearch.Visible = false;

                    string vCaseDateStr = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "從 " + eCaseDate_S_Search.Text.Trim() + " 起至 " + eCaseDate_E_Search.Text.Trim() + " 止" :
                                          ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? eCaseDate_S_Search.Text.Trim() :
                                          ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? eCaseDate_E_Search.Text.Trim() : "";
                    string vDepNoCarStr = ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? "從 " + eDepNo_Car_S_Search.Text.Trim() + " 起至 " + eDepNo_Car_E_Search.Text.Trim() + " 止" :
                                          ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() == "")) ? eDepNo_Car_S_Search.Text.Trim() :
                                          ((eDepNo_Car_S_Search.Text.Trim() == "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? eDepNo_Car_E_Search.Text.Trim() : "全部";
                    string vDepNoManStr = (eDepNo_Man_Search.Text.Trim() != "") ? eDepNo_Man_Search.Text.Trim() : "全部";
                    string vAssignManStr = (eAssignMan_Search.Text.Trim() != "") ? eAssignMan_Search.Text.Trim() : "全部";
                    string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "";
                    string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "";
                    string vRecordNote = "匯出檔案_拖吊內容記錄表" + Environment.NewLine +
                                         "TowTruckAssignList.aspx" + Environment.NewLine +
                                         "拖吊日期：" + vCaseDateStr + Environment.NewLine +
                                         "車輛所屬站別：" + vDepNoCarStr + Environment.NewLine +
                                         "檢修人員單位：" + vDepNoManStr + Environment.NewLine +
                                         "檢修人員：" + vAssignManStr + Environment.NewLine +
                                         "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                         "駕駛員：" + vDriverStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                }
            }
        }

        /// <summary>
        /// 列印拖吊維修費用月統計表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbPrintCostList_Click(object sender, EventArgs e)
        {
            if ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != ""))
            {
                //2023.07.14 改用取回所有資料的語法
                //string vSelStr = GetPrintSelectStr();
                string vSelStr = GetPrintSelectStr_2();
                string vReportName = (DateTime.Parse(eCaseDate_S_Search.Text.Trim()).Year - 1911).ToString() + " 年 " + DateTime.Parse(eCaseDate_S_Search.Text.Trim()).Month.ToString("D2") + " 月 拖吊維修費用月統計表";

                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                using (SqlConnection connTempPrint = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daPrintPoint = new SqlDataAdapter(vSelStr, connTempPrint);
                    connTempPrint.Open();
                    DataTable dtPrintPoint = new DataTable("TowTruckAssignListP");
                    daPrintPoint.Fill(dtPrintPoint);
                    if (dtPrintPoint.Rows.Count > 0)
                    {
                        ReportDataSource rdsPrint = new ReportDataSource("TowTruckAssignListP", dtPrintPoint);

                        rvPrint.LocalReport.DataSources.Clear();
                        rvPrint.LocalReport.ReportPath = @"Report\TowTruckCostListP.rdlc";
                        rvPrint.LocalReport.DataSources.Add(rdsPrint);
                        rvPrint.LocalReport.SetParameters(new ReportParameter("CompanyName", "桃園汽車客運股份有限公司"));
                        rvPrint.LocalReport.SetParameters(new ReportParameter("ReportName", vReportName));
                        rvPrint.LocalReport.Refresh();
                        plPrint.Visible = true;
                        plShowData.Visible = false;
                        plSearch.Visible = false;

                        string vCaseDateStr = ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() != "")) ? "從 " + eCaseDate_S_Search.Text.Trim() + " 起至 " + eCaseDate_E_Search.Text.Trim() + " 止" :
                                              ((eCaseDate_S_Search.Text.Trim() != "") && (eCaseDate_E_Search.Text.Trim() == "")) ? eCaseDate_S_Search.Text.Trim() :
                                              ((eCaseDate_S_Search.Text.Trim() == "") && (eCaseDate_E_Search.Text.Trim() != "")) ? eCaseDate_E_Search.Text.Trim() : "";
                        string vDepNoCarStr = ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? "從 " + eDepNo_Car_S_Search.Text.Trim() + " 起至 " + eDepNo_Car_E_Search.Text.Trim() + " 止" :
                                              ((eDepNo_Car_S_Search.Text.Trim() != "") && (eDepNo_Car_E_Search.Text.Trim() == "")) ? eDepNo_Car_S_Search.Text.Trim() :
                                              ((eDepNo_Car_S_Search.Text.Trim() == "") && (eDepNo_Car_E_Search.Text.Trim() != "")) ? eDepNo_Car_E_Search.Text.Trim() : "全部";
                        string vDepNoManStr = (eDepNo_Man_Search.Text.Trim() != "") ? eDepNo_Man_Search.Text.Trim() : "全部";
                        string vAssignManStr = (eAssignMan_Search.Text.Trim() != "") ? eAssignMan_Search.Text.Trim() : "全部";
                        string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? eCarID_Search.Text.Trim() : "";
                        string vDriverStr = (eDriver_Search.Text.Trim() != "") ? eDriver_Search.Text.Trim() : "";
                        string vRecordNote = "匯出檔案_拖吊維修費用月統計表" + Environment.NewLine +
                                             "TowTruckAssignList.aspx" + Environment.NewLine +
                                             "拖吊日期：" + vCaseDateStr + Environment.NewLine +
                                             "車輛所屬站別：" + vDepNoCarStr + Environment.NewLine +
                                             "檢修人員單位：" + vDepNoManStr + Environment.NewLine +
                                             "檢修人員：" + vAssignManStr + Environment.NewLine +
                                             "牌照號碼：" + vCarIDStr + Environment.NewLine +
                                             "駕駛員：" + vDriverStr;
                        PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    }
                }
            }
            else
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('請先選擇查詢起迄日期！')");
                Response.Write("</" + "Script>");
                eCaseDate_S_Search.Focus();
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void fvTowTruckAssignDetail_DataBound(object sender, EventArgs e)
        {
            string vDateURL_Temp = "";
            string vDateScript_Temp = "";
            string vSQLStr_Temp = "";

            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            switch (fvTowTruckAssignDetail.CurrentMode)
            {
                case FormViewMode.ReadOnly:
                    //2020.10.15 新增 IsChecked
                    CheckBox cbIsChecked_List = (CheckBox)fvTowTruckAssignDetail.FindControl("cbIsChecked_List");
                    if (cbIsChecked_List != null)
                    {
                        Label eIsChecked_List = (Label)fvTowTruckAssignDetail.FindControl("eIsChecked_List");
                        cbIsChecked_List.Enabled = ((vLoginDepNo == "09") || ((vLoginDepNo == "30") && ((vLoginTitle == "130") || (vLoginTitle == "230") || (vLoginTitle == "270"))));
                        cbIsChecked_List.Checked = (eIsChecked_List.Text.Trim() == "V");
                    }
                    break;

                case FormViewMode.Edit:
                    Label eModifyDate_Edit = (Label)fvTowTruckAssignDetail.FindControl("eModifyDate_Edit");
                    Label eModifyMan_Edit = (Label)fvTowTruckAssignDetail.FindControl("eModifyMan_Edit");
                    Label eModifyManName_Edit = (Label)fvTowTruckAssignDetail.FindControl("eModifyManName_Edit");
                    if (eModifyDate_Edit != null)
                    {
                        eModifyDate_Edit.Text = PF.GetChinsesDate(DateTime.Today);
                        eModifyMan_Edit.Text = vLoginID;
                        eModifyManName_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    }
                    TextBox eCaseDate_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eCaseDate_Edit");
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Edit.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Edit.Attributes["onClick"] = vDateScript_Temp;

                    TextBox eLastServiceDate_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eLastServiceDate_Edit");
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eLastServiceDate_Edit.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eLastServiceDate_Edit.Attributes["onClick"] = vDateScript_Temp;
                    eLastServiceDate_Edit.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    DropDownList ddlCaseType_Edit = (DropDownList)fvTowTruckAssignDetail.FindControl("ddlCaseType_Edit");
                    Label eCaseType_Edit = (Label)fvTowTruckAssignDetail.FindControl("eCaseType_Edit");
                    ddlCaseType_Edit.Items.Clear();
                    using (SqlConnection connCaseType_Edit = new SqlConnection(vConnStr))
                    {
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                       " where FKey = '拖吊車叫用記錄表TowTruckAssign  CaseType' " + Environment.NewLine +
                                       " order by ClassNo ";
                        SqlCommand cmdCaseType_Edit = new SqlCommand(vSQLStr_Temp, connCaseType_Edit);
                        connCaseType_Edit.Open();
                        SqlDataReader drCaseType_Edit = cmdCaseType_Edit.ExecuteReader();
                        while (drCaseType_Edit.Read())
                        {
                            ddlCaseType_Edit.Items.Add(new ListItem(drCaseType_Edit["ClassTxt"].ToString().Trim(), drCaseType_Edit["ClassNo"].ToString().Trim()));
                        }
                    }
                    ddlCaseType_Edit.SelectedIndex = (eCaseType_Edit.Text.Trim() != "") ? ddlCaseType_Edit.Items.IndexOf(ddlCaseType_Edit.Items.FindByValue(eCaseType_Edit.Text.Trim())) : 0;

                    DropDownList ddlDamageAnalysis_Edit = (DropDownList)fvTowTruckAssignDetail.FindControl("ddlDamageAnalysis_Edit");
                    Label eDamageAnalysis_Edit = (Label)fvTowTruckAssignDetail.FindControl("eDamageAnalysis_Edit");
                    ddlDamageAnalysis_Edit.Items.Clear();
                    using (SqlConnection connDamageAnalysis_Edit = new SqlConnection(vConnStr))
                    {
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                       " where FKey = '拖吊車叫用記錄表TowTruckAssign  DamageAnalysis' " + Environment.NewLine +
                                       " order by ClassNo";
                        SqlCommand cmdDamageAnalysis_Edit = new SqlCommand(vSQLStr_Temp, connDamageAnalysis_Edit);
                        connDamageAnalysis_Edit.Open();
                        SqlDataReader drDamageAnalysis_Edit = cmdDamageAnalysis_Edit.ExecuteReader();
                        while (drDamageAnalysis_Edit.Read())
                        {
                            ddlDamageAnalysis_Edit.Items.Add(new ListItem(drDamageAnalysis_Edit["ClassTxt"].ToString().Trim(), drDamageAnalysis_Edit["ClassNo"].ToString().Trim()));
                        }
                    }
                    ddlDamageAnalysis_Edit.SelectedIndex = (eDamageAnalysis_Edit.Text.Trim() != "") ? ddlDamageAnalysis_Edit.Items.IndexOf(ddlDamageAnalysis_Edit.Items.FindByValue(eDamageAnalysis_Edit.Text.Trim())) : 0;
                    ddlDamageAnalysis_Edit.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    TextBox eTowingCost_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eTowingCost_Edit");
                    eTowingCost_Edit.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    TextBox eFixFee_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eFixFee_Edit");
                    eFixFee_Edit.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    TextBox eCarKMs_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eCarKMs_Edit");
                    eCarKMs_Edit.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    TextBox eWorkSheetNo_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eWorkSheetNo_Edit");
                    eWorkSheetNo_Edit.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    //2020.10.15 新增 IsChecked
                    CheckBox cbIsChecked_Edit = (CheckBox)fvTowTruckAssignDetail.FindControl("cbIsChecked_Edit");
                    if (cbIsChecked_Edit != null)
                    {
                        Label eIsChecked_Edit = (Label)fvTowTruckAssignDetail.FindControl("eIsChecked_Edit");
                        cbIsChecked_Edit.Enabled = ((vLoginDepNo == "09") || ((vLoginDepNo == "30") && ((vLoginTitle == "130") || (vLoginTitle == "230") || (vLoginTitle == "270"))));
                        cbIsChecked_Edit.Checked = (eIsChecked_Edit.Text.Trim() == "V");
                    }
                    break;

                case FormViewMode.Insert:
                    Label eBuDate_INS = (Label)fvTowTruckAssignDetail.FindControl("eBuDate_INS");
                    Label eBuMan_INS = (Label)fvTowTruckAssignDetail.FindControl("eBuMan_INS");
                    Label eBuManName_INS = (Label)fvTowTruckAssignDetail.FindControl("eBuManName_INS");
                    if (eBuDate_INS != null)
                    {
                        eBuDate_INS.Text = PF.GetChinsesDate(DateTime.Today);
                        eBuMan_INS.Text = vLoginID;
                        eBuManName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vLoginID + "' ", "Name");
                    }
                    TextBox eCaseDate_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eCaseDate_INS");
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_INS.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_INS.Attributes["onClick"] = vDateScript_Temp;

                    TextBox eLastServiceDate_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eLastServiceDate_INS");
                    vDateURL_Temp = "InputDate_ChineseYears.aspx?TextboxID=" + eLastServiceDate_INS.ClientID;
                    vDateScript_Temp = "window.open('" + vDateURL_Temp + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eLastServiceDate_INS.Attributes["onClick"] = vDateScript_Temp;
                    eLastServiceDate_INS.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    DropDownList ddlCaseType_INS = (DropDownList)fvTowTruckAssignDetail.FindControl("ddlCaseType_INS");
                    Label eCaseType_INS = (Label)fvTowTruckAssignDetail.FindControl("eCaseType_INS");
                    ddlCaseType_INS.Items.Clear();
                    using (SqlConnection connCaseType_INS = new SqlConnection(vConnStr))
                    {
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                       " where FKey = '拖吊車叫用記錄表TowTruckAssign  CaseType' " + Environment.NewLine +
                                       " order by ClassNo ";
                        SqlCommand cmdCaseType_INS = new SqlCommand(vSQLStr_Temp, connCaseType_INS);
                        connCaseType_INS.Open();
                        SqlDataReader drCaseType_INS = cmdCaseType_INS.ExecuteReader();
                        while (drCaseType_INS.Read())
                        {
                            ddlCaseType_INS.Items.Add(new ListItem(drCaseType_INS["ClassTxt"].ToString().Trim(), drCaseType_INS["ClassNo"].ToString().Trim()));
                        }
                    }
                    ddlCaseType_INS.SelectedIndex = 0;
                    eCaseType_INS.Text = "";

                    DropDownList ddlDamageAnalysis_INS = (DropDownList)fvTowTruckAssignDetail.FindControl("ddlDamageAnalysis_INS");
                    Label eDamageAnalysis_INS = (Label)fvTowTruckAssignDetail.FindControl("eDamageAnalysis_INS");
                    ddlDamageAnalysis_INS.Items.Clear();
                    using (SqlConnection connDamageAnalysis_INS = new SqlConnection(vConnStr))
                    {
                        vSQLStr_Temp = "select cast('' as varchar) ClassNo, cast('' as varchar) ClassTxt " + Environment.NewLine +
                                       " union all " + Environment.NewLine +
                                       "select ClassNo, ClassTxt from DBDICB " + Environment.NewLine +
                                       " where FKey = '拖吊車叫用記錄表TowTruckAssign  DamageAnalysis' " + Environment.NewLine +
                                       " order by ClassNo";
                        SqlCommand cmdDamageAnalysis_INS = new SqlCommand(vSQLStr_Temp, connDamageAnalysis_INS);
                        connDamageAnalysis_INS.Open();
                        SqlDataReader drDamageAnalysis_INS = cmdDamageAnalysis_INS.ExecuteReader();
                        while (drDamageAnalysis_INS.Read())
                        {
                            ddlDamageAnalysis_INS.Items.Add(new ListItem(drDamageAnalysis_INS["ClassTxt"].ToString().Trim(), drDamageAnalysis_INS["ClassNo"].ToString().Trim()));
                        }
                    }
                    ddlDamageAnalysis_INS.SelectedIndex = 0;
                    eDamageAnalysis_INS.Text = "";
                    ddlDamageAnalysis_INS.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    TextBox eTowingCost_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eTowingCost_INS");
                    eTowingCost_INS.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    TextBox eFixFee_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eFixFee_INS");
                    eFixFee_INS.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    TextBox eCarKMs_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eCarKMs_INS");
                    eCarKMs_INS.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    TextBox eWorkSheetNo_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eWorkSheetNo_INS");
                    eWorkSheetNo_INS.Enabled = ((vLoginDepNo == "09") || (vLoginDepNo == "30"));

                    //2020.10.15 新增 IsChecked
                    CheckBox cbIsChecked_INS = (CheckBox)fvTowTruckAssignDetail.FindControl("cbIsChecked_INS");
                    if (cbIsChecked_INS != null)
                    {
                        Label eIsChecked_INS = (Label)fvTowTruckAssignDetail.FindControl("eIsChecked_INS");
                        cbIsChecked_INS.Checked = false;
                        cbIsChecked_INS.Enabled = ((vLoginDepNo == "09") || ((vLoginDepNo == "30") && ((vLoginTitle == "130") || (vLoginTitle == "230") || (vLoginTitle == "270"))));
                        eIsChecked_INS.Text = "";
                    }
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 修改確認
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Edit_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            Label eCaseNo_Edit = (Label)fvTowTruckAssignDetail.FindControl("eCaseNo_Edit");
            string vCaseNo_Temp = (eCaseNo_Edit != null) ? eCaseNo_Edit.Text.Trim() : "";
            if (vCaseNo_Temp != "")
            {
                string vSQLStr_Temp = "select Max(Items) MaxItems from TowTruckAssignList_History where CaseNo = '" + vCaseNo_Temp + "' ";
                string vMaxItemsStr = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItems");
                int vNewIndex = (vMaxItemsStr != "") ? Int32.Parse(vMaxItemsStr.Trim()) + 1 : 1;
                string vCaseNoItem = vCaseNo_Temp.Trim() + vNewIndex.ToString("D4");
                //寫入異動記錄檔
                vSQLStr_Temp = "insert into TowTruckAssignList_History " + Environment.NewLine +
                               "       (CaseNoItem, Items, ModifyMode, CaseNo, CaseDate, Car_ID, Driver, DepNo_Car, CasePosition, " + Environment.NewLine +
                               "        ParkingPosition, CaseTime, AssignMan, DepNo_Man, FirstSupportMan, SecondSupportMan, " + Environment.NewLine +
                               "        Determination, FaultParts, FaultReason, Dispose, FollowUp, Improvements, Remark, " + Environment.NewLine +
                               "        BuMan, BuDate, ModifyMan, ModifyDate, CarKMs, TowingCost, FixFee, " + Environment.NewLine +
                               //"        LastServiceDate, DamageAnalysis, CaseType, WorkSheetNo) " + Environment.NewLine +
                               "        LastServiceDate, DamageAnalysis, CaseType, WorkSheetNo, IsChecked) " + Environment.NewLine +
                               "select '" + vCaseNoItem.Trim() + "', '" + vNewIndex.ToString("D4") + "', 'Edit', " + Environment.NewLine +
                               "       CaseNo, CaseDate, Car_ID, Driver, DepNo_Car, CasePosition, " + Environment.NewLine +
                               "       ParkingPosition, CaseTime, AssignMan, DepNo_Man, FirstSupportMan, SecondSupportMan, " + Environment.NewLine +
                               "       Determination, FaultParts, FaultReason, Dispose, FollowUp, Improvements, Remark, " + Environment.NewLine +
                               "       BuMan, BuDate, ModifyMan, ModifyDate, CarKMs, TowingCost, FixFee, " + Environment.NewLine +
                               //"       LastServiceDate, DamageAnalysis, CaseType, WorkSheetNo " + Environment.NewLine +
                               "       LastServiceDate, DamageAnalysis, CaseType, WorkSheetNo, IsChecked " + Environment.NewLine +
                               "  from TowTruckAssignList " + Environment.NewLine +
                               " where CaseNo = '" + vCaseNo_Temp.Trim() + "' ";
                //
                try
                {
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);

                    //清除修改命令的參數
                    TextBox eCar_ID_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eCar_ID_Edit");
                    TextBox eDepNo_Car_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eDepNo_Car_Edit");
                    TextBox eDriver_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eDriver_Edit");
                    TextBox eCaseDate_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eCaseDate_Edit");
                    TextBox eCaseTime_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eCaseTime_Edit");
                    TextBox eCasePosition_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eCasePosition_Edit");
                    TextBox eParkingPosition_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eParkingPosition_Edit");
                    TextBox eDetermination_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eDetermination_Edit");
                    TextBox eAssignMan_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eAssignMan_Edit");
                    Label eDepNo_Man_Edit = (Label)fvTowTruckAssignDetail.FindControl("eDepNo_Man_Edit");
                    TextBox eFirstSupportMan_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eFirstSupportMan_Edit");
                    TextBox eSecondSupportMan_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eSecondSupportMan_Edit");
                    TextBox eFaultParts_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eFaultParts_Edit");
                    TextBox eFaultReason_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eFaultReason_Edit");
                    TextBox eDispose_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eDispose_Edit");
                    TextBox eFollowUp_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eFollowUp_Edit");
                    TextBox eImprovements_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eImprovements_Edit");
                    TextBox eRemark_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eRemark_Edit");
                    Label eBuMan_Edit = (Label)fvTowTruckAssignDetail.FindControl("eBuMan_Edit");
                    Label eBuDate_Edit = (Label)fvTowTruckAssignDetail.FindControl("eBuDate_Edit");
                    Label eModifyMan_Edit = (Label)fvTowTruckAssignDetail.FindControl("eModifyMan_Edit");
                    Label eModifyDate_Edit = (Label)fvTowTruckAssignDetail.FindControl("eModifyDate_Edit");
                    TextBox eCarKMs_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eCarKMs_Edit");
                    TextBox eTowingCost_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eTowingCost_Edit");
                    TextBox eFixFee_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eFixFee_Edit");
                    TextBox eLastServiceDate_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eLastServiceDate_Edit");
                    Label eCaseType_Edit = (Label)fvTowTruckAssignDetail.FindControl("eCaseType_Edit");
                    TextBox eWorkSheetNo_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eWorkSheetNo_Edit");
                    Label eDamageAnalysis_Edit = (Label)fvTowTruckAssignDetail.FindControl("eDamageAnalysis_Edit");
                    //2020.10.15 新增 IsChecked
                    Label eIsChecked_Edit = (Label)fvTowTruckAssignDetail.FindControl("eIsChecked_Edit");

                    string vSQLStr_Edit = "UPDATE TowTruckAssignList " + Environment.NewLine +
                                          "   SET CaseDate = @CaseDate, Car_ID = @Car_ID, Driver = @Driver, DepNo_Car = @DepNo_Car, " + Environment.NewLine +
                                          "       CasePosition = @CasePosition, ParkingPosition = @ParkingPosition, CaseTime = @CaseTime, " + Environment.NewLine +
                                          "       AssignMan = @AssignMan, DepNo_Man = @DepNo_Man, FirstSupportMan = @FirstSupportMan, " + Environment.NewLine +
                                          "       SecondSupportMan = @SecondSupportMan, Determination = @Determination, FaultParts = @FaultParts, " + Environment.NewLine +
                                          "       FaultReason = @FaultReason, Dispose = @Dispose, FollowUp = @FollowUp, Improvements = @Improvements, " + Environment.NewLine +
                                          "       Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate, CarKMs = @CarKMs, TowingCost = @TowingCost, " + Environment.NewLine +
                                          "       FixFee = @FixFee, LastServiceDate = @LastServiceDate, DamageAnalysis = @DamageAnalysis, CaseType = @CaseType, " + Environment.NewLine +
                                          "       WorkSheetNo = @WorkSheetNo, IsChecked = @IsChecked " + Environment.NewLine +
                                          " WHERE (CaseNo = @CaseNo)";
                    sdsTowTruckAssignDetail.UpdateCommand = vSQLStr_Edit;

                    sdsTowTruckAssignDetail.UpdateParameters.Clear();
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("Car_ID", DbType.String, (eCar_ID_Edit.Text.Trim() != "") ? eCar_ID_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("DepNo_Car", DbType.String, (eDepNo_Car_Edit.Text.Trim() != "") ? eDepNo_Car_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("Driver", DbType.String, (eDriver_Edit.Text.Trim() != "") ? eDriver_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("CaseDate", DbType.DateTime, (eCaseDate_Edit.Text.Trim() != "") ? eCaseDate_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("CaseTime", DbType.String, (eCaseTime_Edit.Text.Trim() != "") ? eCaseTime_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("CasePosition", DbType.String, (eCasePosition_Edit.Text.Trim() != "") ? eCasePosition_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("ParkingPosition", DbType.String, (eParkingPosition_Edit.Text.Trim() != "") ? eParkingPosition_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("Determination", DbType.String, (eDetermination_Edit.Text.Trim() != "") ? eDetermination_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan_Edit.Text.Trim() != "") ? eAssignMan_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("DepNo_Man", DbType.String, (eDepNo_Man_Edit.Text.Trim() != "") ? eDepNo_Man_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("FirstSupportMan", DbType.String, (eFirstSupportMan_Edit.Text.Trim() != "") ? eFirstSupportMan_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("SecondSupportMan", DbType.String, (eSecondSupportMan_Edit.Text.Trim() != "") ? eSecondSupportMan_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("FaultParts", DbType.String, (eFaultParts_Edit.Text.Trim() != "") ? eFaultParts_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("FaultReason", DbType.String, (eFaultReason_Edit.Text.Trim() != "") ? eFaultReason_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("Dispose", DbType.String, (eDispose_Edit.Text.Trim() != "") ? eDispose_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("FollowUp", DbType.String, (eFollowUp_Edit.Text.Trim() != "") ? eFollowUp_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("Improvements", DbType.String, (eImprovements_Edit.Text.Trim() != "") ? eImprovements_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("Remark", DbType.String, (eRemark_Edit.Text.Trim() != "") ? eRemark_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("BuMan", DbType.String, (eBuMan_Edit.Text.Trim() != "") ? eBuMan_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("BuDate", DbType.DateTime, (eBuDate_Edit.Text.Trim() != "") ? eBuDate_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("ModifyMan", DbType.String, (eModifyMan_Edit.Text.Trim() != "") ? eModifyMan_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("ModifyDate", DbType.DateTime, (eModifyDate_Edit.Text.Trim() != "") ? eModifyDate_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("CarKMs", DbType.Double, (eCarKMs_Edit.Text.Trim() != "") ? eCarKMs_Edit.Text.Trim() : "0"));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("TowingCost", DbType.Double, (eTowingCost_Edit.Text.Trim() != "") ? eTowingCost_Edit.Text.Trim() : "0"));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("FixFee", DbType.Double, (eFixFee_Edit.Text.Trim() != "") ? eFixFee_Edit.Text.Trim() : "0"));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("LastServiceDate", DbType.DateTime, (eLastServiceDate_Edit.Text.Trim() != "") ? eLastServiceDate_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("CaseType", DbType.String, (eCaseType_Edit.Text.Trim() != "") ? eCaseType_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("WorkSheetNo", DbType.String, (eWorkSheetNo_Edit.Text.Trim() != "") ? eWorkSheetNo_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("DamageAnalysis", DbType.String, (eDamageAnalysis_Edit.Text.Trim() != "") ? eDamageAnalysis_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("IsChecked", DbType.String, (eIsChecked_Edit.Text.Trim() != "") ? eIsChecked_Edit.Text.Trim() : String.Empty));
                    sdsTowTruckAssignDetail.UpdateParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp));

                    sdsTowTruckAssignDetail.Update();
                    fvTowTruckAssignDetail.ChangeMode(FormViewMode.ReadOnly);
                    fvTowTruckAssignDetail.DataBind();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void eCar_ID_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eCarID_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eCar_ID_Edit");
            if ((eCarID_Edit != null) && (eCarID_Edit.Text.Trim() != ""))
            {
                string vSQLStr_Temp = "select (isnull(CompanyNo, '') + ',' + " + Environment.NewLine +
                                      "        isnull(Car_Class, '') + ',' + " + Environment.NewLine +
                                      "        cast(Year(GetLicDate) - 1911 as varchar) + '/' + cast(Month(GetLicDate) as varchar) + '/' + cast(day(GetLicDate) as varchar) + ',' + " + Environment.NewLine +
                                      "        isnull(Point, '') + ',' + " + Environment.NewLine +
                                      "        isnull(Driver, '') + ',' + " + Environment.NewLine +
                                      "        cast(datediff(Month, GetLicDate, GetDate()) / 12 as varchar) + ',' + " + Environment.NewLine +
                                      "        cast(datediff(month, GetLicDate, GetDate()) % 12 as varchar) " + Environment.NewLine +
                                      "       ) CarData " + Environment.NewLine +
                                      "  from Car_InfoA where Car_ID = '" + eCarID_Edit.Text.Trim() + "' ";
                string vResultStr = PF.GetValue(vConnStr, vSQLStr_Temp, "CarData");
                string[] aCarData = vResultStr.Split(',');
                string vDepNo_Car = aCarData[0];
                string vCar_Class = aCarData[1];
                DateTime vGetLicDate = (aCarData[2] != "") ? DateTime.Parse((DateTime.Parse(aCarData[2]).Year - 1911).ToString() + "/" + DateTime.Parse(aCarData[2]).ToString("MM/dd")) : DateTime.Today;
                string vGetLicDateStr = PF.GetChinsesDate(vGetLicDate);
                string vPoint = aCarData[3];
                string vDriver = aCarData[4];
                string vCarAge_Y = aCarData[5];
                string vCarAge_M = aCarData[6];
                TextBox eDepNo_Car_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eDepNo_Car_Edit");
                eDepNo_Car_Edit.Text = vDepNo_Car.Trim();
                Label eDepName_Car_Edit = (Label)fvTowTruckAssignDetail.FindControl("eDepName_Car_Edit");
                eDepName_Car_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Department where DepNo = '" + vDepNo_Car + "' ", "Name");
                Label eGetLicDate_Edit = (Label)fvTowTruckAssignDetail.FindControl("eGetLicDate_Edit");
                eGetLicDate_Edit.Text = vGetLicDateStr.Trim();
                Label eCarAge_Year_Edit = (Label)fvTowTruckAssignDetail.FindControl("eCarAge_Year_Edit");
                eCarAge_Year_Edit.Text = vCarAge_Y.Trim();
                Label eCarAge_Month_Edit = (Label)fvTowTruckAssignDetail.FindControl("eCarAge_Month_Edit");
                eCarAge_Month_Edit.Text = vCarAge_M.Trim();
                Label eCar_Class_Edit = (Label)fvTowTruckAssignDetail.FindControl("eCar_Class_Edit");
                eCar_Class_Edit.Text = vCar_Class.Trim();
                Label eCar_ClassName_Edit = (Label)fvTowTruckAssignDetail.FindControl("eCar_ClassName_Edit");
                eCar_ClassName_Edit.Text = PF.GetValue(vConnStr, "select ClassTxt from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and ClassNo = '" + vCar_Class.Trim() + "' ", "ClassTXT");
                Label ePoint_Edit = (Label)fvTowTruckAssignDetail.FindControl("ePoint_Edit");
                ePoint_Edit.Text = vPoint.Trim();
                Label ePointName_Edit = (Label)fvTowTruckAssignDetail.FindControl("ePointName_Edit");
                ePointName_Edit.Text = PF.GetValue(vConnStr, "select ClassTXT from DBDICB where FKey = '車輛資料作業    Car_infoA       POINT' and ClassNo = '" + vPoint.Trim() + "' ", "ClassTXT");
                TextBox eDriver_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eDriver_Edit");
                Label eDriverName_Edit = (Label)fvTowTruckAssignDetail.FindControl("eDriverName_Edit");
                eDriver_Edit.Text = vDriver.Trim();
                eDriverName_Edit.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vDriver + "' ", "Name");
                //2020.06.23 里程數不帶...USER 自行輸入
                //TextBox eCarKMs_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eCarKMs_Edit");
                //vSQLStr_Temp = "select round(sum(ActualKM), 2) TotalKMs from RunSheetB where Car_ID = '" + eCarID_Edit.Text.Trim() + "' and isnull(ReduceReason, '') = '' Group by Car_ID";
                //eCarKMs_Edit.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "TotalKMs");
                TextBox eLastServiceDate_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eLastServiceDate_Edit");
                vSQLStr_Temp = "select top 1 cast((year(FixDateIn) - 1911) as nvarchar) + '/' + cast(Month(FixDateIn) as nvarchar) + '/' + Cast(day(FixDateIn) as nvarchar) as FixDateIn_C from FixWorkA where Car_ID = '" + eCarID_Edit.Text.Trim() + "' and [Service] = '1' order by FixDateIn DESC";
                eLastServiceDate_Edit.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "FixDateIn_C");
            }
        }

        protected void ddlCaseType_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlCaseType_Edit = (DropDownList)fvTowTruckAssignDetail.FindControl("ddlCaseType_Edit");
            Label eCaseType_Edit = (Label)fvTowTruckAssignDetail.FindControl("eCaseType_Edit");
            eCaseType_Edit.Text = ddlCaseType_Edit.SelectedValue.Trim();
        }

        protected void eDepNo_Car_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Car_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eDepNo_Car_Edit");
            Label eDepName_Car_Edit = (Label)fvTowTruckAssignDetail.FindControl("eDepName_Car_Edit");
            string vDepNo_Temp = eDepNo_Car_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Car_Edit.Text = vDepNo_Temp.Trim();
            eDepName_Car_Edit.Text = vDepName_Temp.Trim();
        }

        protected void eDriver_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eDriver_Edit");
            Label eDriverName_Edit = (Label)fvTowTruckAssignDetail.FindControl("eDriverName_Edit");
            string vDriver_Temp = eDriver_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp == "")
            {
                vDriverName_Temp = vDriver_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_Edit.Text = vDriver_Temp.Trim();
            eDriverName_Edit.Text = vDriverName_Temp.Trim();
        }

        protected void eAssignMan_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eAssignMan_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eAssignMan_Edit");
            Label eAssignManName_Edit = (Label)fvTowTruckAssignDetail.FindControl("eAssignManName_Edit");
            string vAssignMan_Temp = eAssignMan_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan_Temp + "' ";
            string vAssignManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vAssignManName_Temp == "")
            {
                vAssignManName_Temp = vAssignMan_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vAssignManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vAssignMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eAssignMan_Edit.Text = vAssignMan_Temp.Trim();
            eAssignManName_Edit.Text = vAssignManName_Temp.Trim();
            if (vAssignMan_Temp.Trim() != "")
            {
                Label eDepNo_Man_Edit = (Label)fvTowTruckAssignDetail.FindControl("eDepNo_Man_Edit");
                eDepNo_Man_Edit.Text = PF.GetValue(vConnStr, "select DepNo from Employee where EmpNo = '" + vAssignMan_Temp.Trim() + "' ", "DepNo");
            }
        }

        protected void eFirstSupportMan_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eFirstSupportMan_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eFirstSupportMan_Edit");
            Label eFirstSupportManName_Edit = (Label)fvTowTruckAssignDetail.FindControl("eFirstSupportManName_Edit");
            string vFirstSupportMan_Temp = eFirstSupportMan_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vFirstSupportMan_Temp + "' ";
            string vFirstSupportManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vFirstSupportManName_Temp == "")
            {
                vFirstSupportManName_Temp = vFirstSupportMan_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vFirstSupportManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vFirstSupportMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eFirstSupportMan_Edit.Text = vFirstSupportMan_Temp.Trim();
            eFirstSupportManName_Edit.Text = vFirstSupportManName_Temp.Trim();
        }

        protected void eSecondSupportMan_Edit_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eSecondSupportMan_Edit = (TextBox)fvTowTruckAssignDetail.FindControl("eSecondSupportMan_Edit");
            Label eSecondSupportManName_Edit = (Label)fvTowTruckAssignDetail.FindControl("eSecondSupportManName_Edit");
            string vSecondSupportMan_Temp = eSecondSupportMan_Edit.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vSecondSupportMan_Temp + "' ";
            string vSecondSupportManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vSecondSupportManName_Temp == "")
            {
                vSecondSupportManName_Temp = vSecondSupportMan_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vSecondSupportManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vSecondSupportMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eSecondSupportMan_Edit.Text = vSecondSupportMan_Temp.Trim();
            eSecondSupportManName_Edit.Text = vSecondSupportManName_Temp.Trim();
        }

        protected void ddlDamageAnalysis_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlDamageAnalysis_Edit = (DropDownList)fvTowTruckAssignDetail.FindControl("ddlDamageAnalysis_Edit");
            Label eDamageAnalysis_Edit = (Label)fvTowTruckAssignDetail.FindControl("eDamageAnalysis_Edit");
            eDamageAnalysis_Edit.Text = ddlDamageAnalysis_Edit.SelectedValue.Trim();
        }

        /// <summary>
        /// 新增確定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_INS_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            try
            {
                TextBox eCaseDate_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eCaseDate_INS");
                DateTime vCaseDate_Temp = (eCaseDate_INS.Text.Trim() != "") ? DateTime.Parse(eCaseDate_INS.Text.Trim()) : DateTime.Today;
                string vCaseNo_Head = vCaseDate_Temp.Year.ToString() + vCaseDate_Temp.Month.ToString("D2");
                string vSQLStr_Temp = "select Max(CaseNo) MaxCaseNo from TowTruckAssignList where CaseNo like '" + vCaseNo_Head + "%' ";
                string vOldIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxCaseNo").Replace(vCaseNo_Head, "");
                int vNewIndex = (vOldIndex != "") ? Int32.Parse(vOldIndex) + 1 : 1;
                string vCaseNo_New = vCaseNo_Head + vNewIndex.ToString("D4");
                TextBox eCar_ID_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eCar_ID_INS");
                TextBox eDriver_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eDriver_INS");
                TextBox eDepNo_Car_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eDepNo_Car_INS");
                TextBox eCasePosition_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eCasePosition_INS");
                TextBox eParkingPosition_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eParkingPosition_INS");
                TextBox eCaseTime_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eCaseTime_INS");
                TextBox eAssignMan_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eAssignMan_INS");
                Label eDepNo_Man_INS = (Label)fvTowTruckAssignDetail.FindControl("eDepNo_Man_INS");
                TextBox eFirstSupportMan_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eFirstSupportMan_INS");
                TextBox eSecondSupportMan_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eSecondSupportMan_INS");
                TextBox eDetermination_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eDetermination_INS");
                TextBox eFaultParts_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eFaultParts_INS");
                TextBox eFaultReason_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eFaultReason_INS");
                TextBox eDispose_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eDispose_INS");
                TextBox eFollowUp_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eFollowUp_INS");
                TextBox eImprovements_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eImprovements_INS");
                TextBox eRemark_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eRemark_INS");
                TextBox eCarKMs_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eCarKMs_INS");
                TextBox eTowingCost_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eTowingCost_INS");
                TextBox eFixFee_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eFixFee_INS");
                TextBox eLastServiceDate_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eLastServiceDate_INS");
                Label eDamageAnalysis_INS = (Label)fvTowTruckAssignDetail.FindControl("eDamageAnalysis_INS");
                Label eCaseType_INS = (Label)fvTowTruckAssignDetail.FindControl("eCaseType_INS");
                TextBox eWorkSheetNo_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eWorkSheetNo_INS");
                //2020.10.15 新增 IsChecked
                Label eIsChecked_INS = (Label)fvTowTruckAssignDetail.FindControl("eIsChecked_INS");

                string vSQLStr_INS = "INSERT INTO TowTruckAssignList " + Environment.NewLine +
                                     "       (CaseNo, CaseDate, Car_ID, Driver, DepNo_Car, CasePosition, ParkingPosition, CaseTime, AssignMan, DepNo_Man, " + Environment.NewLine +
                                     "        FirstSupportMan, SecondSupportMan, Determination, FaultParts, FaultReason, Dispose, FollowUp, Improvements, " + Environment.NewLine +
                                     "        Remark, BuMan, BuDate, CarKMs, TowingCost, FixFee, LastServiceDate, DamageAnalysis, CaseType, WorkSheetNo, IsChecked) " + Environment.NewLine +
                                     "VALUES (@CaseNo, @CaseDate, @Car_ID, @Driver, @DepNo_Car, @CasePosition, @ParkingPosition, @CaseTime, @AssignMan, @DepNo_Man, " + Environment.NewLine +
                                     "        @FirstSupportMan, @SecondSupportMan, @Determination, @FaultParts, @FaultReason, @Dispose, @FollowUp, @Improvements, " + Environment.NewLine +
                                     "        @Remark, @BuMan, @BuDate, @CarKMs, @TowingCost, @FixFee, @LastServiceDate, @DamageAnalysis, @CaseType, @WorkSheetNo, @IsChecked) ";
                sdsTowTruckAssignDetail.InsertCommand = vSQLStr_INS;

                sdsTowTruckAssignDetail.InsertParameters.Clear();
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_New.Trim()));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("CaseDate", DbType.DateTime, vCaseDate_Temp.ToShortDateString()));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("Car_ID", DbType.String, (eCar_ID_INS.Text.Trim() != "") ? eCar_ID_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("Driver", DbType.String, (eDriver_INS.Text.Trim() != "") ? eDriver_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("DepNo_Car", DbType.String, (eDepNo_Car_INS.Text.Trim() != "") ? eDepNo_Car_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("CasePosition", DbType.String, (eCasePosition_INS.Text.Trim() != "") ? eCasePosition_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("ParkingPosition", DbType.String, (eParkingPosition_INS.Text.Trim() != "") ? eParkingPosition_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("CaseTime", DbType.String, (eCaseTime_INS.Text.Trim() != "") ? eCaseTime_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("AssignMan", DbType.String, (eAssignMan_INS.Text.Trim() != "") ? eAssignMan_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("DepNo_Man", DbType.String, (eDepNo_Man_INS.Text.Trim() != "") ? eDepNo_Man_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("FirstSupportMan", DbType.String, (eFirstSupportMan_INS.Text.Trim() != "") ? eFirstSupportMan_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("SecondSupportMan", DbType.String, (eSecondSupportMan_INS.Text.Trim() != "") ? eSecondSupportMan_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("Determination", DbType.String, (eDetermination_INS.Text.Trim() != "") ? eDetermination_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("FaultParts", DbType.String, (eFaultParts_INS.Text.Trim() != "") ? eFaultParts_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("FaultReason", DbType.String, (eFaultReason_INS.Text.Trim() != "") ? eFaultReason_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("Dispose", DbType.String, (eDispose_INS.Text.Trim() != "") ? eDispose_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("FollowUp", DbType.String, (eFollowUp_INS.Text.Trim() != "") ? eFollowUp_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("Improvements", DbType.String, (eImprovements_INS.Text.Trim() != "") ? eImprovements_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("Remark", DbType.String, (eRemark_INS.Text.Trim() != "") ? eRemark_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("BuMan", DbType.String, vLoginID));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("BuDate", DbType.DateTime, DateTime.Today.ToShortDateString()));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("CarKMs", DbType.Double, (eCarKMs_INS.Text.Trim() != "") ? eCarKMs_INS.Text.Trim() : "0"));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("TowingCost", DbType.Double, (eTowingCost_INS.Text.Trim() != "") ? eTowingCost_INS.Text.Trim() : "0"));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("FixFee", DbType.Double, (eFixFee_INS.Text.Trim() != "") ? eFixFee_INS.Text.Trim() : "0"));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("LastServiceDate", DbType.DateTime, (eLastServiceDate_INS.Text.Trim() != "") ? eLastServiceDate_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("DamageAnalysis", DbType.String, (eDamageAnalysis_INS.Text.Trim() != "") ? eDamageAnalysis_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("CaseType", DbType.String, (eCaseType_INS.Text.Trim() != "") ? eCaseType_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("WorkSheetNo", DbType.String, (eWorkSheetNo_INS.Text.Trim() != "") ? eWorkSheetNo_INS.Text.Trim() : String.Empty));
                sdsTowTruckAssignDetail.InsertParameters.Add(new Parameter("IsChecked", DbType.String, (eIsChecked_INS.Text.Trim() != "") ? eIsChecked_INS.Text.Trim() : String.Empty));

                sdsTowTruckAssignDetail.Insert();
                fvTowTruckAssignDetail.ChangeMode(FormViewMode.ReadOnly);
                fvTowTruckAssignDetail.DataBind();
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }

        protected void ddlCaseType_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlCaseType_INS = (DropDownList)fvTowTruckAssignDetail.FindControl("ddlCaseType_INS");
            Label eCaseType_INS = (Label)fvTowTruckAssignDetail.FindControl("eCaseType_INS");
            eCaseType_INS.Text = ddlCaseType_INS.SelectedValue.Trim();
        }

        protected void eCar_ID_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eCarID_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eCar_ID_INS");
            if ((eCarID_INS != null) && (eCarID_INS.Text.Trim() != ""))
            {
                string vSQLStr_Temp = "select (isnull(CompanyNo, '') + ',' + " + Environment.NewLine +
                                      "        isnull(Car_Class, '') + ',' + " + Environment.NewLine +
                                      "        cast(Year(GetLicDate) - 1911 as varchar) + '/' + cast(Month(GetLicDate) as varchar) + '/' + cast(day(GetLicDate) as varchar) + ',' + " + Environment.NewLine +
                                      "        isnull(Point, '') + ',' + " + Environment.NewLine +
                                      "        isnull(Driver, '') + ',' + " + Environment.NewLine +
                                      "        cast(datediff(Month, GetLicDate, GetDate()) / 12 as varchar) + ',' + " + Environment.NewLine +
                                      "        cast(datediff(month, GetLicDate, GetDate()) % 12 as varchar) " + Environment.NewLine +
                                      "       ) CarData " + Environment.NewLine +
                                      "  from Car_InfoA where Car_ID = '" + eCarID_INS.Text.Trim() + "' ";
                string vResultStr = PF.GetValue(vConnStr, vSQLStr_Temp, "CarData");
                string[] aCarData = vResultStr.Split(',');
                string vDepNo_Car = aCarData[0];
                string vCar_Class = aCarData[1];
                DateTime vGetLicDate = (aCarData[2] != "") ? DateTime.Parse((DateTime.Parse(aCarData[2]).Year - 1911).ToString() + "/" + DateTime.Parse(aCarData[2]).ToString("MM/dd")) : DateTime.Today;
                string vGetLicDateStr = PF.GetChinsesDate(vGetLicDate);
                string vPoint = aCarData[3];
                string vDriver = aCarData[4];
                string vCarAge_Y = aCarData[5];
                string vCarAge_M = aCarData[6];
                TextBox eDepNo_Car_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eDepNo_Car_INS");
                eDepNo_Car_INS.Text = vDepNo_Car.Trim();
                Label eDepName_Car_INS = (Label)fvTowTruckAssignDetail.FindControl("eDepName_Car_INS");
                eDepName_Car_INS.Text = PF.GetValue(vConnStr, "select [Name] from Department where DepNo = '" + vDepNo_Car + "' ", "Name");
                Label eGetLicDate_INS = (Label)fvTowTruckAssignDetail.FindControl("eGetLicDate_INS");
                eGetLicDate_INS.Text = vGetLicDateStr.Trim();
                Label eCarAge_Year_INS = (Label)fvTowTruckAssignDetail.FindControl("eCarAge_Year_INS");
                eCarAge_Year_INS.Text = vCarAge_Y.Trim();
                Label eCarAge_Month_INS = (Label)fvTowTruckAssignDetail.FindControl("eCarAge_Month_INS");
                eCarAge_Month_INS.Text = vCarAge_M.Trim();
                Label eCar_Class_INS = (Label)fvTowTruckAssignDetail.FindControl("eCar_Class_INS");
                eCar_Class_INS.Text = vCar_Class.Trim();
                Label eCar_ClassName_INS = (Label)fvTowTruckAssignDetail.FindControl("eCar_ClassName_INS");
                eCar_ClassName_INS.Text = PF.GetValue(vConnStr, "select ClassTXT from DBDICB where FKey = '車輛資料作業    Car_infoA       CAR_CLASS' and CLassNo = '" + vCar_Class.Trim() + "' ", "ClassTXT");
                Label ePoint_INS = (Label)fvTowTruckAssignDetail.FindControl("ePoint_INS");
                ePoint_INS.Text = vPoint.Trim();
                Label ePointName_INS = (Label)fvTowTruckAssignDetail.FindControl("ePointName_INS");
                ePointName_INS.Text = PF.GetValue(vConnStr, "select ClassTXT from DBDICB where FKey = '車輛資料作業    Car_infoA       POINT' and ClassNo = '" + vPoint.Trim() + "' ", "ClassTXT");
                TextBox eDriver_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eDriver_INS");
                Label eDriverName_INS = (Label)fvTowTruckAssignDetail.FindControl("eDriverName_INS");
                eDriver_INS.Text = vDriver.Trim();
                eDriverName_INS.Text = PF.GetValue(vConnStr, "select [Name] from Employee where EmpNo = '" + vDriver + "' ", "Name");
                //2020.06.23 里程數不帶...USER 自行輸入
                //TextBox eCarKMs_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eCarKMs_INS");
                //vSQLStr_Temp = "select round(sum(ActualKM), 2) TotalKMs from RunSheetB where Car_ID = '" + eCarID_INS.Text.Trim() + "' and isnull(ReduceReason, '') = '' Group by Car_ID";
                //eCarKMs_INS.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "TotalKMs");
                TextBox eLastServiceDate_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eLastServiceDate_INS");
                vSQLStr_Temp = "select top 1 cast((year(FixDateIn) - 1911) as nvarchar) + '/' + cast(Month(FixDateIn) as nvarchar) + '/' + Cast(day(FixDateIn) as nvarchar) as FixDateIn_C from FixWorkA where Car_ID = '" + eCarID_INS.Text.Trim() + "' and [Service] = '1' order by FixDateIn DESC";
                eLastServiceDate_INS.Text = PF.GetValue(vConnStr, vSQLStr_Temp, "FixDateIn_C");
            }
        }

        protected void eDepNo_Car_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDepNo_Car_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eDepNo_Car_INS");
            Label eDepName_Car_INS = (Label)fvTowTruckAssignDetail.FindControl("eDepName_Car_INS");
            string vDepNo_Temp = eDepNo_Car_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
            string vDepName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDepName_Temp == "")
            {
                vDepName_Temp = vDepNo_Temp.Trim();
                vSQLStr_Temp = "select DepNo from Department where [Name] = '" + vDepName_Temp.Trim() + "' ";
                vDepNo_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "DepNo");
            }
            eDepNo_Car_INS.Text = vDepNo_Temp.Trim();
            eDepName_Car_INS.Text = vDepName_Temp.Trim();
        }

        protected void eDriver_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eDriver_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eDriver_INS");
            Label eDriverName_INS = (Label)fvTowTruckAssignDetail.FindControl("eDriverName_INS");
            string vDriver_Temp = eDriver_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vDriver_Temp + "' ";
            string vDriverName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vDriverName_Temp == "")
            {
                vDriverName_Temp = vDriver_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vDriver_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eDriver_INS.Text = vDriver_Temp.Trim();
            eDriverName_INS.Text = vDriverName_Temp.Trim();
        }

        protected void eAssignMan_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eAssignMan_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eAssignMan_INS");
            Label eAssignManName_INS = (Label)fvTowTruckAssignDetail.FindControl("eAssignManName_INS");
            string vAssignMan_Temp = eAssignMan_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vAssignMan_Temp + "' ";
            string vAssignManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vAssignManName_Temp == "")
            {
                vAssignManName_Temp = vAssignMan_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vAssignManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vAssignMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eAssignMan_INS.Text = vAssignMan_Temp.Trim();
            eAssignManName_INS.Text = vAssignManName_Temp.Trim();
            if (vAssignMan_Temp.Trim() != "")
            {
                Label eDepNo_Man_INS = (Label)fvTowTruckAssignDetail.FindControl("eDepNo_Man_INS");
                eDepNo_Man_INS.Text = PF.GetValue(vConnStr, "select DepNo from Employee where EmpNo = '" + vAssignMan_Temp.Trim() + "' ", "DepNo");
            }
        }

        protected void eFirstSupportMan_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eFirstSupportMan_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eFirstSupportMan_INS");
            Label eFirstSupportManName_INS = (Label)fvTowTruckAssignDetail.FindControl("eFirstSupportManName_INS");
            string vFirstSupportMan_Temp = eFirstSupportMan_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vFirstSupportMan_Temp + "' ";
            string vFirstSupportManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vFirstSupportManName_Temp == "")
            {
                vFirstSupportManName_Temp = vFirstSupportMan_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vFirstSupportManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vFirstSupportMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eFirstSupportMan_INS.Text = vFirstSupportMan_Temp.Trim();
            eFirstSupportManName_INS.Text = vFirstSupportManName_Temp.Trim();
        }

        protected void eSecondSupportMan_INS_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            TextBox eSecondSupportMan_INS = (TextBox)fvTowTruckAssignDetail.FindControl("eSecondSupportMan_INS");
            Label eSecondSupportManName_INS = (Label)fvTowTruckAssignDetail.FindControl("eSecondSupportManName_INS");
            string vSecondSupportMan_Temp = eSecondSupportMan_INS.Text.Trim();
            string vSQLStr_Temp = "select [Name] from Employee where EmpNo = '" + vSecondSupportMan_Temp + "' ";
            string vSecondSupportManName_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "Name");
            if (vSecondSupportManName_Temp == "")
            {
                vSecondSupportManName_Temp = vSecondSupportMan_Temp.Trim();
                vSQLStr_Temp = "select top 1 EmpNo from Employee where [Name] = '" + vSecondSupportManName_Temp.Trim() + "' order by isnull(LeaveDay, '9999/12/31') DESC";
                vSecondSupportMan_Temp = PF.GetValue(vConnStr, vSQLStr_Temp, "EmpNo");
            }
            eSecondSupportMan_INS.Text = vSecondSupportMan_Temp.Trim();
            eSecondSupportManName_INS.Text = vSecondSupportManName_Temp.Trim();
        }

        protected void ddlDamageAnalysis_INS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlDamageAnalysis_INS = (DropDownList)fvTowTruckAssignDetail.FindControl("ddlDamageAnalysis_INS");
            Label eDamageAnalysis_INS = (Label)fvTowTruckAssignDetail.FindControl("eDamageAnalysis_INS");
            eDamageAnalysis_INS.Text = ddlDamageAnalysis_INS.SelectedValue.Trim();
        }

        /// <summary>
        /// 刪除一筆記錄
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbDelete_List_Click(object sender, EventArgs e)
        {
            Label eCaseNo_List = (Label)fvTowTruckAssignDetail.FindControl("eCaseNo_List");
            if (eCaseNo_List != null)
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vCaseNo_Temp = eCaseNo_List.Text.Trim();
                string vExecSQL_Temp = "select Max(Items) MaxItems from TowTruckAssignList_History where CaseNo = '" + vCaseNo_Temp + "' ";
                string vMaxItems_Str = PF.GetValue(vConnStr, vExecSQL_Temp, "MaxItems");
                int vNewIndex = (vMaxItems_Str != "") ? Int32.Parse(vMaxItems_Str) + 1 : 1;
                string vCaseNoItem = vCaseNo_Temp + vNewIndex.ToString("D4");
                vExecSQL_Temp = "insert into TowTruckAssignList_History " + Environment.NewLine +
                                "       (CaseNoItem, Items, ModifyMode, CaseNo, CaseDate, Car_ID, Driver, DepNo_Car, CasePosition, " + Environment.NewLine +
                                "        ParkingPosition, CaseTime, AssignMan, DepNo_Man, FirstSupportMan, SecondSupportMan, " + Environment.NewLine +
                                "        Determination, FaultParts, FaultReason, Dispose, FollowUp, Improvements, Remark, " + Environment.NewLine +
                                "        BuMan, BuDate, ModifyMan, ModifyDate, IsChecked) " + Environment.NewLine +
                                "select '" + vCaseNoItem.Trim() + "', '" + vNewIndex.ToString("D4") + "', 'DEL', " + Environment.NewLine +
                                "       CaseNo, CaseDate, Car_ID, Driver, DepNo_Car, CasePosition, " + Environment.NewLine +
                                "       ParkingPosition, CaseTime, AssignMan, DepNo_Man, FirstSupportMan, SecondSupportMan, " + Environment.NewLine +
                                "       Determination, FaultParts, FaultReason, Dispose, FollowUp, Improvements, Remark, " + Environment.NewLine +
                                "       BuMan, BuDate, ModifyMan, ModifyDate, IsChecked " + Environment.NewLine +
                                "  from TowtruckAssignList " + Environment.NewLine +
                                " where CaseNo = '" + vCaseNo_Temp.Trim() + "' ";
                try
                {
                    PF.ExecSQL(vConnStr, vExecSQL_Temp);
                    sdsTowTruckAssignDetail.DeleteParameters.Clear();
                    sdsTowTruckAssignDetail.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp));
                    sdsTowTruckAssignDetail.Delete();
                }
                catch (Exception eMessage)
                {
                    Response.Write("<Script language='Javascript'>");
                    Response.Write("alert('" + eMessage.Message.Trim() + eMessage.ToString() + "')");
                    Response.Write("</" + "Script>");
                    //throw;
                }
            }
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plShowData.Visible = true;
            plSearch.Visible = true;
        }

        protected void sdsTowTruckAssignDetail_Deleted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridTowTruckAssignList.DataBind();
            }
        }

        protected void sdsTowTruckAssignDetail_Inserted(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridTowTruckAssignList.DataBind();
            }
        }

        protected void sdsTowTruckAssignDetail_Updated(object sender, SqlDataSourceStatusEventArgs e)
        {
            if (e.Exception == null)
            {
                gridTowTruckAssignList.DataBind();
            }
        }

        protected void cbIsChecked_Edit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsChecked_Edit = (CheckBox)fvTowTruckAssignDetail.FindControl("cbIsChecked_Edit");
            Label eIsChecked_Edit = (Label)fvTowTruckAssignDetail.FindControl("eIsChecked_Edit");
            eIsChecked_Edit.Text = (cbIsChecked_Edit.Checked) ? "V" : "";
        }

        protected void cbIsChecked_INS_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cbIsChecked_INS = (CheckBox)fvTowTruckAssignDetail.FindControl("cbIsChecked_INS");
            Label eIsChecked_INS = (Label)fvTowTruckAssignDetail.FindControl("eIsChecked_INS");
            eIsChecked_INS.Text = (cbIsChecked_INS.Checked) ? "V" : "";
        }
    }
}
