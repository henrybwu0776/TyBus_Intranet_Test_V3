using Amaterasu_Function;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Reporting.WebForms;

namespace TyBus_Intranet_Test_V3
{
    public partial class AnecdoteCase : Page
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
        private string vStationNo = "";
        private DateTime vToday = DateTime.Today;
        private DataTable dtShowData_A;
        private DataTable dtDataDetail_A;
        private DataTable dtShowData_B;
        private DataTable dtDataDetail_B;
        //資料繫結用的全域變數
        //A 檔表頸
        public string vdCaseNo_A;
        public string vdHasInsurance;
        public string vdDepNo;
        public string vdDepName;
        public string vdBuildDate;
        public string vdBuildMan;
        public string vdBuildManName;
        public string vdCar_ID;
        public string vdDriver;
        public string vdDriverName;
        public string vdInsuMan;
        public string vdAnecdotalResRatio;
        public string vdIsNoDeduction;
        public string vdDeductionDate;
        public string vdRemark_A;
        public string vdCaseOccurrence;
        public string vdERPCouseNo;
        public string vdCaseClose;
        public string vdIsExemption;
        public string vdPaidAmount;
        public string vdPenalty;
        public string vdPenaltyRatio;
        public string vdInsuAmount;
        public string vdIDCardNo;
        public string vdBirthday;
        public string vdAssumeday;
        public string vdTelephoneNo;
        public string vdAddress;
        public string vdPersonDamage;
        public string vdCarDamage;
        public string vdReportDate;
        public string vdCaseDate;
        public string vdCaseTime;
        public string vdOutReportNo;
        public string vdCasePosition;
        public string vdPoliceUnit;
        public string vdPoliceName;
        public string vdHasVideo;
        public string vdNoVideoReason;
        public string vdHasCaseData;
        public string vdModifyMan_A;
        public string vdModifyDate_A;
        public string vdHasAccReport;
        //B 檔明細
        public string vdCaseNo_B;
        public string vdItems;
        public string vdCaseNoItems;
        public string vdRelationship;
        public string vdRelCar_ID;
        public string vdEstimatedAmount;
        public string vdThirdInsurance;
        public string vdCompInsurance;
        public string vdDriverSharing;
        public string vdCompanySharing;
        public string vdCarDamageAMT;
        public string vdPersonDamageAMT;
        public string vdRelationComp;
        public string vdReconciliationDate;
        public string vdPassengerInsu;
        public string vdRemark_B;
        public string vdRelGender;
        public string vdRelTelNo1;
        public string vdRelTelNo2;
        public string vdRelCarType;
        public string vdRelPersonDamage;
        public string vdRelCarDamage;
        public string vdRelNote;
        public string vdModifyMan_B;
        public string vdModifyDate_B;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["LoginID"] != null)
            {
                //顯示民國年
                CultureInfo Cal = new CultureInfo("zh-TW");
                Cal.DateTimeFormat.Calendar = new TaiwanCalendar();
                Cal.DateTimeFormat.YearMonthPattern = "民國 yyy 年 MM 月";
                System.Threading.Thread.CurrentThread.CurrentCulture = Cal;

                //從 SESSION 取回使用者帳號資料
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
                    //有登入成功取得使用者帳號
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath); //取得資料庫連線字串
                    }
                    //定義畫面上事件日期欄位各自的 onclick 事件
                    //因為畫面日期的定義是民國年，所以對應的日期選擇視窗是 InputDate_ChineseYears.aspx
                    //查詢介面出險日期_起
                    string vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_Start_Search.ClientID;
                    string vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_Start_Search.Attributes["onClick"] = vCaseDateScript;
                    //查詢介面出險日期_迄
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eCaseDate_End_Search.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eCaseDate_End_Search.Attributes["onClick"] = vCaseDateScript;
                    //主檔_報告日期
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eReportDate.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eReportDate.Attributes["onClick"] = vCaseDateScript;
                    //主檔_出險日期
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBuildDate.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBuildDate.Attributes["onClick"] = vCaseDateScript;
                    //主檔_精勤減發日期
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eDeductionDate.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eDeductionDate.Attributes["onClick"] = vCaseDateScript;
                    //主檔_出生日期
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eBirthday.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eBirthday.Attributes["onClick"] = vCaseDateScript;
                    //明細_和解日期
                    vCaseDateURL = "InputDate_ChineseYears.aspx?TextboxID=" + eReconciliationDate.ClientID;
                    vCaseDateScript = "window.open('" + vCaseDateURL + "','','height=315, width=350,status=no,toolbar=no,menubar=no,location=no','')";
                    eReconciliationDate.Attributes["onClick"] = vCaseDateScript;

                    if (!IsPostBack)
                    {
                        fuExcel.Visible = (vLoginDepNo == "09");
                        bbImportData.Visible = (vLoginDepNo == "09");
                        vStationNo = PF.GetValue(vConnStr, "select DepNo from Department where DepNo = '" + vLoginDepNo + "' and InSHReport = 'V'", "DepNo");

                        eDepNo_Start_Search.Text = vStationNo.Trim();
                        eDepNo_End_Search.Text = "";
                        eDepNo_Start_Search.Enabled = (vStationNo == "");
                        eDepNo_End_Search.Enabled = (vStationNo == "");

                        plPrint.Visible = false;
                        plSearch.Visible = true;
                        plMainDataShow.Visible = true;
                        plMainData.Visible = false;
                        plDetailDataShow.Visible = true;
                        plDetailData.Visible = false;
                        Session["AnecdoteMode"] = "LIST";
                        MainButtonStatus("LIST");
                        SubButtonStatus("LIST");
                        MainTextboxStatus("LIST");
                        SubTextboxStatus("LIST");
                    }
                    else
                    {
                        //進行資料繫結
                        CaseDataBind();
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
        /// 取回查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr()
        {
            string vHasInsuStr = (chkHasInsurance.Checked) ? " and isnull(HasInsurance, 0) = 1" + Environment.NewLine : " and isnull(HasInsurance, 0) = 0" + Environment.NewLine;
            string vCaseCloseStr = (chkCaseClose.Checked) ? " and isnull(CaseClose, 0) = 1" + Environment.NewLine : " and isnull(CaseClose, 0) = 0" + Environment.NewLine;
            string vDepNoWhereStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? " and DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                    "";
            string vCaseDateWhereStr = ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() != "")) ?
                                       " and BuildDate between '" + PF.TransDateString(eCaseDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eCaseDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() == "")) ?
                                       " and BuildDate = '" + PF.TransDateString(eCaseDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eCaseDate_Start_Search.Text.Trim() == "") && (eCaseDate_End_Search.Text.Trim() != "")) ?
                                       " and BuildDate = '" + PF.TransDateString(eCaseDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       "";
            string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? " and Car_ID like '%" + eCarID_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vDriverStr = (eDriver_Search.Text.Trim() != "") ? " and Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "SELECT CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, " + Environment.NewLine +
                                "       (SELECT NAME FROM Employee WHERE (EMPNO = AnecdoteCase.BuildMan)) AS BuildManName, " + Environment.NewLine +
                                "       Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, " + Environment.NewLine +
                                //2024.05.08 新增欄位
                                //"       IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose " + Environment.NewLine +
                                "       IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose, " + Environment.NewLine +
                                "       IsExemption, PaidAmount, InsuAmount, Penalty, PenaltyRatio " + Environment.NewLine +
                                "  FROM AnecdoteCase " + Environment.NewLine +
                                " where isnull(CaseNo, '') <> '' " + Environment.NewLine +
                                vHasInsuStr +
                                vCaseCloseStr +
                                vDepNoWhereStr +
                                vCaseDateWhereStr +
                                vCarIDStr +
                                vDriverStr +
                                " order by CaseNo DESC";
            return vResultStr;
        }

        protected void CaseDataBind()
        {
            if (vLoginID != "")
            {
                if (vConnStr == "")
                {
                    vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                }
                string vSelectStr = GetSelectStr();
                using (SqlConnection connMainData = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daMainData = new SqlDataAdapter(vSelectStr, connMainData);
                    if (dtShowData_A != null)
                    {
                        dtShowData_A.Clear();
                    }
                    else
                    {
                        dtShowData_A = new DataTable();
                    }
                    connMainData.Open();
                    daMainData.Fill(dtShowData_A);
                    gridAnecdoteCaseA_List.DataSourceID = "";
                    gridAnecdoteCaseA_List.DataSource = dtShowData_A;
                    gridAnecdoteCaseA_List.DataBind();
                    plMainData.Visible = (dtShowData_A.Rows.Count > 0);
                }
            }
        }

        /// <summary>
        /// 取回EXCEL用的查詢字串
        /// </summary>
        /// <returns></returns>
        private string GetSelectStr_Excel()
        {
            string vHasInsuStr = (chkHasInsurance.Checked) ? " and HasInsurance = 1" + Environment.NewLine : " and isnull(HasInsurance, 0) = 0" + Environment.NewLine;
            string vCaseCloseStr = (chkCaseClose.Checked) ? " and CaseClose = 1" + Environment.NewLine : " and isnull(CaseClose, 0) = 0" + Environment.NewLine;
            string vDepNoWhereStr = ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and DepNo between '" + eDepNo_Start_Search.Text.Trim() + "' and '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDepNo_Start_Search.Text.Trim() != "") && (eDepNo_End_Search.Text.Trim() == "")) ? " and DepNo = '" + eDepNo_Start_Search.Text.Trim() + "' " + Environment.NewLine :
                                    ((eDepNo_Start_Search.Text.Trim() == "") && (eDepNo_End_Search.Text.Trim() != "")) ? " and DepNo = '" + eDepNo_End_Search.Text.Trim() + "' " + Environment.NewLine :
                                    "";
            string vCaseDateWhereStr = ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() != "")) ?
                                       " and BuildDate between '" + PF.TransDateString(eCaseDate_Start_Search.Text.Trim(), "B") + "' and '" + PF.TransDateString(eCaseDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eCaseDate_Start_Search.Text.Trim() != "") && (eCaseDate_End_Search.Text.Trim() == "")) ?
                                       " and BuildDate = '" + PF.TransDateString(eCaseDate_Start_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       ((eCaseDate_Start_Search.Text.Trim() == "") && (eCaseDate_End_Search.Text.Trim() != "")) ?
                                       " and BuildDate = '" + PF.TransDateString(eCaseDate_End_Search.Text.Trim(), "B") + "' " + Environment.NewLine :
                                       "";
            string vCarIDStr = (eCarID_Search.Text.Trim() != "") ? " and Car_ID like '%" + eCarID_Search.Text.Trim() + "%' " + Environment.NewLine : "";
            string vDriverStr = (eDriver_Search.Text.Trim() != "") ? " and Driver = '" + eDriver_Search.Text.Trim() + "' " + Environment.NewLine : "";
            string vResultStr = "SELECT CaseNo, " + Environment.NewLine +
                                "       case when HasInsurance = 1 then '是' else '否' end HasInsurance, " + Environment.NewLine +
                                "       DepNo, DepName, BuildDate, BuildMan, " + Environment.NewLine +
                                "       (SELECT NAME FROM Employee WHERE (EMPNO = AnecdoteCase.BuildMan)) AS BuildManName, " + Environment.NewLine +
                                "       Car_ID, Driver, DriverName, InsuMan, AnecdotalResRatio, " + Environment.NewLine +
                                "       case when IsNoDeduction = 1 then '是' else '否' end IsNoDeduction, " + Environment.NewLine +
                                //2024.05.08 新增欄位
                                //"       DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose " + Environment.NewLine +
                                "       DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose, " + Environment.NewLine +
                                "       IsExemption, PaidAmount, InsuAmount, Penalty, PenaltyRatio " + Environment.NewLine +
                                "  FROM AnecdoteCase " + Environment.NewLine +
                                " where isnull(CaseNo, '') <> '' " + Environment.NewLine +
                                vHasInsuStr +
                                vCaseCloseStr +
                                vDepNoWhereStr +
                                vCaseDateWhereStr +
                                vCarIDStr +
                                vDriverStr +
                                " order by CaseNo DESC";
            return vResultStr;
        }

        /// <summary>
        /// 查詢條件_駕駛員
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eDriver_Search_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDriver = eDriver_Search.Text.Trim();
            string vDriverName = "";
            string vTempStr_Driver = "select [Name] from Employee where EmpNo = '" + vDriver + "' ";
            vDriverName = PF.GetValue(vConnStr, vTempStr_Driver, "Name");
            if (vDriverName.Trim() == "")
            {
                vDriverName = vDriver;
                vTempStr_Driver = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName + "' order by Assumeday DESC";
                vDriver = PF.GetValue(vConnStr, vTempStr_Driver, "EmpNo");
            }
            eDriver_Search.Text = vDriver.Trim();
            eDriverName_Search.Text = vDriverName.Trim();
        }

        /// <summary>
        /// 從 EXCEL 匯入資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbImportData_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            CaseDataBind();
        }

        protected void bbExcel_Click(object sender, EventArgs e)
        {

        }

        protected void bbCancel_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void eDepNo_TextChanged(object sender, EventArgs e)
        {
            string vDepNo = eDepNo.Text.Trim();
            string vTempStr_DepNo = "select [Name] from Department where DeNo = '" + vDepNo + "' ";
            string vDepName = PF.GetValue(vConnStr, vTempStr_DepNo, "Name");
            if (string.IsNullOrEmpty(vDepName))
            {
                vDepName = vDepNo.Trim();
                vTempStr_DepNo = "select top 1 DepNo deom Department where [Name] = '" + vDepName + "' order by DepNo ";
                vDepNo = PF.GetValue(vConnStr, vTempStr_DepNo, "DepNo");
            }
            eDepNo.Text = vDepNo.Trim();
            eDepName.Text = vDepName.Trim();
        }

        /// <summary>
        /// 主檔按鈕狀態
        /// </summary>
        /// <param name="fButtonMode"></param>
        private void MainButtonStatus(string fButtonMode)
        {
            bbNew_Main.Visible = (fButtonMode == "LIST");
            bbEdit_Main.Visible = ((fButtonMode == "LIST") && (gridAnecdoteCaseA_List.SelectedIndex != -1));
            bbDel_Main.Visible = ((fButtonMode == "LIST") && (gridAnecdoteCaseA_List.SelectedIndex != -1));
            bbPrintNote.Visible = (fButtonMode == "LIST");
            bbPrintReport.Visible = (fButtonMode == "LIST");
            bbERPSync.Visible = ((fButtonMode == "LIST") && (gridAnecdoteCaseA_List.SelectedIndex != -1));
            bbOK_Main.Visible = !(fButtonMode == "LIST");
            bbCancel_Main.Visible = !(fButtonMode == "LIST");
        }

        /// <summary>
        /// 主檔文字輸入狀態
        /// </summary>
        /// <param name="fMode"></param>
        private void MainTextboxStatus(string fMode)
        {
            foreach (Control vItems in plMainData.Controls)
            {
                if (vItems is TextBox)
                {
                    (vItems as TextBox).Enabled = (fMode.ToUpper() != "LIST");
                    if (fMode.ToUpper() == "INS") (vItems as TextBox).Text = "";
                }
                else if (vItems is CheckBox)
                {
                    (vItems as CheckBox).Enabled = (fMode.ToUpper() != "LIST");
                }
            }
        }

        /// <summary>
        /// 明細按鈕狀態
        /// </summary>
        /// <param name="fButtonMode"></param>
        private void SubButtonStatus(string fButtonMode)
        {
            bbNew_B.Visible = (fButtonMode == "LIST");
            bbEdit_B.Visible = ((fButtonMode == "LIST") && (gridAnecdoteCaseB_List.SelectedIndex != -1));
            bbDel_B.Visible = ((fButtonMode == "LIST") && (gridAnecdoteCaseB_List.SelectedIndex != -1));
            bbOK_B.Visible = !(fButtonMode == "LIST");
            bbCancel_B.Visible = !(fButtonMode == "LIST");
        }

        /// <summary>
        /// 明細文字輸入狀態
        /// </summary>
        /// <param name="fMode"></param>
        private void SubTextboxStatus(string fMode)
        {
            foreach (Control vItems in plDetailData.Controls)
            {
                if (vItems is TextBox)
                {
                    (vItems as TextBox).Enabled = (fMode.ToUpper() != "LIST");
                    if (fMode.ToUpper() == "INS") (vItems as TextBox).Text = "";
                }
                else if (vItems is CheckBox)
                {
                    (vItems as CheckBox).Enabled = (fMode.ToUpper() != "LIST");
                }
            }
        }

        /// <summary>
        /// 根據駕駛員工號或姓名取回相關資訊
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eDriver_TextChanged(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }

            string vDriver = eDriver.Text.Trim();
            string vTempStr_Driver = "select [Name] from Employee where EmpNo = '" + vDriver + "' and Type = '20' ";
            string vDriverName = PF.GetValue(vConnStr, vTempStr_Driver, "Name");
            if (string.IsNullOrEmpty(vDriverName))
            {
                vDriverName = eDriver.Text.Trim();
                vTempStr_Driver = "select top 1 EmpNo from Employee where [Name] = '" + vDriverName + "' order by Assumeday DESC";
                vDriver = PF.GetValue(vConnStr, vTempStr_Driver, "EmpNo");
            }
            if (vDriver.Trim() != "")
            {
                vTempStr_Driver = "select e.EmpNo, e.[Name] as EmpName, convert(varchar(10), e.Birthday, 111) as Birthday, convert(varchar(10), e.Assumeday, 111) as Assumeday, " + Environment.NewLine +
                                  "       e.Addr1, e.IDCardNo, case when e.Cellphone is not null then e.CellPhone else e.TelNo1 end as DriverTelNo, " + Environment.NewLine +
                                  "       e.DepNo, d.[Name] as DepName " + Environment.NewLine +
                                  "  from Employee e left join department d on d.DepNo = e.DepNo " + Environment.NewLine +
                                  " where e.EmpNo = '" + vDriver + "' and e.Type = '20'";
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdTemp = new SqlCommand(vTempStr_Driver, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    if (drTemp.HasRows)
                    {
                        while (drTemp.Read())
                        {
                            eDriverName.Text = drTemp["EmpName"].ToString().Trim();
                            eBirthday.Text = drTemp["Birthday"].ToString().Trim();
                            eAssumeday.Text = drTemp["Assumeday"].ToString().Trim();
                            eAddress.Text = drTemp["Arrd1"].ToString().Trim();
                            eIDCardNo.Text = drTemp["IDCardNo"].ToString().Trim();
                            eTelephoneNo.Text = drTemp["DriverTelNo"].ToString().Trim();
                            eDepNo.Text = drTemp["DepNo"].ToString().Trim();
                            eDepName.Text = drTemp["DepName"].ToString().Trim();
                        }
                    }
                }
            }
            else
            {
                eErrorMSG_A.Text = "查無此駕駛，請重新查詢";
                eErrorMSG_A.Visible = true;
                eDriver.Text = "";
                eDriver.Focus();
            }
        }

        /// <summary>
        /// 根據車號帶點駕駛員相關資訊
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eCar_ID_TextChanged(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCar_ID = eCar_ID.Text.Trim();
            string vTempStr_Car_ID = "select a.Car_ID, a.CompanyNo, d.[Name] as DepName, a.Driver, e.[Name] as DriverName, e.IDCardNo, " + Environment.NewLine +
                                     "       convert(varchar(10), e.Birthday, 111) as Birthday, e.[Addr1], " + Environment.NewLine +
                                     "       convert(varchar(10), e.Assumeday, 111) as Assumeday, " + Environment.NewLine +
                                     "       case when e.CellPhone is not null then e.CellPhone else e.TelNo1 end DriverTelNo " + Environment.NewLine +
                                     "  from Car_InfoA a left join Department d on d.DepNo = a.CompanyNo " + Environment.NewLine +
                                     "                   left join Employee e on e.EmpNo = a.Driver " + Environment.NewLine +
                                     " where Car_ID = '" + vCar_ID + "'";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vTempStr_Car_ID, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    while (drTemp.Read())
                    {
                        eDriver.Text = drTemp["Driver"].ToString().Trim();
                        eDriverName.Text = drTemp["DriverName"].ToString().Trim();
                        eDepNo.Text = drTemp["CompanyNo"].ToString().Trim();
                        eDepName.Text = drTemp["DepName"].ToString().Trim();
                        eBirthday.Text = drTemp["Birthday"].ToString().Trim();
                        eIDCardNo.Text = drTemp["IDCardNo"].ToString().Trim();
                        eAddress.Text = drTemp["Addr1"].ToString().Trim();
                        eAssumeday.Text = drTemp["Assumeday"].ToString().Trim();
                        eTelephoneNo.Text = drTemp["DriverTelNo"].ToString().Trim();
                    }
                }
                else
                {
                    eCar_ID.Text = "";
                    eCar_ID.Focus();
                    eErrorMSG_A.Text = "輸入車號不正確，查無此車號";
                    eErrorMSG_A.Visible = true;
                }
            }
        }

        protected void cbHasInsurance_CheckedChanged(object sender, EventArgs e)
        {
            eHasInsurance.Text = (cbHasInsurance.Checked) ? "True" : "False";
        }

        protected void cbCaseClose_CheckedChanged(object sender, EventArgs e)
        {
            eCaseClose.Text = (cbCaseClose.Checked) ? "True" : "False";
        }

        /// <summary>
        /// 變更建檔人
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void eBuildMan_TextChanged(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vBuildMan = eBuildMan.Text.Trim();
            string vTempStr_BuildMan = "select [Name] from Employee where EmpNo = '" + vBuildMan + "' ";
            string vBuildManName = PF.GetValue(vConnStr, vTempStr_BuildMan, "Name");
            if (string.IsNullOrEmpty(vBuildManName))
            {
                vBuildManName = eBuildMan.Text.Trim();
                vTempStr_BuildMan = "select top 1 EmpNo from Employee where [Name] = '" + vBuildManName + "' order by Assumeday DESC";
                vBuildMan = PF.GetValue(vConnStr, vTempStr_BuildMan, "EmpNo");
            }
            eBuildMan.Text = vBuildMan.Trim();
            eBuildManName.Text = vBuildManName.Trim();
        }

        protected void cbIsExemption_CheckedChanged(object sender, EventArgs e)
        {
            eIsExemption.Text = (cbIsExemption.Checked) ? "Y" : "X";
        }

        protected void cbIsNoDeduction_CheckedChanged(object sender, EventArgs e)
        {
            eIsNoDeduction.Text = (cbIsNoDeduction.Checked) ? "True" : "False";
        }

        protected void cbHasVideo_CheckedChanged(object sender, EventArgs e)
        {
            eHasVideo.Text = (cbHasVideo.Checked) ? "Y" : "X";
        }

        protected void cbHasAccReport_CheckedChanged(object sender, EventArgs e)
        {
            eHasAccReport.Text = (cbHasAccReport.Checked) ? "Y" : "X";
        }

        protected void gridAnecdoteCaseA_List_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridAnecdoteCaseA_List.PageIndex = e.NewPageIndex;
            gridAnecdoteCaseA_List.DataBind();
        }

        /// <summary>
        /// 
        /// </summary>
        private void OpenData_A()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCaseNo = gridAnecdoteCaseA_List.SelectedDataKey.Value.ToString().Trim();
            if (!string.IsNullOrEmpty(vCaseNo))
            {
                string vSelectStr_CaseA = "select a.CaseNo, a.HasInsurance, a.DepNo, a.DepName, convert(varchar(10), a.BuildDate, 111) as BuildDate, a.BuildMan, " + Environment.NewLine +
                                          "       e.[Name] as BuildManName, a.Car_ID, a.Driver, a.DriverName, a.InsuMan, a.AnecdotalResRatio, a.IsNoDeduction, " + Environment.NewLine +
                                          "       convert(varchar(10), a.DeductionDate, 111) as DeductionDate, a.Remark as Remark_A, a.CaseOccurrence, a.ERPCouseNo, a.CaseClose, " + Environment.NewLine +
                                          "       a.IsExemption, a.PaidAmount, a.Penalty, a.PenaltyRatio, a.InsuAmount, a.IDCardNo, convert(varchar(10), a.Birthday, 111) as Birthday, " + Environment.NewLine +
                                          "       convert(varchar(10), a.Assumeday, 111) as Assumeday, a.TelephoneNo, a.[Address],a.PersonDamage, a.CarDamage,  " + Environment.NewLine +
                                          "       convert(varchar(10), a.ReportDate, 111) as ReportDate, convert(varchar(10), a.CaseDate, 111) as CaseDate, " + Environment.NewLine +
                                          "       a.CaseTime, a.OutReportNo, a.CasePosition, a.PoliceUnit, a.PoliceName, a.HasVideo, a.NoVideoReason, a.HasCaseData, " + Environment.NewLine +
                                          "       a.HasAccReport, a.ModifyMan, convert(varchar(10), a.ModifyDate, 111) as ModifyDate " + Environment.NewLine +
                                          "  from AnecdoteCase a left join Employee e on e.EmpNo = a.BuildMan " + Environment.NewLine +
                                          " where a.CaseNo = @CaseNo ";
                using (SqlConnection connDetailA = new SqlConnection(vConnStr))
                {
                    if (dtDataDetail_A != null)
                    {
                        dtDataDetail_A.Clear();
                    }
                    else
                    {
                        dtDataDetail_A = new DataTable();
                    }
                    SqlCommand cmdDetailA = new SqlCommand(vSelectStr_CaseA, connDetailA);
                    cmdDetailA.Parameters.Clear();
                    cmdDetailA.Parameters.Add(new SqlParameter("CaseNo", vCaseNo));
                    connDetailA.Open();
                    SqlDataReader drDetailA = cmdDetailA.ExecuteReader();
                    if (drDetailA.HasRows)
                    {
                        while (drDetailA.Read())
                        {
                            vdCaseNo_A = drDetailA["CaseNo"].ToString().Trim();
                            vdHasInsurance = drDetailA["HasInsurance"].ToString().Trim();
                            cbHasInsurance.Checked = (vdHasInsurance.ToLower() == "true");
                            vdDepNo = drDetailA["DepNo"].ToString().Trim();
                            vdDepName = drDetailA["DepName"].ToString().Trim();
                            vdBuildDate = drDetailA["BuildDate"].ToString().Trim();
                            vdBuildMan = drDetailA["BuildMan"].ToString().Trim();
                            vdBuildManName = drDetailA["BuildManName"].ToString().Trim();
                            vdCar_ID = drDetailA["Car_ID"].ToString().Trim();
                            vdDriver = drDetailA["Driver"].ToString().Trim();
                            vdDriverName = drDetailA["DriverName"].ToString().Trim();
                            vdInsuMan = drDetailA["InsuMan"].ToString().Trim();
                            vdAnecdotalResRatio = drDetailA["AnecdotalResRatio"].ToString().Trim();
                            vdIsNoDeduction = drDetailA["IsNodeduction"].ToString().Trim();
                            cbIsNoDeduction.Checked = (vdIsNoDeduction.ToLower() == "true");
                            vdDeductionDate = drDetailA["DeductionDate"].ToString().Trim();
                            vdRemark_A = drDetailA["Remark_A"].ToString().Trim();
                            vdCaseOccurrence = drDetailA["CaseOccurrence"].ToString().Trim();
                            vdERPCouseNo = drDetailA["ERPCouseNo"].ToString().Trim();
                            vdCaseClose = drDetailA["CaseClose"].ToString().Trim();
                            cbCaseClose.Checked = (vdCaseClose.ToLower() == "true");
                            vdIsExemption = drDetailA["IsExemption"].ToString().Trim();
                            cbIsExemption.Checked = (vdIsExemption == "Y");
                            vdPaidAmount = drDetailA["PaidAmount"].ToString().Trim();
                            vdPenalty = drDetailA["Penalty"].ToString().Trim();
                            vdPenaltyRatio = drDetailA["PenaltyRatio"].ToString().Trim();
                            vdInsuAmount = drDetailA["InsuAmount"].ToString().Trim();
                            vdIDCardNo = drDetailA["IDCardNo"].ToString().Trim();
                            vdBirthday = drDetailA["Birthday"].ToString().Trim();
                            vdAssumeday = drDetailA["Assumeday"].ToString().Trim();
                            vdTelephoneNo = drDetailA["TelephoneNo"].ToString().Trim();
                            vdAddress = drDetailA["Address"].ToString().Trim();
                            vdPersonDamage = drDetailA["PersonDamage"].ToString().Trim();
                            vdCarDamage = drDetailA["CarDamage"].ToString().Trim();
                            vdReportDate = drDetailA["ReportDate"].ToString().Trim();
                            vdCaseDate = drDetailA["CaseDate"].ToString().Trim();
                            vdCaseTime = drDetailA["CaseTime"].ToString().Trim();
                            vdOutReportNo = drDetailA["OutReportNo"].ToString().Trim();
                            vdCasePosition = drDetailA["CasePosition"].ToString().Trim();
                            vdPoliceUnit = drDetailA["PoliceUnit"].ToString().Trim();
                            vdPoliceName = drDetailA["PoliceName"].ToString().Trim();
                            vdHasVideo = drDetailA["HasVideo"].ToString().Trim();
                            cbHasVideo.Checked = (vdHasVideo == "Y");
                            vdNoVideoReason = drDetailA["NoVideoReason"].ToString().Trim();
                            vdHasCaseData = drDetailA["HasCaseData"].ToString().Trim();
                            rbHasCaseData.SelectedIndex = (vdHasCaseData != "") ? Int32.Parse(vdHasCaseData) : 1;
                            vdModifyMan_A = drDetailA["ModifyMan"].ToString().Trim();
                            vdModifyDate_A = drDetailA["ModifyDate"].ToString().Trim();
                            vdHasAccReport = drDetailA["HasAccReport"].ToString().Trim();
                            cbHasAccReport.Checked = (vdHasAccReport == "Y");
                            plMainData.DataBind();
                        }
                    }
                }
                //檢查是不是有明細資料
                vSelectStr_CaseA = "select CaseNo, Items, CaseNoItems, Relationship, RelCarType, RelCar_ID, RelTelNo1, " + Environment.NewLine +
                                   "       convert(varchar(10), ReconciliationDate, 111) ReconciliationDate " + Environment.NewLine +
                                   "  from AnecdoteCaseB " + Environment.NewLine +
                                   " where CaseNo = @CaseNo_B ";
                using (SqlConnection connListB = new SqlConnection(vConnStr))
                {
                    SqlDataAdapter daList_B = new SqlDataAdapter(vSelectStr_CaseA, connListB);
                    daList_B.SelectCommand.Parameters.Clear();
                    daList_B.SelectCommand.Parameters.Add(new SqlParameter("CaseNo_B", vCaseNo));
                    connListB.Open();
                    DataTable dtList_B = new DataTable();
                    daList_B.Fill(dtList_B);
                    gridAnecdoteCaseB_List.DataSource = dtList_B;
                    gridAnecdoteCaseB_List.DataBind();
                }
            }
        }

        /// <summary>
        /// 展示主檔資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gridAnecdoteCaseA_List_SelectedIndexChanged(object sender, EventArgs e)
        {
            MainButtonStatus("LIST");
            SubButtonStatus("LIST");
            MainTextboxStatus("LIST");
            SubTextboxStatus("LIST");
            OpenData_A();
        }

        protected void bbCloseReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = false;
            plMainDataShow.Visible = true;
            plDetailDataShow.Visible = true;
        }

        protected void gridAnecdoteCaseB_List_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridAnecdoteCaseB_List.PageIndex = e.NewPageIndex;
            gridAnecdoteCaseB_List.DataBind();
        }

        /// <summary>
        /// 顯示 B 檔資料
        /// </summary>
        private void OpenData_B()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (gridAnecdoteCaseB_List.SelectedDataKey != null)
            {
                string vCaseNoItems = gridAnecdoteCaseB_List.SelectedDataKey.Value.ToString().Trim();
                if (!string.IsNullOrEmpty(vCaseNoItems))
                {
                    string vSelectStr_CaseB = "SELECT CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, " + Environment.NewLine +
                                              "       DriverSharing, CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, " + Environment.NewLine +
                                              "       PassengerInsu, Remark, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, " + Environment.NewLine +
                                              "       RelNote, ModifyMan, ModifyDate " + Environment.NewLine +
                                              "  from dbo.AnecdoteCaseB " + Environment.NewLine +
                                              " where CaseNoItems = @CaseNoItems ";
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSelectStr_CaseB, connTemp);
                        connTemp.Open();
                        cmdTemp.Parameters.Clear();
                        cmdTemp.Parameters.Add(new SqlParameter("CaseNoItems", vCaseNoItems));
                        SqlDataReader drTemp = cmdTemp.ExecuteReader();
                        if (drTemp.HasRows)
                        {
                            while (drTemp.Read())
                            {
                                vdCaseNo_B = drTemp["CaseNo"].ToString().Trim();
                                vdItems = drTemp["Items"].ToString().Trim();
                                vdCaseNoItems = drTemp["CaseNoItems"].ToString().Trim();
                                vdRelationship = drTemp["Relationship"].ToString().Trim();
                                vdRelCar_ID = drTemp["RelCar_ID"].ToString().Trim();
                                vdEstimatedAmount = (drTemp["EstimatedAmount"] != null) ? drTemp["EstimatedAmount"].ToString().Trim() : "";
                                vdThirdInsurance = drTemp["ThirdInsurance"].ToString().Trim();
                                vdCompInsurance = drTemp["CompInsurance"].ToString().Trim();
                                vdDriverSharing = (drTemp["DriverSharing"] != null) ? drTemp["DriverSharing"].ToString().Trim() : "";
                                vdCompanySharing = (drTemp["CompanySharing"] != null) ? drTemp["CompanySharing"].ToString().Trim() : "";
                                vdCarDamageAMT = (drTemp["CarDamageAMT"] != null) ? drTemp["CarDamageAMT"].ToString().Trim() : "";
                                vdPersonDamageAMT = (drTemp["PersonDamageAMT"] != null) ? drTemp["PersonDamageAMT"].ToString().Trim() : "";
                                vdRelationComp = drTemp["RelationComp"].ToString().Trim();
                                vdReconciliationDate = drTemp["ReconciliationDate"].ToString().Trim();
                                vdPassengerInsu = drTemp["PassengerInsu"].ToString().Trim();
                                vdRemark_B = drTemp["Remark"].ToString().Trim();
                                vdRelGender = drTemp["RelGender"].ToString().Trim();
                                vdRelTelNo1 = drTemp["RelTelNo1"].ToString().Trim();
                                vdRelTelNo2 = drTemp["RelTelNo2"].ToString().Trim();
                                vdRelCarType = drTemp["RelCarType"].ToString().Trim();
                                vdRelPersonDamage = drTemp["RelPersonDamage"].ToString().Trim();
                                vdRelCarDamage = drTemp["RelCarDamage"].ToString().Trim();
                                vdRelNote = drTemp["RelNote"].ToString().Trim();
                                vdModifyMan_B = drTemp["ModifyMan"].ToString().Trim();
                                DateTime vModifyDate_B;
                                vdModifyDate_B = DateTime.TryParse(drTemp["ModifyDate"].ToString().Trim(), out vModifyDate_B) ? vModifyDate_B.ToShortDateString() : "";
                            }
                            plDetailData.DataBind();
                            plDetailData.Visible = true;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 展示明細資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gridAnecdoteCaseB_List_SelectedIndexChanged(object sender, EventArgs e)
        {
            MainButtonStatus("LIST");
            SubButtonStatus("LIST");
            MainTextboxStatus("LIST");
            SubTextboxStatus("LIST");
            OpenData_B();
        }

        protected void ddlRelCarType_SelectedIndexChanged(object sender, EventArgs e)
        {
            eRelCarType.Text = ddlRelCarType.SelectedValue.Trim();
        }

        protected void rbHasCaseData_SelectedIndexChanged(object sender, EventArgs e)
        {
            eHasCaseData.Text = (rbHasCaseData.SelectedIndex != -1) ? rbHasCaseData.SelectedIndex.ToString().Trim() : "0";
        }

        /// <summary>
        /// 主檔新增
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbNew_Main_Click(object sender, EventArgs e)
        {
            Session["AnecdoteMode"] = "INS";
            MainButtonStatus("INS");
            SubButtonStatus("LIST");
            MainTextboxStatus("INS");
            SubTextboxStatus("LIST");
            eReportDate.Text = (vToday.Year - 1911).ToString() + "/" + vToday.ToString("MM/dd");
            eInsuMan.Focus();
        }

        /// <summary>
        /// 主檔修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbEdit_Main_Click(object sender, EventArgs e)
        {
            Session["AnecdoteMode"] = "EDIT";
            MainButtonStatus("EDIT");
            SubButtonStatus("LIST");
            MainTextboxStatus("EDIT");
            SubTextboxStatus("LIST");
            eReportDate.Focus();
        }

        protected void bbDel_Main_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            if (eCaseNo_A.Text.Trim() != "")
            {
                try
                {
                    //取回要刪除的肇事單號
                    string vCaseNo_Temp = eCaseNo_A.Text.Trim();
                    string vRecordNote = "刪除肇事單資料_" + vCaseNo_Temp + Environment.NewLine +
                                         "刪除日期：" + vToday.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                         "肇事單號：[" + vCaseNo_Temp + "]" + Environment.NewLine +
                                         "刪除人工號：" + vLoginID + Environment.NewLine +
                                         "AnecdoteCase.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //因為是刪主檔，所以要連 B 檔和 C 檔都一起刪，可是 C 檔沒有異動檔所以只針對 B 檔做異動
                    string vHistoryNo = vToday.Year.ToString("D4") + vToday.Month.ToString("D2") + "DELA";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndex_Str = vMaxNo.Replace(vHistoryNo, "").Trim();
                    int vIndex = (vIndex_Str != "") ? Int32.Parse(vIndex_Str) + 1 : 1;
                    string vNewHistoryNo = vHistoryNo + vIndex.ToString("D4");
                    vSQLStr_Temp = "select max(ItemsH) MaxItem from AnecdoteCaseBHistory where HistoryNo = '" + vNewHistoryNo + "' ";
                    string vMaxItemsH = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    string vNewItemsH = (vMaxItemsH.Trim() != "") ? Int32.Parse(vMaxItemsH).ToString("D4") : "0001";
                    string vNewHistoryNoItems = vNewHistoryNo.Trim() + vNewItemsH.Trim();
                    //複製 B 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseBHistory] " + Environment.NewLine +
                                   "            ([HistoryNo], [ItemsH], [HistoryNoItems], [CaseNo], [Items], [CaseNoItems], [Relationship], [RelCar_ID], " + Environment.NewLine +
                                   "             [EstimatedAmount], [ThirdInsurance], [CompInsurance], [DriverSharing], [CompanySharing], [CarDamageAMT], " + Environment.NewLine +
                                   "             [PersonDamageAMT], [RelationComp], [ReconciliationDate], [Remark], [ModifyType], [ModifyDate], [ModifyMan], " + Environment.NewLine +
                                   "             [PassengerInsu], [RelGender], [RelTelNo1], [RelTelNo2], [RelCarType], [RelPersonDamage], [RelCarDamage], " + Environment.NewLine +
                                   "             [RelNote])" + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', '" + vNewItemsH + "', '" + vNewHistoryNoItems + "', [CaseNo], [Items], [CaseNoItems], [Relationship], " + Environment.NewLine +
                                   "       [RelCar_ID], [EstimatedAmount], [ThirdInsurance], [CompInsurance], [DriverSharing], [CompanySharing], [CarDamageAMT], " + Environment.NewLine +
                                   "       [PersonDamageAMT], [RelationComp], [ReconciliationDate], [Remark], 'DELA', GetDate(), '" + vLoginID + "', [PassengerInsu], " + Environment.NewLine +
                                   "       [RelGender], [RelTelNo1], [RelTelNo2], [RelCarType], [RelPersonDamage], [RelCarDamage], [RelNote] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCaseB] " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //複製 A 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseHistory] " + Environment.NewLine +
                                   "            ([HistoryNo], [CaseNo], [HasInsurance], [DepNo], [DepName], [BuildDate], [BuildMan], [Car_ID], [Driver], " + Environment.NewLine +
                                   "             [DriverName], [InsuMan], [AnecdotalResRatio], [IsNoDeduction], [DeductionDate], [Remark], [ModifyType], " + Environment.NewLine +
                                   "             [ModifyDate], [ModifyMan], [CaseOccurrence], [ERPCouseNo], [CaseClose], [IsExemption], [PaidAmount], " + Environment.NewLine +
                                   "             [Penalty], [PenaltyRatio], [InsuAmount], [IDCardNo], [Birthday], [Assumeday], [TelephoneNo], [Address], " + Environment.NewLine +
                                   "             [PersonDamage], [CarDamage], [ReportDate], [CaseDate], [CaseTime], [OutReportNo], [PoliceUnit], [PoliceName], " + Environment.NewLine +
                                   "             [HasVideo], [NoVideoReason], [HasCaseData]) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', [CaseNo], [HasInsurance], [DepNo], [DepName], [BuildDate], [BuildMan], [Car_ID], [Driver], " + Environment.NewLine +
                                   "       [DriverName], [InsuMan], [AnecdotalResRatio], [IsNoDeduction], [DeductionDate], [Remark], 'DELA', GetDate(), '" + vLoginID + "', " + Environment.NewLine +
                                   "       [CaseOccurrence], [ERPCouseNo], [CaseClose], [IsExemption], [PaidAmount], [Penalty], [PenaltyRatio], [InsuAmount], [IDCardNo], " + Environment.NewLine +
                                   "       [Birthday], [Assumeday], [TelephoneNo], [Address], [PersonDamage], [CarDamage], [ReportDate], [CaseDate], [CaseTime], " + Environment.NewLine +
                                   "       [OutReportNo], [PoliceUnit], [PoliceName], [HasVideo], [NoVideoReason], [HasCaseData] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCase] " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //開始刪除
                    using (SqlDataSource dsDelTemp = new SqlDataSource())
                    {
                        dsDelTemp.ConnectionString = vConnStr;
                        //刪除 C 檔
                        dsDelTemp.DeleteCommand = "delete AnecdoteCaseC where CaseNo = @CaseNo ";
                        dsDelTemp.DeleteParameters.Clear();
                        dsDelTemp.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                        dsDelTemp.Delete();
                        //刪除 B 檔
                        dsDelTemp.DeleteCommand = "delete AnecdoteCaseB where CaseNo = @CaseNo ";
                        dsDelTemp.DeleteParameters.Clear();
                        dsDelTemp.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                        dsDelTemp.Delete();
                        //刪除 A 檔
                        dsDelTemp.DeleteCommand = "delete AnecdoteCase where CaseNo = @CaseNo ";
                        dsDelTemp.DeleteParameters.Clear();
                        dsDelTemp.DeleteParameters.Add(new Parameter("CaseNo", DbType.String, vCaseNo_Temp.Trim()));
                        dsDelTemp.Delete();
                    }
                    CaseDataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message.Trim();
                    eErrorMSG_A.Visible = true;
                }
            }
            else
            {
                eErrorMSG_A.Text = "請先選擇要刪除的肇事單";
                eErrorMSG_A.Visible = true;
            }
        }

        /// <summary>
        /// 確定主檔新增/異動
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbOK_Main_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vMode = Session["AnecdoteMode"].ToString().Trim().ToUpper();
            switch (vMode)
            {
                case "INS":
                    CreateAnecdoteA();
                    break;
                case "EDIT":
                    ModifyAnecdoteA();
                    break;
            }
            MainButtonStatus("LIST");
            SubButtonStatus("LIST");
            MainTextboxStatus("LIST");
            SubTextboxStatus("LIST");
            OpenData_A();
        }

        /// <summary>
        /// 新增主檔資料
        /// </summary>
        private void CreateAnecdoteA()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            //產生新單號
            string vCaseNo_INS;
            string vCaseNo_FirstCode = (DateTime.Today.Year - 1911).ToString() + DateTime.Today.Month.ToString("D2") + "A";
            string vSQLStr_INS = "select max(CaseNo) MaxNo from AnecdoteCase where CaseNo like '" + vCaseNo_FirstCode + "%' ";
            string vMaxNo = PF.GetValue(vConnStr, vSQLStr_INS, "MaxNo");
            string vIndexStr = (vMaxNo.Trim() != "") ? vMaxNo.Replace(vCaseNo_FirstCode, "").Trim() : "0";
            int vIndex = Int32.Parse(vIndexStr) + 1;
            vCaseNo_INS = vCaseNo_FirstCode + vIndex.ToString("D4");
            //寫入記錄
            string vRecordNote = "新增肇事單主檔資料_" + vCaseNo_INS + Environment.NewLine +
                                 "建檔日期：" + vToday.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                 "肇事單號：[" + vCaseNo_INS + "]" + Environment.NewLine +
                                 "建檔人工號：" + vLoginID + Environment.NewLine +
                                 "AnecdoteCase.aspx";
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            //開始寫入
            vSQLStr_INS = "insert into AnecdoteCase " + Environment.NewLine +
                          "       (CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, DriverName, InsuMan, " + Environment.NewLine +
                          "        AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, " + Environment.NewLine +
                          "        PaidAmount, Penalty, PenaltyRatio, InsuAmount, IDCardNo, Birthday, Assumeday, TelephoneNo, Address, " + Environment.NewLine +
                          "        PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, OutReportNo, PoliceUnit, PoliceName, HasVideo, " + Environment.NewLine +
                          "        NoVideoReason, HasCaseData) " + Environment.NewLine +
                          "values (@CaseNo, @HasInsurance, @DepNo, @DepName, @BuildDate, @BuildMan, @Car_ID, @Driver, @DriverName, @InsuMan, " + Environment.NewLine +
                          "        @AnecdotalResRatio, @IsNoDeduction, @DeductionDate, @Remark, @CaseOccurrence, @ERPCouseNo, @CaseClose, @IsExemption, " + Environment.NewLine +
                          "        @PaidAmount, @Penalty, @PenaltyRatio, @InsuAmount, @IDCardNo, @Birthday, @Assumeday, @TelephoneNo, @Address, " + Environment.NewLine +
                          "        @PersonDamage, @CarDamage, @ReportDate, @CaseDate, @CaseTime, @OutReportNo, @PoliceUnit, @PoliceName, @HasVideo, " + Environment.NewLine +
                          "        @NoVideoReason, @HasCaseData)";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlCommand cmdTemp = new SqlCommand(vSQLStr_INS, connTemp);
                connTemp.Open();
                cmdTemp.Parameters.Clear();
                cmdTemp.Parameters.Add(new SqlParameter("HasInsurance", (eHasInsurance.Text.Trim() != "") ? eHasInsurance.Text.Trim() : "False"));
                cmdTemp.Parameters.Add(new SqlParameter("DepNo", (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("DepName", (eDepName.Text.Trim() != "") ? eDepName.Text.Trim() : String.Empty));
                DateTime vTempDate;
                cmdTemp.Parameters.Add(new SqlParameter("BuildDate", DateTime.TryParse(eBuildDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("BuildMan", vLoginID));
                cmdTemp.Parameters.Add(new SqlParameter("Car_ID", (eCar_ID.Text.Trim() != "") ? eCar_ID.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("Driver", (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("DriverName", (eDriverName.Text.Trim() != "") ? eDriverName.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("InsuMan", (eInsuMan.Text.Trim() != "") ? eInsuMan.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("AnecdotalResRatio", (eAnecdotalResRatio.Text.Trim() != "") ? eAnecdotalResRatio.Text.Trim() : "0.0"));
                cmdTemp.Parameters.Add(new SqlParameter("IsNoDeduction", (eIsNoDeduction.Text.Trim() != "") ? eIsNoDeduction.Text.Trim() : "False"));
                cmdTemp.Parameters.Add(new SqlParameter("DeductionDate", DateTime.TryParse(eDeductionDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("Remark", (eRemark_A.Text.Trim() != "") ? eRemark_A.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("ModifyMan", vLoginID));
                cmdTemp.Parameters.Add(new SqlParameter("CaseOccurrence", (eCaseOccurrence.Text.Trim() != "") ? eCaseOccurrence.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("ERPCouseNo", (eERPCouseNo.Text.Trim() != "") ? eERPCouseNo.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("CaseClose", (eCaseClose.Text.Trim() != "") ? eCaseClose.Text.Trim() : "False"));
                cmdTemp.Parameters.Add(new SqlParameter("IsExemption", (eIsExemption.Text.Trim() != "") ? eIsExemption.Text.Trim() : "X"));
                cmdTemp.Parameters.Add(new SqlParameter("PaidAmount", (ePaidAmount.Text.Trim() != "") ? ePaidAmount.Text.Trim() : "0.0"));
                cmdTemp.Parameters.Add(new SqlParameter("Penalty", (ePenalty.Text.Trim() != "") ? ePenalty.Text.Trim() : "0.0"));
                cmdTemp.Parameters.Add(new SqlParameter("PenaltyRatio", (ePenaltyRatio.Text.Trim() != "") ? ePenaltyRatio.Text.Trim() : "0.0"));
                cmdTemp.Parameters.Add(new SqlParameter("InsuAmount", (eInsuAmount.Text.Trim() != "") ? eInsuAmount.Text.Trim() : "0.0"));
                cmdTemp.Parameters.Add(new SqlParameter("IDCardNo", (eIDCardNo.Text.Trim() != "") ? eIDCardNo.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("Birthday", DateTime.TryParse(eBirthday.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("Assumeday", DateTime.TryParse(eAssumeday.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("TelePhoneNo", (eTelephoneNo.Text.Trim() != "") ? eTelephoneNo.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("Address", (eAddress.Text.Trim() != "") ? eAddress.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("PersonDamage", (ePersonDamage.Text.Trim() != "") ? ePersonDamage.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("CarDamage", (eCarDamage.Text.Trim() != "") ? eCarDamage.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("ReportDate", DateTime.TryParse(eReportDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("CaseDate", DateTime.TryParse(eCaseDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("CaseTime", (eCaseTime.Text.Trim() != "") ? eCaseTime.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("OutReportNo", (eOutReportNo.Text.Trim() != "") ? eOutReportNo.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("PoliceUnit", (ePoliceUnit.Text.Trim() != "") ? ePoliceUnit.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("PoliceName", (ePoliceName.Text.Trim() != "") ? ePoliceName.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("HasVideo", (eHasVideo.Text.Trim() != "") ? eHasVideo.Text.Trim() : "X"));
                cmdTemp.Parameters.Add(new SqlParameter("NoVideoReason", (eNoVideoReason.Text.Trim() != "") ? eNoVideoReason.Text.Trim() : String.Empty));
                cmdTemp.Parameters.Add(new SqlParameter("HasCaseData", (eHasCaseData.Text.Trim() != "") ? eHasCaseData.Text.Trim() : "X"));
                cmdTemp.Parameters.Add(new SqlParameter("CaseNo", vCaseNo_INS));
                cmdTemp.ExecuteNonQuery();
            }
            CaseDataBind();
        }

        /// <summary>
        /// 異動主檔資料
        /// </summary>
        private void ModifyAnecdoteA()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            if (eCaseNo_A.Text.Trim() != "")
            {
                string vCaseNo_Temp = eCaseNo_A.Text.Trim();
                try
                {
                    string vRecordNote = "修改肇事單主檔資料_" + vCaseNo_Temp + Environment.NewLine +
                                         "修改日期：" + vToday.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                         "肇事單號：[" + vCaseNo_Temp + "]" + Environment.NewLine +
                                         "修改人工號：" + vLoginID + Environment.NewLine +
                                         "AnecdoteCase.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    string vHistoryNo = vToday.Year.ToString("D4") + vToday.Month.ToString("D2") + "EDITA";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndex_Str = vMaxNo.Replace(vHistoryNo, "").Trim();
                    int vIndex = (vIndex_Str != "") ? Int32.Parse(vIndex_Str) + 1 : 1;
                    string vNewHistoryNo = vHistoryNo + vIndex.ToString("D4");
                    //複製 A 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO [dbo].[AnecdoteCaseHistory] " + Environment.NewLine +
                                   "            ([HistoryNo], [CaseNo], [HasInsurance], [DepNo], [DepName], [BuildDate], [BuildMan], [Car_ID], [Driver], " + Environment.NewLine +
                                   "             [DriverName], [InsuMan], [AnecdotalResRatio], [IsNoDeduction], [DeductionDate], [Remark], [ModifyType], " + Environment.NewLine +
                                   "             [ModifyDate], [ModifyMan], [CaseOccurrence], [ERPCouseNo], [CaseClose], [IsExemption], [PaidAmount], " + Environment.NewLine +
                                   "             [Penalty], [PenaltyRatio], [InsuAmount], [IDCardNo], [Birthday], [Assumeday], [TelephoneNo], [Address], " + Environment.NewLine +
                                   "             [PersonDamage], [CarDamage], [ReportDate], [CaseDate], [CaseTime], [OutReportNo], [PoliceUnit], [PoliceName], " + Environment.NewLine +
                                   "             [HasVideo], [NoVideoReason], [HasCaseData]) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', [CaseNo], [HasInsurance], [DepNo], [DepName], [BuildDate], [BuildMan], [Car_ID], [Driver], " + Environment.NewLine +
                                   "       [DriverName], [InsuMan], [AnecdotalResRatio], [IsNoDeduction], [DeductionDate], [Remark], 'EDITA', GetDate(), '" + vLoginID + "', " + Environment.NewLine +
                                   "       [CaseOccurrence], [ERPCouseNo], [CaseClose], [IsExemption], [PaidAmount], [Penalty], [PenaltyRatio], [InsuAmount], [IDCardNo], " + Environment.NewLine +
                                   "       [Birthday], [Assumeday], [TelephoneNo], [Address], [PersonDamage], [CarDamage], [ReportDate], [CaseDate], [CaseTime], " + Environment.NewLine +
                                   "       [OutReportNo], [PoliceUnit], [PoliceName], [HasVideo], [NoVideoReason], [HasCaseData] " + Environment.NewLine +
                                   "  from [dbo].[AnecdoteCase] " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //開始更新
                    vSQLStr_Temp = "update [dbo].[AnecdoteCase] " + Environment.NewLine +
                                   "   set HasInsurance = @HasInsurance, DepNo = @DepNo, DepName = @DepName, BuildDate = @BuildDate, Car_ID = @Car_ID, " + Environment.NewLine +
                                   "       Driver = @Driver, DriverName = @DriverName, InsuMan = @InsuMan, AnecdotalResRatio = @AnecdotalResRatio, " + Environment.NewLine +
                                   "       IsNoDeduction = @IsNoDeduction, DeductionDate = @DeductionDate, Remark = @Remark, ModifyDate = GetDate(), " + Environment.NewLine +
                                   "       ModifyMan = @ModifyMan, CaseOccurrence = @CaseOccurrence, ERPCouseNo = @ERPCouseNo, CaseClose = @CaseClose, " + Environment.NewLine +
                                   "       IsExemption = @IsExemption, PaidAmount = @PaidAmount, Penalty = @Penalty, PenaltyRatio = @PenaltyRatio, " + Environment.NewLine +
                                   "       InsuAmount = @InsuAmount, IDCardNo = @IDCardNo, Birthday = @Birthday, Assumeday = @Assumeday, TelephoneNo = @TelephoneNo, " + Environment.NewLine +
                                   "       Address = @Address, PersonDamage = @PersonDamage, CarDamage = @CarDamage, ReportDate = @ReportDate, CaseDate = @CaseDate, " + Environment.NewLine +
                                   "       CaseTime = @CaseTime, OutReportNo = @OutReportNo, PoliceUnit = @PoliceUnit, PoliceName = @PoliceName, HasVideo = @HasVideo, " + Environment.NewLine +
                                   "       NoVideoReason = @NoVideoReason, HasCaseData = @HasCaseData, BuildMan = @BuildMan " + Environment.NewLine +
                                   " where CaseNo = @CaseNo";
                    using (SqlConnection connTemp = new SqlConnection(vConnStr))
                    {
                        SqlCommand cmdTemp = new SqlCommand(vSQLStr_Temp, connTemp);
                        cmdTemp.Parameters.Clear();
                        cmdTemp.Parameters.Add(new SqlParameter("HasInsurance", (eHasInsurance.Text.Trim() != "") ? eHasInsurance.Text.Trim() : "False"));
                        cmdTemp.Parameters.Add(new SqlParameter("DepNo", (eDepNo.Text.Trim() != "") ? eDepNo.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("DepName", (eDepName.Text.Trim() != "") ? eDepName.Text.Trim() : String.Empty));
                        DateTime vBuildDate;
                        cmdTemp.Parameters.Add(new SqlParameter("BuildDate", DateTime.TryParse(eBuildDate.Text.Trim(), out vBuildDate) ? vBuildDate.ToShortDateString() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("BuildMan", (eBuildMan.Text.Trim() != "") ? eBuildMan.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("Car_ID", (eCar_ID.Text.Trim() != "") ? eCar_ID.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("Driver", (eDriver.Text.Trim() != "") ? eDriver.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("DriverName", (eDriverName.Text.Trim() != "") ? eDriverName.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("InsuMan", (eInsuMan.Text.Trim() != "") ? eInsuMan.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("AnecdotalResRatio", (eAnecdotalResRatio.Text.Trim() != "") ? eAnecdotalResRatio.Text.Trim() : "0.0"));
                        cmdTemp.Parameters.Add(new SqlParameter("IsNoDeduction", (eIsNoDeduction.Text.Trim() != "") ? eIsNoDeduction.Text.Trim() : "False"));
                        DateTime vDeductionDate;
                        cmdTemp.Parameters.Add(new SqlParameter("DeductionDate", DateTime.TryParse(eDeductionDate.Text.Trim(), out vDeductionDate) ? vDeductionDate.ToShortDateString() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("Remark", (eRemark_A.Text.Trim() != "") ? eRemark_A.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("ModifyMan", vLoginID));
                        cmdTemp.Parameters.Add(new SqlParameter("CaseOccurrence", (eCaseOccurrence.Text.Trim() != "") ? eCaseOccurrence.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("ERPCouseNo", (eERPCouseNo.Text.Trim() != "") ? eERPCouseNo.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("CaseClose", (eCaseClose.Text.Trim() != "") ? eCaseClose.Text.Trim() : "False"));
                        cmdTemp.Parameters.Add(new SqlParameter("IsExemption", (eIsExemption.Text.Trim() != "") ? eIsExemption.Text.Trim() : "X"));
                        cmdTemp.Parameters.Add(new SqlParameter("PaidAmount", (ePaidAmount.Text.Trim() != "") ? ePaidAmount.Text.Trim() : "0.0"));
                        cmdTemp.Parameters.Add(new SqlParameter("Penalty", (ePenalty.Text.Trim() != "") ? ePenalty.Text.Trim() : "0.0"));
                        cmdTemp.Parameters.Add(new SqlParameter("PenaltyRatio", (ePenaltyRatio.Text.Trim() != "") ? ePenaltyRatio.Text.Trim() : "0.0"));
                        cmdTemp.Parameters.Add(new SqlParameter("InsuAmount", (eInsuAmount.Text.Trim() != "") ? eInsuAmount.Text.Trim() : "0.0"));
                        cmdTemp.Parameters.Add(new SqlParameter("IDCardNo", (eIDCardNo.Text.Trim() != "") ? eIDCardNo.Text.Trim() : String.Empty));
                        DateTime vBirthday;
                        cmdTemp.Parameters.Add(new SqlParameter("Birthday", DateTime.TryParse(eBirthday.Text.Trim(), out vBirthday) ? vBirthday.ToShortDateString() : String.Empty));
                        DateTime vAssumeday;
                        cmdTemp.Parameters.Add(new SqlParameter("Assumeday", DateTime.TryParse(eAssumeday.Text.Trim(), out vAssumeday) ? vAssumeday.ToShortDateString() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("TelePhoneNo", (eTelephoneNo.Text.Trim() != "") ? eTelephoneNo.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("Address", (eAddress.Text.Trim() != "") ? eAddress.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("PersonDamage", (ePersonDamage.Text.Trim() != "") ? ePersonDamage.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("CarDamage", (eCarDamage.Text.Trim() != "") ? eCarDamage.Text.Trim() : String.Empty));
                        DateTime vReportDate;
                        cmdTemp.Parameters.Add(new SqlParameter("ReportDate", DateTime.TryParse(eReportDate.Text.Trim(), out vReportDate) ? vReportDate.ToShortDateString() : String.Empty));
                        DateTime vCaseDate;
                        cmdTemp.Parameters.Add(new SqlParameter("CaseDate", DateTime.TryParse(eCaseDate.Text.Trim(), out vCaseDate) ? vCaseDate.ToShortDateString() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("CaseTime", (eCaseTime.Text.Trim() != "") ? eCaseTime.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("OutReportNo", (eOutReportNo.Text.Trim() != "") ? eOutReportNo.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("PoliceUnit", (ePoliceUnit.Text.Trim() != "") ? ePoliceUnit.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("PoliceName", (ePoliceName.Text.Trim() != "") ? ePoliceName.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("HasVideo", (eHasVideo.Text.Trim() != "") ? eHasVideo.Text.Trim() : "X"));
                        cmdTemp.Parameters.Add(new SqlParameter("NoVideoReason", (eNoVideoReason.Text.Trim() != "") ? eNoVideoReason.Text.Trim() : String.Empty));
                        cmdTemp.Parameters.Add(new SqlParameter("HasCaseData", (eHasCaseData.Text.Trim() != "") ? eHasCaseData.Text.Trim() : "X"));
                        cmdTemp.Parameters.Add(new SqlParameter("CaseNo", vCaseNo_Temp));
                        connTemp.Open();
                        cmdTemp.ExecuteNonQuery();
                    }
                    CaseDataBind();
                }
                catch (Exception eMessage)
                {
                    eErrorMSG_A.Text = eMessage.Message.Trim();
                    eErrorMSG_A.Visible = true;
                }
            }
            else
            {
                eErrorMSG_A.Text = "請先選擇要編輯的肇事單";
                eErrorMSG_A.Visible = true;
            }
        }

        /// <summary>
        /// 取消主檔新增/異動
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCancel_Main_Click(object sender, EventArgs e)
        {
            CaseDataBind();
            MainButtonStatus("LIST");
            SubButtonStatus("LIST");
            MainTextboxStatus("LIST");
            SubTextboxStatus("LIST");
        }

        protected void bbPrintReport_Click(object sender, EventArgs e)
        {
            plPrint.Visible = true;
            plMainDataShow.Visible = false;
            plDetailDataShow.Visible = false;
        }

        protected void bbPrintNote_Click(object sender, EventArgs e)
        {
            plPrint.Visible = true;
            plMainDataShow.Visible = false;
            plDetailDataShow.Visible = false;
        }

        /// <summary>
        /// 與 ERP 同步
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbERPSync_Click(object sender, EventArgs e)
        {
            eErrorMSG_A.Text = "";
            eErrorMSG_A.Visible = false;
            string vSQLStr_Temp = "";
            string vCaseNo = eCaseNo_A.Text.Trim();
            string vBuileDate_TempStr = eBuildDate.Text.Trim();
            string[] vaBuildDate_Temp = vBuileDate_TempStr.Split('/');
            int vYear_Temp = ((vaBuildDate_Temp != null) && (vaBuildDate_Temp[0].Trim() != "")) ? Int32.Parse(vaBuildDate_Temp[0].Trim()) : 0;
            int vMonth_Temp = ((vaBuildDate_Temp != null) && (vaBuildDate_Temp[1].Trim() != "")) ? Int32.Parse(vaBuildDate_Temp[1].Trim()) : 0;
            int vDay_Temp = ((vaBuildDate_Temp != null) && (vaBuildDate_Temp[2].Trim() != "")) ? Int32.Parse(vaBuildDate_Temp[2].Trim()) : 0;
            string vCouseDate_New = (vYear_Temp + vMonth_Temp + vDay_Temp == 0) ? "" :
                                    (vYear_Temp < 1911) ? (vYear_Temp + 1911).ToString("D4") + "/" + vMonth_Temp.ToString("D2") + "/" + vDay_Temp.ToString("D2") :
                                     vYear_Temp.ToString("D4") + "/" + vMonth_Temp.ToString("D2") + "/" + vDay_Temp.ToString("D2");
            string vCaseYM = (DateTime.Today.Year - 1911).ToString("D3") + DateTime.Today.Month.ToString("D2");
            vSQLStr_Temp = "select MAX(CouseNo) MaxNo from Car_CouseTrouble where CouseNo like '" + vCaseYM + "%' ";
            string vMaxIndex = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
            int vIndex = (vMaxIndex != "") ? Int32.Parse(vMaxIndex.Replace(vCaseYM, "")) + 1 : 1;
            string vCouseNo_New = vCaseYM + vIndex.ToString("D4");
            string vDepNo_New = eDepNo.Text.Trim();
            string vCarID_New = eCar_ID.Text.Trim();
            vSQLStr_Temp = "select (isnull(Car_ID, '') + ',' + isnull(Car_Class, '') + ',' + " + Environment.NewLine +
                           "        isnull(convert(varchar(10),ProdDate, 111), '') + ',' + " + Environment.NewLine +
                           "        isnull(BCENo, '') + ',' + isnull(Exhaust, '')) CarData " + Environment.NewLine +
                           "  from Car_InfoA " + Environment.NewLine +
                           " where Car_ID = '" + vCarID_New + "' ";
            string vCarData = PF.GetValue(vConnStr, vSQLStr_Temp, "CarData");
            string[] vaCarData_New = vCarData.Split(',');
            string vCarClass_New = (vaCarData_New != null) ? vaCarData_New[1].ToString().Trim() : "";
            string vProdDate_New = (vaCarData_New != null) ? vaCarData_New[2].ToString().Trim() : "";
            string vBCENo_New = (vaCarData_New != null) ? vaCarData_New[3].ToString().Trim() : "";
            string vExhaust_New = (vaCarData_New != null) ? vaCarData_New[4].ToString().Trim() : "";
            string vEsempNo_New = eBuildMan.Text.Trim();
            string vDriver_New = eDriver.Text.Trim();
            string vMemo_New = eRemark_A.Text.Trim();
            string vCouseCondition_New = eCaseOccurrence.Text.Trim();
            vSQLStr_Temp = "insert into Car_CouseTrouble(CouseNo, DepNo, CouseDate, Car_ID, Car_Class, ProdDate, Displacement, EngineNo, EsempNo, Driver, Memo, Course_Condition) " + Environment.NewLine +
                           "values(@CouseNo, @DepNo, @CouseDate, @Car_ID, @Car_Class, @ProdDate, @Displacement, @EngineNo, @EsempNo, @Driver, @Memo, @Course_Condition)";
            try
            {
                using (SqlConnection connInsert = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdInsert = new SqlCommand(vSQLStr_Temp, connInsert);
                    cmdInsert.Parameters.Clear();
                    cmdInsert.Parameters.Add(new SqlParameter("CouseNo", vCouseNo_New));
                    cmdInsert.Parameters.Add(new SqlParameter("DepNo", vDepNo_New));
                    cmdInsert.Parameters.Add(new SqlParameter("CouseDate", vCouseDate_New));
                    cmdInsert.Parameters.Add(new SqlParameter("Car_ID", vCarID_New));
                    cmdInsert.Parameters.Add(new SqlParameter("Car_Class", vCarClass_New));
                    cmdInsert.Parameters.Add(new SqlParameter("ProdDate", vProdDate_New));
                    cmdInsert.Parameters.Add(new SqlParameter("Displacement", vExhaust_New));
                    cmdInsert.Parameters.Add(new SqlParameter("EngineNo", vBCENo_New));
                    cmdInsert.Parameters.Add(new SqlParameter("EsempNo", vEsempNo_New));
                    cmdInsert.Parameters.Add(new SqlParameter("Driver", vDriver_New));
                    cmdInsert.Parameters.Add(new SqlParameter("Memo", vMemo_New));
                    cmdInsert.Parameters.Add(new SqlParameter("Course_Condition", vCouseCondition_New));
                    connInsert.Open();
                    cmdInsert.ExecuteNonQuery();
                }
                vSQLStr_Temp = "update AnecdoteCase set ERPCouseNo = @ERPCouseNo where CaseNo = @CaseNo";
                using (SqlConnection connUpdate = new SqlConnection(vConnStr))
                {
                    SqlCommand cmdUpdate = new SqlCommand(vSQLStr_Temp, connUpdate);
                    cmdUpdate.Parameters.Clear();
                    cmdUpdate.Parameters.Add(new SqlParameter("ERPCouseNo", vCouseNo_New));
                    cmdUpdate.Parameters.Add(new SqlParameter("CaseNo", vCaseNo));
                    connUpdate.Open();
                    cmdUpdate.ExecuteNonQuery();
                }
                gridAnecdoteCaseA_List.DataBind();
                string vRecordNote = "與ERP同步" + Environment.NewLine +
                                     "AnecdoteCase.aspx" + Environment.NewLine +
                                     "線上系統單號：" + vCaseNo + Environment.NewLine +
                                     "ERP單號：" + vCouseNo_New;
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            }
            catch (Exception eMessage)
            {
                eErrorMSG_A.Text = eMessage.Message;
                eErrorMSG_A.Visible = true;
            }
        }

        /// <summary>
        /// 計算罰款分擔
        /// </summary>
        /// <param name="fInsuAmount"></param>
        /// <param name="fPenaltyRatio"></param>
        /// <returns></returns>
        private double CalPenaltyAmount(double fInsuAmount, double fPenaltyRatio)
        {
            double vPenalty = 0.0;
            double vBaseAmount = fInsuAmount / 4.0;
            double vBaseAmount2 = (fInsuAmount - 200000) / 4.0;
            double vCalRatio_Low = 0.0;
            double vCalRatio_High = 0.0;

            if (fPenaltyRatio <= 20.0)
            {
                vCalRatio_Low = 0.1;
                vCalRatio_High = 0.05;
            }
            else if ((fPenaltyRatio > 20.0) && (fPenaltyRatio <= 30.0))
            {
                vCalRatio_Low = 0.1;
                vCalRatio_High = 0.05;
            }
            else if ((fPenaltyRatio > 30.0) && (fPenaltyRatio <= 40.0))
            {
                vCalRatio_Low = 0.15;
                vCalRatio_High = 0.1;
            }
            else if ((fPenaltyRatio > 40.0) && (fPenaltyRatio <= 50.0))
            {
                vCalRatio_Low = 0.15;
                vCalRatio_High = 0.1;
            }
            else if ((fPenaltyRatio > 50.0) && (fPenaltyRatio <= 60.0))
            {
                vCalRatio_Low = 0.2;
                vCalRatio_High = 0.1;
            }
            else if ((fPenaltyRatio > 60.0) && (fPenaltyRatio <= 70.0))
            {
                vCalRatio_Low = 0.25;
                vCalRatio_High = 0.15;
            }
            else if ((fPenaltyRatio > 70.0) && (fPenaltyRatio <= 80.0))
            {
                vCalRatio_Low = 0.25;
                vCalRatio_High = 0.15;
            }
            else if (fPenaltyRatio > 80.0)
            {
                vCalRatio_Low = 0.25;
                vCalRatio_High = 0.15;
            }
            vPenalty = (fInsuAmount <= 200000) ?
                       (double)Math.Round(vBaseAmount * vCalRatio_Low, 0, MidpointRounding.AwayFromZero) :
                       (double)Math.Round(((50000 * vCalRatio_Low) + (vBaseAmount2 * vCalRatio_High)), 0, MidpointRounding.AwayFromZero);
            return vPenalty;
        }

        protected void ePaidAmount_TextChanged(object sender, EventArgs e)
        {
            double vPenaltyRatio;
            double vInsuAmount;
            if ((double.TryParse(eInsuAmount.Text.Trim(), out vInsuAmount)) && (double.TryParse(ePenaltyRatio.Text.Trim(), out vPenaltyRatio)))
            {
                ePenalty.Text = CalPenaltyAmount(vInsuAmount, vPenaltyRatio).ToString();
            }
        }

        protected void bbNew_B_Click(object sender, EventArgs e)
        {
            Session["AnecdoteDetailMode"] = "INS";
            SubButtonStatus("INS");
            SubTextboxStatus("INS");
            plDetailData.Visible = true;
        }

        protected void bbEdit_B_Click(object sender, EventArgs e)
        {
            Session["AnecdoteDetailMode"] = "EDIT";
            SubButtonStatus("EDIT");
            SubTextboxStatus("EDIT");
            //plDetailDataShow.Visible = true;
        }

        protected void bbDel_B_Click(object sender, EventArgs e)
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            try
            {
                string vCaseNoItems = eCaseNoItems.Text.Trim();
                //如果肇事明細編號是空值就不進行任何動作
                if (!String.IsNullOrEmpty(vCaseNoItems))
                {
                    if (vConnStr == "")
                    {
                        vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
                    }
                    //雖然是刪除 B 檔，還是要在 A 檔的異動記錄加入一筆資料
                    DateTime vModifyDate = DateTime.Today;
                    string vHistoryNo = vModifyDate.Year.ToString("D4") + vModifyDate.Month.ToString("D2") + "DELB";
                    string vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndex_Str = (vMaxNo.Replace(vHistoryNo, "").Trim() != "") ? vMaxNo.Replace(vHistoryNo, "").Trim() : "0";
                    int vIndex = Int32.Parse(vIndex_Str) + 1;
                    string vNewHistoryNo = vHistoryNo + vIndex.ToString("D4");
                    string vCaseNo_Temp = eCaseNo_B.Text.Trim();
                    //寫入 A 檔異動記錄
                    vSQLStr_Temp = "INSERT INTO AnecdoteCaseHistory " + Environment.NewLine +
                                   "            (HistoryNo, CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, " + Environment.NewLine +
                                   "             DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, ModifyType, " + Environment.NewLine +
                                   "             ModifyDate, ModifyMan, CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, PaidAmount, " + Environment.NewLine +
                                   "             Penalty, PenaltyRatio, InsuAmount, IDCardNo, Birthday, Assumeday, TelephoneNo, Address, " + Environment.NewLine +
                                   "             PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, OutReportNo, PoliceUnit, PoliceName, " + Environment.NewLine +
                                   "             HasVideo, NoVideoReason, HasCaseData) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, " + Environment.NewLine +
                                   "       DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, 'EDITA', GetDate(), '" + vLoginID + "', " + Environment.NewLine +
                                   "       CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, PaidAmount, Penalty, PenaltyRatio, InsuAmount, IDCardNo, " + Environment.NewLine +
                                   "       Birthday, Assumeday, TelephoneNo, Address, PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, " + Environment.NewLine +
                                   "       OutReportNo, PoliceUnit, PoliceName, HasVideo, NoVideoReason, HasCaseData " + Environment.NewLine +
                                   "  from AnecdoteCase " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_Temp + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入 B 檔異動記錄
                    vSQLStr_Temp = "select max(Items) MaxItem from AnecdoteCaseBHistory where HistoryNo = '" + vNewHistoryNo + "' ";
                    vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    int vItems = (vMaxNo != "") ? Int32.Parse(vMaxNo.Trim()) + 1 : 1;
                    //複製 B 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO AnecdoteCaseBHistory " + Environment.NewLine +
                                   "            (HistoryNo, ItemsH, HistoryNoItems, CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, " + Environment.NewLine +
                                   "             EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, " + Environment.NewLine +
                                   "             PersonDamageAMT, RelationComp, ReconciliationDate, Remark, ModifyType, ModifyDate, ModifyMan, " + Environment.NewLine +
                                   "             PassengerInsu, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, RelNote)" + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', '" + vItems.ToString("D4") + "', '" + vNewHistoryNo + vItems.ToString("D4") + "', "+Environment.NewLine+
                                   "       CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, "+Environment.NewLine+
                                   "       CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, Remark, 'DELB', GetDate(),  " + Environment.NewLine +
                                   "       '" + vLoginID + "', PassengerInsu, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, RelNote " + Environment.NewLine +
                                   "  from AnecdoteCaseB " + Environment.NewLine +
                                   " where CaseNoItems = '" + vCaseNoItems + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);

                    //開始刪除
                    using (SqlConnection connDel_B = new SqlConnection(vConnStr))
                    {
                        string vDelStr_B = "delete from AnecdoteCaseB where CaseNoItems = @CaseNoItems";
                        SqlCommand cmdDel_B = new SqlCommand(vDelStr_B, connDel_B);
                        connDel_B.Open();
                        cmdDel_B.Parameters.Clear();
                        cmdDel_B.Parameters.Add(new SqlParameter("CaseNoItems", vCaseNoItems));
                        cmdDel_B.ExecuteNonQuery();
                    }
                    OpenData_B();
                }
                else
                {
                    eErrorMSG_B.Text = "請選擇要刪除的明細";
                    eErrorMSG_B.Visible = true;
                }
            }
            catch (Exception eMessage)
            {
                eErrorMSG_B.Text = eMessage.Message;
                eErrorMSG_B.Visible = true;
            }
        }

        protected void bbOK_B_Click(object sender, EventArgs e)
        {
            switch (Session["AnecdoteDetailMode"].ToString().Trim())
            {
                case "INS":
                    CreateAnecdoteB();
                    break;

                case "EDIT":
                    ModifyAnecdoteB();
                    break;
            }
            MainButtonStatus("LIST");
            SubButtonStatus("LIST");
            MainTextboxStatus("LIST");
            SubTextboxStatus("LIST");
            OpenData_B();
        }

        /// <summary>
        /// 新增 B 檔資料
        /// </summary>
        private void CreateAnecdoteB()
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCaseNo_TempB = eCaseNo_B.Text.Trim();
            if (!string.IsNullOrEmpty(vCaseNo_TempB))
            {
                string vTempSQL = "select Max(Items) MaxItems from AnecdoteCaseB where CaseNo = '" + vCaseNo_TempB + "' ";
                string vMaxItems = PF.GetValue(vConnStr, vTempSQL, "MaxItems");
                int vTempINT;
                string vNewItems = Int32.TryParse(vMaxItems, out vTempINT) ? (vTempINT + 1).ToString("D4") : "0001";
                string vCaseNoItems = vCaseNo_TempB + vNewItems;
                //寫入記錄
                string vRecordNote = "新增肇事單明細資料_" + vCaseNo_TempB + Environment.NewLine +
                                     "建檔日期：" + vToday.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                     "肇事單號：[" + vCaseNo_TempB + "]" + Environment.NewLine +
                                     "明細項次：[" + vNewItems + "]" + Environment.NewLine +
                                     "建檔人工號：" + vLoginID + Environment.NewLine +
                                     "AnecdoteCase.aspx";
                PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

                vTempSQL = "insert into AnecdoteCaseB " + Environment.NewLine +
                           "       (CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, " + Environment.NewLine +
                           "        CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, PassengerInsu, Remark, RelGender, RelTelNo1, RelTelNo2, RelCarType, " + Environment.NewLine +
                           "        RelPersonDamage, RelCarDamage, RelNote)" + Environment.NewLine +
                           "values (@CaseNo, @Items, @CaseNoItems, @Relationship, @RelCar_ID, @EstimatedAmount, @ThirdInsurance, @CompInsurance, @DriverSharing, @CompanySharing, " + Environment.NewLine +
                           "        @CarDamageAMT, @PersonDamageAMT, @RelationComp, @ReconciliationDate, @PassengerInsu, @Remark, @RelGender, @RelTelNo1, @RelTelNo2, @RelCarType, " + Environment.NewLine +
                           "        @RelPersonDamage, @RelCarDamage, @RelNote)";
                using (SqlConnection connINS_B = new SqlConnection(vConnStr))
                {
                    double vTempDouble;
                    DateTime vTempDate;
                    SqlCommand cmdINS_B = new SqlCommand(vTempSQL, connINS_B);
                    try
                    {
                        connINS_B.Open();
                        cmdINS_B.Parameters.Clear();
                        cmdINS_B.Parameters.Add(new SqlParameter("CaseNo", vCaseNo_TempB));
                        cmdINS_B.Parameters.Add(new SqlParameter("Items", vNewItems));
                        cmdINS_B.Parameters.Add(new SqlParameter("CaseNoItems", vCaseNoItems));
                        cmdINS_B.Parameters.Add(new SqlParameter("Relationship", eRelationComp.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("RelCar_ID", eRelCar_ID.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("EstimatedAmount", double.TryParse(eEstimatedAmount.Text.Trim(), out vTempDouble) ? vTempDouble.ToString() : "0.0"));
                        cmdINS_B.Parameters.Add(new SqlParameter("ThirdInsueance", double.TryParse(eThirdInsurance.Text.Trim(), out vTempDouble) ? vTempDouble.ToString() : "0.0"));
                        cmdINS_B.Parameters.Add(new SqlParameter("CompInsurance", double.TryParse(eCompInsurance.Text.Trim(), out vTempDouble) ? vTempDouble.ToString() : "0.0"));
                        cmdINS_B.Parameters.Add(new SqlParameter("DriverSharing", double.TryParse(eDriverSharing.Text.Trim(), out vTempDouble) ? vTempDouble.ToString() : "0.0"));
                        cmdINS_B.Parameters.Add(new SqlParameter("CompanySharing", double.TryParse(eCompanySharing.Text.Trim(), out vTempDouble) ? vTempDouble.ToString() : "0.0"));
                        cmdINS_B.Parameters.Add(new SqlParameter("CarDamageAMT", double.TryParse(eCarDamageAMT.Text.Trim(), out vTempDouble) ? vTempDouble.ToString() : "0.0"));
                        cmdINS_B.Parameters.Add(new SqlParameter("PersonDamageAMT", double.TryParse(ePersonDamageAMT.Text.Trim(), out vTempDouble) ? vTempDouble.ToString() : "0.0"));
                        cmdINS_B.Parameters.Add(new SqlParameter("RelationComp", eRelationComp.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("ReconciliationDate", DateTime.TryParse(eReconciliationDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                        cmdINS_B.Parameters.Add(new SqlParameter("PassengerInsu", double.TryParse(ePassengerInsu.Text.Trim(), out vTempDouble) ? vTempDouble.ToString() : "0.0"));
                        cmdINS_B.Parameters.Add(new SqlParameter("Remark", eRemark_B.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("RelGender", eRelGender.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("RelTelNo1", eRelTelNo1.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("RelTelNo2", eRelTelNo2.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("RelCarType", eRelCarType.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("RelPersonDamage", eRelPersonDamage.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("RelCarDamage", eRelCarDamage.Text.Trim()));
                        cmdINS_B.Parameters.Add(new SqlParameter("RelNote", eRelNote.Text.Trim()));
                        cmdINS_B.ExecuteNonQuery();
                    }
                    catch (Exception eMessage)
                    {
                        eErrorMSG_B.Text = eMessage.Message.Trim();
                        eErrorMSG_B.Visible = true;
                    }
                }
            }
        }

        /// <summary>
        /// 異動 B 檔資料
        /// </summary>
        private void ModifyAnecdoteB()
        {
            eErrorMSG_B.Text = "";
            eErrorMSG_B.Visible = false;
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vCaseNo_TempB = eCaseNo_B.Text.Trim();
            string vItems_TempB = eItems.Text.Trim();
            string vCaseNoItems_TempB = eCaseNoItems.Text.Trim();
            string vSQLStr_Temp = "";
            if (!string.IsNullOrEmpty(vCaseNoItems_TempB))
            {
                try
                {
                    //寫入記錄
                    string vRecordNote = "修改肇事單明細資料_" + vCaseNo_TempB + Environment.NewLine +
                                         "建檔日期：" + vToday.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine +
                                         "肇事單號：[" + vCaseNo_TempB + "]" + Environment.NewLine +
                                         "明細項次：[" + vItems_TempB + "]" + Environment.NewLine +
                                         "異動人工號：" + vLoginID + Environment.NewLine +
                                         "AnecdoteCase.aspx";
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);

                    //雖然是刪除 B 檔，還是要在 A 檔的異動記錄加入一筆資料
                    DateTime vModifyDate = DateTime.Today;
                    string vHistoryNo = vModifyDate.Year.ToString("D4") + vModifyDate.Month.ToString("D2") + "MODB";
                    vSQLStr_Temp = "select max(HistoryNo) MaxNo from AnecdoteCaseHistory where HistoryNo like '" + vHistoryNo + "%' ";
                    string vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxNo");
                    string vIndex_Str = (vMaxNo.Replace(vHistoryNo, "").Trim() != "") ? vMaxNo.Replace(vHistoryNo, "").Trim() : "0";
                    int vIndex = Int32.Parse(vIndex_Str) + 1;
                    string vNewHistoryNo = vHistoryNo + vIndex.ToString("D4");
                    //寫入 A 檔異動記錄
                    vSQLStr_Temp = "INSERT INTO AnecdoteCaseHistory " + Environment.NewLine +
                                   "            (HistoryNo, CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, " + Environment.NewLine +
                                   "             DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, ModifyType, " + Environment.NewLine +
                                   "             ModifyDate, ModifyMan, CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, PaidAmount, " + Environment.NewLine +
                                   "             Penalty, PenaltyRatio, InsuAmount, IDCardNo, Birthday, Assumeday, TelephoneNo, Address, " + Environment.NewLine +
                                   "             PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, OutReportNo, PoliceUnit, PoliceName, " + Environment.NewLine +
                                   "             HasVideo, NoVideoReason, HasCaseData) " + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', CaseNo, HasInsurance, DepNo, DepName, BuildDate, BuildMan, Car_ID, Driver, " + Environment.NewLine +
                                   "       DriverName, InsuMan, AnecdotalResRatio, IsNoDeduction, DeductionDate, Remark, 'EDITA', GetDate(), '" + vLoginID + "', " + Environment.NewLine +
                                   "       CaseOccurrence, ERPCouseNo, CaseClose, IsExemption, PaidAmount, Penalty, PenaltyRatio, InsuAmount, IDCardNo, " + Environment.NewLine +
                                   "       Birthday, Assumeday, TelephoneNo, Address, PersonDamage, CarDamage, ReportDate, CaseDate, CaseTime, " + Environment.NewLine +
                                   "       OutReportNo, PoliceUnit, PoliceName, HasVideo, NoVideoReason, HasCaseData " + Environment.NewLine +
                                   "  from AnecdoteCase " + Environment.NewLine +
                                   " where CaseNo = '" + vCaseNo_TempB + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);
                    //寫入 B 檔異動記錄
                    vSQLStr_Temp = "select max(Items) MaxItem from AnecdoteCaseBHistory where HistoryNo = '" + vNewHistoryNo + "' ";
                    vMaxNo = PF.GetValue(vConnStr, vSQLStr_Temp, "MaxItem");
                    int vItems = (vMaxNo != "") ? Int32.Parse(vMaxNo.Trim()) + 1 : 1;
                    //複製 B 檔原資料到異動檔
                    vSQLStr_Temp = "INSERT INTO AnecdoteCaseBHistory " + Environment.NewLine +
                                   "            (HistoryNo, ItemsH, HistoryNoItems, CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, " + Environment.NewLine +
                                   "             EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, CompanySharing, CarDamageAMT, " + Environment.NewLine +
                                   "             PersonDamageAMT, RelationComp, ReconciliationDate, Remark, ModifyType, ModifyDate, ModifyMan, " + Environment.NewLine +
                                   "             PassengerInsu, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, RelNote)" + Environment.NewLine +
                                   "select '" + vNewHistoryNo + "', '" + vItems.ToString("D4") + "', '" + vNewHistoryNo + vItems.ToString("D4") + "', " + Environment.NewLine +
                                   "       CaseNo, Items, CaseNoItems, Relationship, RelCar_ID, EstimatedAmount, ThirdInsurance, CompInsurance, DriverSharing, " + Environment.NewLine +
                                   "       CompanySharing, CarDamageAMT, PersonDamageAMT, RelationComp, ReconciliationDate, Remark, 'DELB', GetDate(),  " + Environment.NewLine +
                                   "       '" + vLoginID + "', PassengerInsu, RelGender, RelTelNo1, RelTelNo2, RelCarType, RelPersonDamage, RelCarDamage, RelNote " + Environment.NewLine +
                                   "  from AnecdoteCaseB " + Environment.NewLine +
                                   " where CaseNoItems = '" + vCaseNoItems_TempB + "' ";
                    PF.ExecSQL(vConnStr, vSQLStr_Temp);

                    //寫入更新資料
                    using (SqlConnection connModifyB=new SqlConnection(vConnStr))
                    {
                        vSQLStr_Temp = "update AnecdoteCaseB " + Environment.NewLine +
                                       "   set Relationship = @Relationship, RelCar_ID = @RelCar_ID, EstimatedAmount = @EstimatedAmount, ThirdInsurance = @ThirdInsurance, " + Environment.NewLine +
                                       "       CompInsurance = @CompInsurance, DriverSharing = @DriverSharing, CompanySharing = @CompanySharing, CarDamageAMT = @CarDamageAMT, " + Environment.NewLine +
                                       "       PersonDamageAMT = @PersonDamageAMT, RelationComp = @RelationComp, ReconciliationDate = @ReconciliationDate, Remark = @Remark, " + Environment.NewLine +
                                       "       ModifyDate = GetDate(), ModifyMan = '" + vLoginID + "', PassengerInsu = @PassengerInsu, RelGender = @RelGender, " + Environment.NewLine +
                                       "       RelTelNo1 = @RelTelNo1, RelTelNo2 = @RelTelNo2, RelCarType = @RelCarType, RelPersonDamage = @RelPersonDamage, " + Environment.NewLine +
                                       "       RelCarDamage = @RelCarDamage, RelNote = @RelNote " + Environment.NewLine +
                                       " where CaseNoItems = @CaseNoItems ";
                        SqlCommand cmdModifyB = new SqlCommand(vSQLStr_Temp, connModifyB);
                        connModifyB.Open();
                        int vTempINT;
                        double vTempFloat;
                        DateTime vTempDate;
                        cmdModifyB.Parameters.Clear();
                        cmdModifyB.Parameters.Add(new SqlParameter("Relationship", (eRelationship.Text.Trim() != "") ? eRelationship.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("RelCar_ID", (eRelCar_ID.Text.Trim() != "") ? eRelCar_ID.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("EstimatedAmount", Int32.TryParse(eEstimatedAmount.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                        cmdModifyB.Parameters.Add(new SqlParameter("ThirdInsurance", Int32.TryParse(eThirdInsurance.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                        cmdModifyB.Parameters.Add(new SqlParameter("CompInsurance", Int32.TryParse(eCompInsurance.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                        cmdModifyB.Parameters.Add(new SqlParameter("DriverSharing", Int32.TryParse(eDriverSharing.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                        cmdModifyB.Parameters.Add(new SqlParameter("CompanySharing", Int32.TryParse(eCompanySharing.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                        cmdModifyB.Parameters.Add(new SqlParameter("CarDamageAMT", Int32.TryParse(eCarDamageAMT.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                        cmdModifyB.Parameters.Add(new SqlParameter("PersonDamageAMT", Int32.TryParse(ePersonDamageAMT.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                        cmdModifyB.Parameters.Add(new SqlParameter("RelationComp", Int32.TryParse(eRelationComp.Text.Trim(), out vTempINT) ? vTempINT.ToString() : "0"));
                        cmdModifyB.Parameters.Add(new SqlParameter("ReconciliationDate", DateTime.TryParse(eReconciliationDate.Text.Trim(), out vTempDate) ? vTempDate.ToShortDateString() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("Remark", (eRemark_B.Text.Trim() != "") ? eRemark_B.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("PassengerInsu", double.TryParse(ePassengerInsu.Text.Trim(), out vTempFloat) ? vTempFloat.ToString() : "0.0"));
                        cmdModifyB.Parameters.Add(new SqlParameter("RelGender", (eRelGender.Text.Trim() != "") ? eRelGender.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("RelTelNo1", (eRelTelNo1.Text.Trim() != "") ? eRelTelNo1.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("RelTelNo2", (eRelTelNo2.Text.Trim() != "") ? eRelTelNo2.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("RelCarType", (eRelCarType.Text.Trim() != "") ? eRelCarType.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("RelPersonDamage", (eRelPersonDamage.Text.Trim() != "") ? eRelPersonDamage.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("RelCarDamage", (eRelCarDamage.Text.Trim() != "") ? eRelCarDamage.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("RelNote", (eRelNote.Text.Trim() != "") ? eRelNote.Text.Trim() : String.Empty));
                        cmdModifyB.Parameters.Add(new SqlParameter("CaseNoItems", vCaseNoItems_TempB));
                        cmdModifyB.ExecuteNonQuery();
                    }
                }
                catch(Exception eMessage)
                {
                    eErrorMSG_B.Text = eMessage.Message;
                    eErrorMSG_B.Visible = true;
                }
            }
        }

        protected void bbCancel_B_Click(object sender, EventArgs e)
        {
            MainButtonStatus("LIST");
            SubButtonStatus("LIST");
            MainTextboxStatus("LIST");
            SubTextboxStatus("LIST");
            OpenData_B();
        }

        protected void gridAnecdoteCaseB_List_DataBound(object sender, EventArgs e)
        {
            plDetailData.Visible = (gridAnecdoteCaseB_List.Rows.Count > 0);
        }
    }
}