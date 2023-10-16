using Amaterasu_Function;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;

namespace TyBus_Intranet_Test_V3
{
    public partial class CalCompanyPay : Page
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
        private DateTime vFirstDate;
        private DateTime vLastDate;

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

                DateTime vToday = DateTime.Today;

                if (vLoginID != "")
                {
                    if (!IsPostBack)
                    {
                        int vYear = ((vToday.Month - 1) <= 0) ? vToday.Year - 1 : vToday.Year;
                        int vMonth = ((vToday.Month - 1) <= 0) ? vToday.Month + 11 : vToday.Month - 1;
                        eCalYM.Text = (eCalYM.Text.Trim() == "") ? (vYear.ToString("D4") + "/" + vMonth.ToString("D2")) : eCalYM.Text.Trim();
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
        /// 計算勞保公司負擔
        /// </summary>
        /// <param name="fJoinDate">加保日</param>
        /// <param name="fExitDate">退保日</param>
        /// <param name="fLaiAMT">加保保額</param>
        /// <param name="fLaiKind">加保類別</param>
        /// <param name="fLaiRatio">勞保費率 + 就業保險費率</param>
        /// <param name="fLaiBossRatio">勞保雇主費率</param>
        /// <param name="fLaiCompRatio">勞保公司負擔比率</param>
        /// <param name="fLaiAccRatio">職災費率</param>
        /// <param name="fLaiSafeRatio">就業安定費率</param>
        /// <param name="fCitzen">員工國籍</param>
        /// <returns>勞保公司負擔金額</returns>
        private double CalLaiAMT(DateTime fJoinDate, DateTime fExitDate, double fLaiAMT, string fLaiKind, double fLaiRatio, double fLaiBossRatio, double fLaiCompRatio, double fLaiAccRatio, double fLaiSafeRatio, string fCitzen)
        {
            //加保天數
            double vCalDays = 0.0;
            double vIOA = 0.0;
            double vTempAMT = 0.0;
            double vResultAMT = 0.0;

            //計算加保天數
            if ((fJoinDate.Ticks == 0) || (fJoinDate > vLastDate) || ((fExitDate.Ticks != 0) && (fExitDate < vFirstDate)))
            {
                vCalDays = 0.0;
            }
            else
            {
                vCalDays = ((fJoinDate <= vFirstDate) && (fExitDate >= vLastDate)) ? 30 : //加保日早於本月第一天且退保日晚於本月最後一天
                           ((fJoinDate <= vFirstDate) && (fExitDate.Ticks == 0)) ? 30 : //加保日早於本月第一天且沒有退保日
                           ((fJoinDate > vFirstDate) && (fExitDate.Ticks == 0)) ? (vLastDate.Day - fJoinDate.Day + 1) : //本月加保且沒有退保日
                           ((fJoinDate < vFirstDate) && ((fExitDate >= vFirstDate) && (fExitDate <= vLastDate))) ? (fExitDate.Day - vFirstDate.Day + 1) : //本月內退保
                           (((fJoinDate > vFirstDate) && (fJoinDate < vLastDate)) && ((fExitDate > vFirstDate) && (fExitDate < vLastDate))) ? (fExitDate.Day - fJoinDate.Day + 1) : //本月內加退保
                           30;
            }

            if (vCalDays > 30)
            {
                vCalDays = 30;
            }

            if ((fLaiKind != "2") && (fCitzen != "2"))
            {
                //非雇主
                //vIOA = Math.Round((fLaiAMT * fLaiRatio * fLaiCompRatio), 0, MidpointRounding.AwayFromZero);
                vIOA = Math.Ceiling((fLaiAMT * fLaiRatio * fLaiCompRatio)) +
                       Math.Round((fLaiAMT * fLaiAccRatio), 0, MidpointRounding.AwayFromZero);
            }
            else
            {
                //雇主
                //vIOA = Math.Round((fLaiAMT * fLaiBossRatio * fLaiCompRatio), 0, MidpointRounding.AwayFromZero);
                vIOA = Math.Ceiling((fLaiAMT * fLaiBossRatio * fLaiCompRatio)) +
                       Math.Round((fLaiAMT * fLaiAccRatio), 0, MidpointRounding.AwayFromZero);
            }

            if (fLaiKind == "9")
            {
                vTempAMT = Math.Round((fLaiAMT * fLaiSafeRatio), 0, MidpointRounding.AwayFromZero);
            }
            else
            {
                //vTempAMT = vIOA +
                //           Math.Round((fLaiAMT * fLaiAccRatio), 0, MidpointRounding.AwayFromZero) +
                //           Math.Round((fLaiAMT * fLaiSafeRatio), 0, MidpointRounding.AwayFromZero); //*/
                vTempAMT = vIOA;
            }
            vResultAMT = Math.Round(vTempAMT * (vCalDays / 30), 0, MidpointRounding.AwayFromZero);
            return vResultAMT;
        }

        /// <summary>
        /// 計算勞保墊償基金提撥金額
        /// </summary>
        /// <param name="fJoinDate">加保日</param>
        /// <param name="fExitDate">退保日</param>
        /// <param name="fLaiAMT">加保保額</param>
        /// <param name="fLaiSafeRatio">勞保墊償基金提撥費率</param>
        /// <returns>勞保墊償基金提撥金額</returns>
        private double CalLaiSafeAMT(DateTime fJoinDate, DateTime fExitDate, double fLaiAMT, double fLaiSafeRatio)
        {
            //加保天數
            double vCalDays = 0.0;
            double vTempAMT = 0.0;
            double vResultAMT = 0.0;

            //計算加保天數
            if ((fJoinDate.Ticks == 0) || (fJoinDate > vLastDate) || ((fExitDate.Ticks != 0) && (fExitDate < vFirstDate)))
            {
                vCalDays = 0.0;
            }
            else
            {
                vCalDays = ((fJoinDate <= vFirstDate) && (fExitDate >= vLastDate)) ? 30 : //加保日早於本月第一天且退保日晚於本月最後一天
                           ((fJoinDate <= vFirstDate) && (fExitDate.Ticks == 0)) ? 30 : //加保日早於本月第一天且沒有退保日
                           ((fJoinDate > vFirstDate) && (fExitDate.Ticks == 0)) ? (vLastDate.Day - fJoinDate.Day + 1) : //本月加保且沒有退保日
                           ((fJoinDate < vFirstDate) && ((fExitDate >= vFirstDate) && (fExitDate <= vLastDate))) ? (fExitDate.Day - vFirstDate.Day + 1) : //本月內退保
                           (((fJoinDate > vFirstDate) && (fJoinDate < vLastDate)) && ((fExitDate > vFirstDate) && (fExitDate < vLastDate))) ? (fExitDate.Day - fJoinDate.Day + 1) : //本月內加退保
                           30;
            }

            if (vCalDays > 30)
            {
                vCalDays = 30;
            }

            vTempAMT = Math.Round((fLaiAMT * fLaiSafeRatio), 0, MidpointRounding.AwayFromZero);
            vResultAMT = Math.Round(vTempAMT * (vCalDays / 30), 0, MidpointRounding.AwayFromZero);
            return vResultAMT;
        }

        /// <summary>
        /// 計算健保公司負擔
        /// </summary>
        /// <param name="fHiJoinDate">健保加保日</param>
        /// <param name="fHiExitDate">健保退保日</param>
        /// <param name="fHiKind">投保類別</param>
        /// <param name="fHiRatio">健保費率</param>
        /// <param name="fHiNum">眷口數</param>
        /// <param name="fHiAMT">健保投保金額</param>
        /// <param name="fHiCompRatio">健保公司負擔比率</param>
        /// <param name="fHiAvgNum">健保平均眷口數</param>
        /// <returns>健保公司負擔金額</returns>
        private double CalHiAMT(DateTime fHiJoinDate, DateTime fHiExitDate, string fHiKind, double fHiRatio, int fHiNum, double fHiAMT, double fHiCompRatio, double fHiAvgNum)
        {
            double vResultAMT = 0.0;
            if ((fHiJoinDate.Ticks == 0) || (fHiJoinDate > vLastDate) || ((fHiExitDate < vFirstDate) && (fHiExitDate.Ticks != 0)) || ((fHiExitDate >= vFirstDate) && (fHiExitDate <= vLastDate)))
            {
                vResultAMT = 0;
            }
            else
            {
                if (fHiKind != "2")
                {
                    vResultAMT = Math.Round(fHiAMT * fHiRatio * fHiCompRatio * (1 + fHiAvgNum), 0, MidpointRounding.AwayFromZero);
                }
                else
                {
                    vResultAMT = 0.0;
                }
            }
            return vResultAMT;
        }

        protected void bbOK_Click(object sender, EventArgs e)
        {
            plShowData.Visible = false;
            gridData.DataSourceID = "";
            //建立一個暫時的資料表
            DataTable dtResult = new DataTable("CompanyPay");
            //定義資料表的欄位性質
            DataColumn colDepNo = new DataColumn("DepNo", Type.GetType("System.String"));
            DataColumn colDepName = new DataColumn("DepName", Type.GetType("System.String"));
            DataColumn colEmpNo = new DataColumn("EmpNo", Type.GetType("System.String"));
            DataColumn colEmpName = new DataColumn("EmpName", Type.GetType("System.String"));
            DataColumn colCalDate = new DataColumn("CalDate", Type.GetType("System.DateTime"));
            DataColumn colLaiCompanyAMT = new DataColumn("LaiCompanyAMT", Type.GetType("System.Double"));
            DataColumn colHiCompanyAMT = new DataColumn("HiCompanyAMT", Type.GetType("System.Double"));
            DataColumn colLaiRetireAMT = new DataColumn("LaiRetireAMT", Type.GetType("System.Double"));
            DataColumn colLaiSafeAMT = new DataColumn("LaiSafeAMT", Type.GetType("System.Double"));
            //建立欄位
            dtResult.Columns.Add(colDepNo);
            dtResult.Columns.Add(colDepName);
            dtResult.Columns.Add(colEmpNo);
            dtResult.Columns.Add(colEmpName);
            dtResult.Columns.Add(colCalDate);
            dtResult.Columns.Add(colLaiCompanyAMT);
            dtResult.Columns.Add(colHiCompanyAMT);
            dtResult.Columns.Add(colLaiRetireAMT);
            dtResult.Columns.Add(colLaiSafeAMT);

            //宣告其他變數
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            //==========================================================================================
            string vRecordYMStr = eCalYM.Text.Trim();
            string vRecordNote = "計算數值：重新計算公司提撥" + Environment.NewLine +
                                 "CalCompanyPay.aspx" + Environment.NewLine +
                                 "計算年月：" + vRecordYMStr;
            PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
            //==========================================================================================
            string vSelectSQLStr = "";
            string vCalYM = eCalYM.Text.Trim() + "/01";
            string vTempStr = "";
            string vDepNo_Temp = "";
            string vDepName_Temp = "";
            string vEmpNo = "";
            string vEmpName = "";
            string vCitzen = ""; //員工國籍
            string vLaiKind = ""; //勞保投保類別
            string vHiKind = ""; //健保投保類別
            int vHiNum = 0; //健保眷口數
            double vHiAvgNum = 0.0; //健保平均眷口數
            DateTime vLaiJoinDate = DateTime.FromBinary(0); //勞保加保日
            DateTime vLaiExitDate = DateTime.FromBinary(0); //勞保退保日
            DateTime vHiJoinDate = DateTime.FromBinary(0); //健保加保日
            DateTime vHiExitDate = DateTime.FromBinary(0); //健保退保日
            int vWorkDays = 0; //上班天數
            int vMonthDays = 0; //當月天數
            double vHiAMT = 0; //健保級距
            double vLaiAMT = 0; //勞保級距
            double vLaiRetireAMT = 0; //勞退級距
            double vLaiRatio = 0; //勞保費率 + 就業保險費率
            double vLaiBossRatio = 0; //勞保雇主費率
            double vLaiCompRatio = 0; //勞保公司負擔比率
            double vLaiAccRatio = 0; //職災費率
            double vLaiSafeRatio = 0; //就業安定費率
            double vHiRatio = 0; //健保費率
            double vHiCompRatio = 0; //健保公司負擔比率
            double vSevDisability = 0; //重殘比率
            double vModDisability = 0; //中殘比率
            double vMildDisability = 0; //輕殘比率
            double vOld_TP = 0; //北市老人
            double vOld_KH = 0; //高市老人
            double vOld_TY = 0; //桃市老人
            double vLowIncome = 0; //中低收入戶
            double vRetireRatio = 0; //勞退比率
                                     //double vMaxLevelRetire = 150000; //勞退級距上限
            double vCashNum71 = 0; //勞保公司負擔
            double vCashNum72 = 0; //健保公司負擔
            double vCashNum74 = 0; //勞退公司提撥
            double vLaiSafeAMT = 0; //勞保墊償基金提撥

            //計算當月第一天跟最後一天
            vFirstDate = DateTime.Parse(PF.GetMonthFirstDay(DateTime.Parse(vCalYM), "B"));
            vLastDate = DateTime.Parse(PF.GetMonthLastDay(DateTime.Parse(vCalYM), "B"));

            //以下取回相關的數字
            //勞保費率 + 就業保險費率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '301' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vLaiRatio = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //勞保雇主費率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '336' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vLaiBossRatio = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //勞保公司負擔比率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '303' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vLaiCompRatio = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //勞保墊償基金提撥費率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '333' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vLaiSafeRatio = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //職災費率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '334' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vLaiAccRatio = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //健保費率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '308' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vHiRatio = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //健保公司負擔比率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '310' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vHiCompRatio = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //健保平均眷口數
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '332' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vHiAvgNum = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //重殘比率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '313' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vSevDisability = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //中殘比率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '312' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vModDisability = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //輕殘比率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '311' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vMildDisability = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //北市老人
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '315' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vOld_TP = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //高市老人
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '316' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vOld_KH = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //桃市老人
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '352' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vOld_TY = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //中低收入
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '331' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vLowIncome = (vTempStr != "") ? double.Parse(vTempStr) : 0;
            //勞退比率
            vSelectSQLStr = "select Content from SysFlag where FormName = 'unHrSet' and ControlItem = '674' ";
            vTempStr = PF.GetValue(vConnStr, vSelectSQLStr, "Content");
            vRetireRatio = (vTempStr != "") ? double.Parse(vTempStr) : 0;

            vSelectSQLStr = "select DepNo, EmpNo, [Name], HiAMT, LaiAMT from PayRec where PayDate = '" + vCalYM + "' and PayDur = '1' order by DepNo, EmpNo";
            using (SqlConnection connCalSource = new SqlConnection(vConnStr))
            {
                SqlCommand cmdCalSource = new SqlCommand(vSelectSQLStr, connCalSource);
                connCalSource.Open();
                SqlDataReader drCalSource = cmdCalSource.ExecuteReader();
                if (drCalSource.HasRows)
                {
                    while (drCalSource.Read())
                    {
                        //取回所屬單位
                        vDepNo_Temp = drCalSource["DepNo"].ToString().Trim();
                        vTempStr = "select [Name] from Department where DepNo = '" + vDepNo_Temp + "' ";
                        vDepName_Temp = PF.GetValue(vConnStr, vTempStr, "Name");
                        //取回員工工號
                        vEmpNo = drCalSource["EmpNo"].ToString().Trim();
                        //取回員工姓名
                        vEmpName = drCalSource["Name"].ToString().Trim();
                        //健保投保金額
                        vHiAMT = double.Parse(drCalSource["HiAMT"].ToString());
                        //勞保投保金額
                        vLaiAMT = double.Parse(drCalSource["LaiAMT"].ToString());
                        //勞退級距
                        //vLaiRetireAMT = (vHiAMT >= vMaxLevelRetire) ? vMaxLevelRetire : vHiAMT;
                        vTempStr = "select isnull(Pension, 0) Pension from Employee where EmpNo = '" + vEmpNo + "' ";
                        vLaiRetireAMT = double.Parse(PF.GetValue(vConnStr, vTempStr, "Pension"));
                        //計算當月天數
                        vMonthDays = PF.GetMonthDays(vFirstDate);
                        //計算上班天數
                        vWorkDays = PF.CalEmployeeWorkDays(vConnStr, vEmpNo, vFirstDate, vLastDate);
                        vWorkDays = (vWorkDays == vMonthDays) ? 30 : (vWorkDays > 30) ? 30 : vWorkDays;
                        //計算勞退金額
                        //vCashNum74 = int.Parse(Math.Ceiling(((vLaiRetireAMT * vRetireRatio) / 30) * vWorkDays).ToString());
                        vCashNum74 = Math.Round((((vLaiRetireAMT * vRetireRatio) / 30) * vWorkDays), 0, MidpointRounding.AwayFromZero);
                        //取回勞健保加保日期類別和員工國籍以及健保眷口數
                        vTempStr = "select isnull(convert(char(10), HiJoinDay, 111), 0) HiJoinDay, " + Environment.NewLine +
                                   "       isnull(convert(char(10), HiExitDay, 111), 0) HiExitDay, " + Environment.NewLine +
                                   "       isnull(convert(char(10), LaiJoinDay, 111), 0) LaiJoinDay, " + Environment.NewLine +
                                   "       isnull(convert(char(10), LaiExitDay, 111), 0) LaiExitDay, " + Environment.NewLine +
                                   "       Citzen, LaiKind, HiKind, HiNum " + Environment.NewLine +
                                   "  from Employee where EmpNo = '" + vEmpNo + "' ";
                        SqlConnection connGetEmpData = new SqlConnection(vConnStr);
                        SqlCommand cmdGetEMpData = new SqlCommand(vTempStr, connGetEmpData);
                        connGetEmpData.Open();
                        SqlDataReader drGetEmpData = cmdGetEMpData.ExecuteReader();
                        if (drGetEmpData.HasRows)
                        {
                            while (drGetEmpData.Read())
                            {
                                vHiJoinDate = (drGetEmpData["HiJoinDay"].ToString().Trim() != "0") ? DateTime.Parse(drGetEmpData["HiJoinDay"].ToString().Trim()) : DateTime.FromBinary(0);
                                vHiExitDate = (drGetEmpData["HiExitDay"].ToString().Trim() != "0") ? DateTime.Parse(drGetEmpData["HiExitDay"].ToString().Trim()) : DateTime.FromBinary(0);
                                vLaiJoinDate = (drGetEmpData["LaiJoinDay"].ToString().Trim() != "0") ? DateTime.Parse(drGetEmpData["LaiJoinDay"].ToString().Trim()) : DateTime.FromBinary(0);
                                vLaiExitDate = (drGetEmpData["LaiExitDay"].ToString().Trim() != "0") ? DateTime.Parse(drGetEmpData["LaiExitDay"].ToString().Trim()) : DateTime.FromBinary(0);
                                vCitzen = drGetEmpData["Citzen"].ToString().Trim();
                                vLaiKind = drGetEmpData["LaiKind"].ToString().Trim();
                                vHiKind = drGetEmpData["HiKind"].ToString().Trim();
                                vHiNum = int.Parse(drGetEmpData["HiNum"].ToString().Trim());
                            }
                        }
                        else
                        {
                            vHiJoinDate = DateTime.FromBinary(0);
                            vHiExitDate = DateTime.FromBinary(0);
                            vLaiJoinDate = DateTime.FromBinary(0);
                            vLaiExitDate = DateTime.FromBinary(0);
                            vLaiKind = "";
                            vHiKind = "";
                            vHiNum = 0;
                            vCitzen = "";
                        }
                        drGetEmpData.Close();
                        cmdGetEMpData.Dispose();
                        //計算勞保公司負擔金額
                        vCashNum71 = CalLaiAMT(vLaiJoinDate, vLaiExitDate, vLaiAMT, vLaiKind, vLaiRatio, vLaiBossRatio, vLaiCompRatio, vLaiAccRatio, vLaiSafeRatio, vCitzen);
                        //計算勞保墊償基金提撥金額
                        vLaiSafeAMT = CalLaiSafeAMT(vLaiJoinDate, vLaiExitDate, vLaiAMT, vLaiSafeRatio);
                        //計算健保公司負擔金額
                        vCashNum72 = CalHiAMT(vHiJoinDate, vHiExitDate, vHiKind, vHiRatio, vHiNum, vHiAMT, vHiCompRatio, vHiAvgNum);
                        //建立一筆資料
                        DataRow rwTemp = dtResult.NewRow();
                        rwTemp["DepNo"] = vDepNo_Temp;
                        rwTemp["DepName"] = vDepName_Temp;
                        rwTemp["EmpNo"] = vEmpNo;
                        rwTemp["EmpName"] = vEmpName;
                        rwTemp["CalDate"] = DateTime.Parse(vCalYM);
                        rwTemp["LaiCompanyAMT"] = vCashNum71;
                        rwTemp["HiCompanyAMT"] = vCashNum72;
                        rwTemp["LaiRetireAMT"] = vCashNum74;
                        rwTemp["LaiSafeAMT"] = vLaiSafeAMT;
                        dtResult.Rows.Add(rwTemp);
                    }
                    gridData.DataSource = dtResult;
                    gridData.DataBind();
                    plShowData.Visible = true;
                }
            }
        }

        protected void bbClose_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        protected void bbExportExcel_Click(object sender, EventArgs e)
        {
            string vHeadText = "";
            string vGetDate = "";
            DateTime vTempDate;
            //Excel 檔案
            XSSFWorkbook wbExcel = new XSSFWorkbook();
            //Excel 工作表
            XSSFSheet wsExcel = (XSSFSheet)wbExcel.CreateSheet("計算結果");
            //設定標題欄位的格式
            XSSFCellStyle csTitle = (XSSFCellStyle)wbExcel.CreateCellStyle();
            csTitle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            csTitle.Alignment = HorizontalAlignment.Center; //水平置中
                                                            //設定字體格式
            XSSFFont fontTitle = (XSSFFont)wbExcel.CreateFont();
            //fontTitle.Boldweight = (short)FontBoldWeight.Bold; //粗體字
            fontTitle.IsBold = true;
            fontTitle.FontHeightInPoints = 20; //字體大小
            csTitle.SetFont(fontTitle);

            //建立標題列
            wsExcel.CreateRow(0);
            for (int i = 0; i < gridData.Columns.Count; i++)
            {
                vHeadText = (gridData.Columns[i].HeaderText == "EmpNo") ? "員工工號" :
                            (gridData.Columns[i].HeaderText == "CalDate") ? "發薪年月" :
                            (gridData.Columns[i].HeaderText == "LaiCompanyAMT") ? "勞保公司負擔" :
                            (gridData.Columns[i].HeaderText == "HiCompanyAMT") ? "健保公司負擔" :
                            (gridData.Columns[i].HeaderText == "LaiRetireAMT") ? "勞退提撥" :
                            gridData.Columns[i].HeaderText;
                wsExcel.GetRow(0).CreateCell(i).SetCellValue(vHeadText);
            }
            for (int j = 0; j < gridData.Rows.Count; j++)
            {
                wsExcel.CreateRow(j + 1);
                for (int k = 0; k < gridData.Columns.Count; k++)
                {
                    if ((gridData.Columns[k].HeaderText == "勞保公司負擔") ||
                        (gridData.Columns[k].HeaderText == "健保公司負擔") ||
                        (gridData.Columns[k].HeaderText == "勞退公司提撥") ||
                        (gridData.Columns[k].HeaderText == "勞保墊償基金提撥"))
                    {
                        //這幾個欄位是數值欄位
                        wsExcel.GetRow(j + 1).CreateCell(k);
                        wsExcel.GetRow(j + 1).GetCell(k).SetCellType(CellType.Numeric);
                        wsExcel.GetRow(j + 1).GetCell(k).SetCellValue(double.Parse(gridData.Rows[j].Cells[k].Text.Trim()));
                    }
                    else
                    {
                        wsExcel.GetRow(j + 1).CreateCell(k);
                        wsExcel.GetRow(j + 1).GetCell(k).SetCellType(CellType.String);
                        if (gridData.Columns[k].HeaderText == "CalDate")
                        {
                            vTempDate = DateTime.Parse(gridData.Rows[j].Cells[k].Text.Trim());
                            vGetDate = (vTempDate.Year - 1911).ToString("D3") + "/" + vTempDate.ToString("MM/dd");
                            wsExcel.GetRow(j + 1).GetCell(k).SetCellValue(vGetDate);
                        }
                        else
                        {
                            wsExcel.GetRow(j + 1).GetCell(k).SetCellValue(gridData.Rows[j].Cells[k].Text.Trim());
                        }
                    }
                }
            }
            try
            {
                /*
                MemoryStream msTarget = new MemoryStream();
                wbExcel.Write(msTarget);
                // 設定強制下載標頭
                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=CalCompanyPay.xlsx"));
                // 輸出檔案
                Response.BinaryWrite(msTarget.ToArray());
                msTarget.Close();

                Response.End(); //*/
                var msTarget = new NPOIMemoryStream();
                msTarget.AllowClose = false;
                wbExcel.Write(msTarget);
                msTarget.Flush();
                msTarget.Seek(0, SeekOrigin.Begin);
                msTarget.AllowClose = true;

                if (msTarget.Length > 0)
                {
                    //==========================================================================================
                    string vRecordYMStr = eCalYM.Text.Trim();
                    string vRecordNote = "匯出檔案：重新計算公司提撥" + Environment.NewLine +
                                         "CalCompanyPay.aspx" + Environment.NewLine +
                                         "計算年月：" + vRecordYMStr;
                    PF.InsertOperateRecord(vConnStr, vLoginID, vComputerName, vRecordNote);
                    //==========================================================================================
                    //先判斷是不是IE
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    HttpBrowserCapabilities brObject = HttpContext.Current.Request.Browser;
                    string TourVerision = brObject.Type;
                    string vFileName = "CalCompanyPay";
                    if ((TourVerision.Substring(0, 2).ToUpper() == "IE") || (TourVerision == "InternetExplorer11"))
                    {
                        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("" + vFileName + ".xls", System.Text.Encoding.UTF8));
                    }
                    else if (TourVerision.Contains("Chrome") || TourVerision.Contains("Firefox"))
                    {
                        // 設定強制下載標頭
                        Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + vFileName + ".xls"));
                    }
                    // 輸出檔案
                    Response.BinaryWrite(msTarget.ToArray());
                    msTarget.Close();

                    Response.End();
                }
            }
            catch (Exception eMessage)
            {
                Response.Write("<Script language='Javascript'>");
                Response.Write("alert('" + eMessage.Message + "')");
                Response.Write("</" + "Script>");
                //throw;
            }
        }
    }
}