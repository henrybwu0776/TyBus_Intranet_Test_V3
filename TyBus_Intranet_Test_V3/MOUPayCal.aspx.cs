using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Amaterasu_Function;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace TyBus_Intranet_Test_V3
{
    public partial class MOUPayCal : Page
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
        private DateTime vCalMonth;
        private DataTable dtTemp = new DataTable();

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
                        vCalMonth = DateTime.Parse(PF.GetMonthFirstDay(vToday.AddMonths(1), "B"));
                        eCalYear.Text = (vCalMonth.Year - 1911).ToString();
                        eCalMonth.Text = (vCalMonth.Month).ToString();
                        //取回契約辦事員名單
                        OpenEmpList();
                        //取回基本工資
                        ePayAmount.Text = PF.GetValue(vConnStr, "select top 1 LiAMT from Labor where isnull([Block], '') = 'Y' order by LiClass", "LiAMT");
                        //預設政府補助 50%
                        eGovPayRatio.Text = "50";
                        //津貼扣款項目暫定為 27
                        eBonusCode.Text = "27";
                        //計算各項金額
                        CalPaymentAmount();
                    }
                    else
                    {
                        vCalMonth = DateTime.Parse((Int32.Parse(eCalYear.Text.Trim()) + 1911).ToString("D4") + "/" + eCalMonth.Text.Trim() + "/01");
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
        /// 取回契約辦事員名單
        /// </summary>
        private void OpenEmpList()
        {
            string vTempStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                vTempStr = "select EmpNo, [Name] from Employee where Title = '271' and isnull(Leaveday, '') = '' order by EmpNo";
                SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    eDriverList.Items.Clear();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["Name"].ToString().Trim(), drTemp["EmpNo"].ToString().Trim());
                        eDriverList.Items.Add(liTemp);
                    }
                }
            }
        }

        private void CalPaymentAmount()
        {
            double vPayAmount;
            double vGovPayRatio;
            double vGovPayAMT;
            double vCompanyPay;
            if ((double.TryParse(ePayAmount.Text.Trim(), out vPayAmount)) && (double.TryParse(eGovPayRatio.Text.Trim(), out vGovPayRatio)))
            {
                vGovPayAMT = Math.Round(vPayAmount * (vGovPayRatio / 100.0), 0, MidpointRounding.AwayFromZero);
                vCompanyPay = vPayAmount - vGovPayAMT;
                eGovPayAMT.Text = vGovPayAMT.ToString();
                eCompanyPay.Text = vCompanyPay.ToString();
            }
            else
            {
                eGovPayAMT.Text = "0";
                eCompanyPay.Text = "0";
            }
        }

        protected void ePayAmount_TextChanged(object sender, EventArgs e)
        {
            CalPaymentAmount();
        }

        protected void eCompanyPay_TextChanged(object sender, EventArgs e)
        {
            double vCompanyPay;
            double vGovPayAMT;
            double vPayAmount = double.Parse(ePayAmount.Text.Trim());
            if (double.TryParse(eCompanyPay.Text.Trim(), out vCompanyPay))
            {
                vGovPayAMT = vPayAmount - vCompanyPay;
                eGovPayAMT.Text = vGovPayAMT.ToString();
            }
        }

        private void OpenData()
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            string vDriverList = "";
            if (eSelectedList.Items.Count > 0)
            {
                for (int i = 0; i < eSelectedList.Items.Count; i++)
                {
                    vDriverList = (i == 0) ? "'" + eSelectedList.Items[i].Value.Trim() + "'" : vDriverList + ", '" + eSelectedList.Items[i].Value.Trim() + "'";
                }
            }
            string vWStr_Driver = (eSelectedList.Items.Count > 0) ?
                                  "   and a.EmpNo in (" + vDriverList + ")" + Environment.NewLine :
                                  "   and b.Title = '271' and isnull(b.Leaveday, '') = '' " + Environment.NewLine;
            string vTempStr = "select a.PayDate 薪資年月, a.PayDur 發薪期別, a.EmpNo 工號, a.[Name] 姓名, b.Assumeday 到職日, " + Environment.NewLine +
                              "       a.NowPay6 生活津貼, a.CashNum27 桃市補助, a.CashNum29 其他加項, a.CashNum51 福利金, a.CashNum52 互助費, a.CashNum53 工會扣款, " + Environment.NewLine +
                              "       a.LiFee 勞保費, a.HiFee 健保費, a.GivCash 應發金額, a.NoGivCash 應扣金額, a.RealCash 實發金額 " + Environment.NewLine +
                              "  from PayRec as a left join Employee as b on b.EmpNo = a.EmpNo " + Environment.NewLine +
                              " where a.PayDate = '" + vCalMonth.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                              vWStr_Driver +
                              " order by a.EmpNo ";
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                SqlDataAdapter daTemp = new SqlDataAdapter(vTempStr, connTemp);
                connTemp.Open();
                daTemp.Fill(dtTemp);
                gvDataList.DataSource = dtTemp;
                gvDataList.DataBind();
                if (dtTemp.Rows.Count > 0)
                {
                    plShowData.Visible = true;
                }
                else
                {
                    plShowData.Visible = false;
                }
            }
        }

        protected void bbCal_Click(object sender, EventArgs e)
        {
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            //根據勾選結果逐一進行計算
            double vCompanyPay = double.Parse(eCompanyPay.Text.Trim());
            double vGovPayAMT = double.Parse(eGovPayAMT.Text.Trim());
            double vPayAmount = double.Parse(ePayAmount.Text.Trim());
            double vPayRatio = double.Parse(eGovPayRatio.Text.Trim()) / 100.0;
            double vTotalAmount = 0.0; //應發總額
            double vAmountPerHours = 0.0;
            double vRealAmount = 0.0; //公司實發
            string vEmpNo = "";
            string vNoWorkFee = "";
            string vTitle = "";
            string vTempStr = "";
            //string vTempDateStr = "";
            string vCash51Str = ""; //福利金
            string vCash52Str = ""; //互助金
            string vCash53Str = ""; //工會費
            double vCash51 = 0.0;
            double vCash52 = 0.0;
            double vCash53 = 0.0;
            double vFullMonthCount = 0.0;
            DateTime vCalFirstDate = DateTime.Parse((Int32.Parse(eCalYear.Text.Trim()) + 1911).ToString("D4") + "/" + Int32.Parse(eCalMonth.Text.Trim()).ToString("D2") + "/01");
            DateTime vCalLastDate = vCalFirstDate.AddMonths(1).AddDays(-1);
            DateTime vAssumeDay;
            string vAssumedayStr = "";
            DateTime vLeaveDay;
            string vLeavedayStr = "";
            //DateTime vStudyBoundStartDay;
            //DateTime vStudyBoundEndDay;
            double vMonthDays = (double)DateTime.DaysInMonth(Int32.Parse(eCalYear.Text.Trim()) + 1911, Int32.Parse(eCalMonth.Text.Trim()));
            double vWorkDays = 0.0;
            double vESCHour_04 = 0.0;
            double vESCHour_05 = 0.0;
            double vESCHour_13 = 0.0;
            double vCash41 = 0.0; //事假扣款
            double vCash42 = 0.0; //病假扣款
            double vCash43 = 0.0; //曠職扣款
            string vHIKind = ""; //健保類別
            string vLaIKind = ""; //勞保類別
            double vLaIAMT = 0.0; //勞保保額
            double vHIAMT = 0.0; //健保保額
            double vPension = 0.0; //勞退級距
            double vLaIFee = 0.0; //勞保保費
            double vCash71 = 0.0; //勞保公司負擔
            double vHIFee = 0.0; //健保保費
            double vCash72 = 0.0; //健保公司負擔
            double vCash57 = 0.0; //勞退自提
            double vCash74 = 0.0; //勞退公司負擔 6%
            string vRemark = ""; //備註

            for (int i = 0; i < eSelectedList.Items.Count; i++)
            {
                vEmpNo = eSelectedList.Items[i].Value.Trim(); //工號
                using (SqlConnection connTemp = new SqlConnection(vConnStr))
                {
                    vTempStr = "select NoWorkFee, Title, Assumeday, Leaveday, HIKind, HIAMT, HINum, LaIKind, LaIAMT, Pension, FixTax, " + Environment.NewLine +
                               "       case when isnull(Assumeday, '') < '" + vCalFirstDate.ToString("yyyy/MM/dd") + "' then day(EOMonth('" + vCalFirstDate.ToString("yyyy/MM/dd") + "')) " + Environment.NewLine +
                               "            else DATEDIFF(dd, Assumeday, EOMonth('" + vCalLastDate.ToString("yyyy/MM/dd") + "')) + 1 end WorkDays, " + Environment.NewLine +
                               "       DATEDIFF(MM, DateAdd(dd, -1, Assumeday), '" + vCalLastDate.ToString("yyyy/MM/dd") + "') as MonthDiff " + Environment.NewLine +
                               "  from Employee " + Environment.NewLine +
                               " where EmpNo = '" + vEmpNo + "' ";
                    SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                    connTemp.Open();
                    SqlDataReader drTemp = cmdTemp.ExecuteReader();
                    while (drTemp.Read())
                    {
                        //是否扣工會費 (3:舊會員扣100 / 2:新會員扣200 / 其他不扣款)
                        vNoWorkFee = drTemp["NoWorkFee"].ToString().Trim();
                        vCash53Str = (vNoWorkFee == "3") ? "100" : (vNoWorkFee == "2") ? "200" : "0"; //工會費
                        double.TryParse(vCash53Str, out vCash53);
                        //取回職等職級資料
                        vTitle = drTemp["Title"].ToString().Trim();
                        //計算在職天數(這裡先不算請假天數)
                        vWorkDays = double.Parse(drTemp["WorkDays"].ToString().Trim());
                        //計算到職月數
                        vFullMonthCount = double.Parse(drTemp["MonthDiff"].ToString().Trim());
                        //取回到職日
                        vAssumeDay = DateTime.Parse(drTemp["Assumeday"].ToString().Trim());
                        vAssumedayStr = vAssumeDay.ToString("yyyy/MM/dd");
                        //取回離職日 (如果有)
                        vLeavedayStr = (DateTime.TryParse(drTemp["Leaveday"].ToString().Trim(), out vLeaveDay)) ? vLeaveDay.ToString("yyyy/MM/dd") : "";
                        //取回健保類別
                        vHIKind = drTemp["HIKind"].ToString().Trim();
                        //取回勞保類別
                        vLaIKind = drTemp["LaIKind"].ToString().Trim();
                        //取回勞保保額
                        vLaIAMT = double.Parse(drTemp["LaIAMT"].ToString().Trim());
                        //取回健保保額
                        vHIAMT = double.Parse(drTemp["HIAMT"].ToString().Trim());
                        //取回勞退級距
                        vPension = double.Parse(drTemp["Pension"].ToString().Trim());
                        //取回勞退自提金額
                        vCash57 = double.Parse(drTemp["FixTax"].ToString().Trim());
                    }
                }
                //從職等職級資料取回福利金金額
                vTempStr = "select FitSet from SalaryLevel where LevelName = '" + vTitle + "' ";
                vCash51Str = PF.GetValue(vConnStr, vTempStr, "FitSet"); //福利金
                double.TryParse(vCash51Str, out vCash51);
                //從薪資系統變數取回互助金金額
                vTempStr = "select Content from SYSFLAG where FormName = 'unHRSet' and ControlItem = '664'";
                vCash52Str = PF.GetValue(vConnStr, vTempStr, "Content"); //互助金
                double.TryParse(vCash52Str, out vCash52);
                //在職天數不等於當月天數時，依比例計算薪資總額
                vTotalAmount = (vWorkDays != vMonthDays) ? Math.Round(vPayAmount * (vWorkDays / vMonthDays), 0, MidpointRounding.AwayFromZero) : vPayAmount;
                vCompanyPay = Math.Round(vTotalAmount * vPayRatio, 0, MidpointRounding.AwayFromZero);
                vGovPayAMT = vTotalAmount - vCompanyPay;
                //計算時薪 (用畫面上設定的薪資總額除以當月天數再除以8)
                vAmountPerHours = Math.Round(vPayAmount / (vMonthDays * 8), 0, MidpointRounding.AwayFromZero);
                //取回事假時數並計算扣薪
                vTempStr = "select sum(Hours) TotalHours from ESCDuty " + Environment.NewLine +
                           " where Realday between '" + vCalFirstDate.ToString("yyyy/MM/dd") + "' and '" + vCalLastDate.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                           "   and ESCType = '04' and ApplyMan = '" + vEmpNo + "' ";
                vCash41 = (double.TryParse(PF.GetValue(vConnStr, vTempStr, "TotalHours"), out vESCHour_04)) ? vESCHour_04 * vAmountPerHours : 0.0;
                //取回病假時數並計算扣薪
                vTempStr = "select sum(Hours) TotalHours from ESCDuty " + Environment.NewLine +
                           " where Realday between '" + vCalFirstDate.ToString("yyyy/MM/dd") + "' and '" + vCalLastDate.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                           "   and ESCType = '05' and ApplyMan = '" + vEmpNo + "' ";
                vCash42 = (double.TryParse(PF.GetValue(vConnStr, vTempStr, "TotalHours"), out vESCHour_05)) ? vESCHour_05 * vAmountPerHours : 0.0;
                //取回曠職時數並計算扣薪
                vTempStr = "select sum(Hours) TotalHours from ESCDuty " + Environment.NewLine +
                           " where Realday between '" + vCalFirstDate.ToString("yyyy/MM/dd") + "' and '" + vCalLastDate.ToString("yyyy/MM/dd") + "' " + Environment.NewLine +
                           "   and ESCType = '13' and ApplyMan = '" + vEmpNo + "' ";
                vCash43 = (double.TryParse(PF.GetValue(vConnStr, vTempStr, "TotalHours"), out vESCHour_13)) ? vESCHour_13 * vAmountPerHours : 0.0;
                //計算勞保費
                vLaIFee = CalLaiFee(vMonthDays, vWorkDays, vLaIAMT, vLaIKind);
                //計算健保費
                vHIFee = CalHIFee(vMonthDays, vWorkDays, vHIAMT, vHIKind);
                //計算勞保公司提撥
                vCash71 = CalLaiCompany(vMonthDays, vWorkDays, vLaIAMT, vLaIKind);
                //計算健保公司提撥
                vCash72 = CalHICompany(vMonthDays, vWorkDays, vHIAMT, vHIKind);
                //計算勞退公司提撥
                vCash74 = CalCash74(vMonthDays, vWorkDays, vPension);
                //計算實付金額
                vRealAmount = vCompanyPay - vCash41 - vCash42 - vCash43 - vLaIFee - vHIFee;
                //設定備註內容
                vRemark = (vRealAmount > 0.0) ? "發薪天數：" + vWorkDays.ToString("D0") : "發薪天數：" + vWorkDays.ToString("D0") + Environment.NewLine + "公司應付總額小於應扣總額，差額：" + Math.Abs(vRealAmount);
                //公司應發金額小於應扣總額，實發金額歸零  
                if (vRealAmount < 0.0)
                {
                    vRealAmount = 0.0;
                }
            }
        }

        /// <summary>
        /// 計算勞保費
        /// </summary>
        /// <param name="fMonthdays">當月天數</param>
        /// <param name="fWorkdays">在職天數</param>
        /// <param name="fLaIAMT">勞保保額</param>
        /// <param name="fLaIAMT">勞保保額</param>
        /// <param name="fLaIKind">勞保類別</param>
        /// <returns></returns>
        protected double CalLaiFee(double fMonthdays, double fWorkdays, double fLaIAMT, string fLaIKind)
        {
            string vTempStr = "";
            double vLaIRatio = 0.0; //勞保費率
            double vLaIRatio_Self = 0.0; //勞保自付比例
            double vHRSet304 = 0.0; //輕殘
            double vHRSet305 = 0.0; //中殘
            double vHRSet306 = 0.0; //重殘
            double vHRSet334 = 0.0; //職災
            double vCalDays = 0.0;
            double vResultAMT = 0.0;
            if (fMonthdays != fWorkdays)
            {
                vCalDays = (fMonthdays < 30) ? fWorkdays + (30 - fMonthdays) : fWorkdays;
            }
            else
            {
                vCalDays = 30.0;
            }
            vTempStr = "select ControlItem, Content from SysFlag where FormName = 'unHRSet' and ControlItem between '300' and '335' order by ControlItem ";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connGetData = new SqlConnection(vConnStr))
            {
                SqlCommand cmdGetData = new SqlCommand(vTempStr, connGetData);
                connGetData.Open();
                SqlDataReader drGetData = cmdGetData.ExecuteReader();
                while (drGetData.Read())
                {
                    switch (drGetData["ControlItem"].ToString().Trim())
                    {
                        case "301":
                            vLaIRatio = double.Parse(drGetData["Content"].ToString().Trim());
                            break;
                        case "302":
                            vLaIRatio_Self = double.Parse(drGetData["Content"].ToString().Trim());
                            break;
                        case "304":
                            vHRSet304 = double.Parse(drGetData["Content"].ToString().Trim());
                            break;
                        case "305":
                            vHRSet305 = double.Parse(drGetData["Content"].ToString().Trim());
                            break;
                        case "306":
                            vHRSet306 = double.Parse(drGetData["Content"].ToString().Trim());
                            break;
                        case "334":
                            vHRSet334 = double.Parse(drGetData["Content"].ToString().Trim());
                            break;
                    }
                }
            }
            //vResultAMT = fLaIAMT * vCalDays * vLaIRatio_Self * vLaIRatio;
            return vResultAMT;
        }

        /// <summary>
        /// 計算勞保公司負擔
        /// </summary>
        /// <param name="fMonthdays">當月天數</param>
        /// <param name="fWorkdays">在職天數</param>
        /// <param name="fLaIAMT">勞保保額</param>
        /// <param name="fLaIKind">勞保類別</param>
        /// <returns></returns>
        protected double CalLaiCompany(double fMonthdays, double fWorkdays, double fLaIAMT, string fLaIKind)
        {
            double vCalDays = 0.0;
            double vResultAMT = 0.0;
            if (fMonthdays != fWorkdays)
            {
                vCalDays = (fMonthdays < 30) ? fWorkdays + (30 - fMonthdays) : fWorkdays;
            }
            else
            {
                vCalDays = 30.0;
            }
            return vResultAMT;
        }

        /// <summary>
        /// 計算健保費
        /// </summary>
        /// <param name="fMonthdays">當月天數</param>
        /// <param name="fWorkdays">在職天數</param>
        /// <param name="fHIAMT">健保保額</param>
        /// <param name="fHIKind">健保類別</param>
        /// <returns></returns>
        protected double CalHIFee(double fMonthdays, double fWorkdays, double fHIAMT, string fHIKind)
        {
            double vResultAMT = 0.0;

            return vResultAMT;
        }

        /// <summary>
        /// 計算健保公司負擔
        /// </summary>
        /// <param name="fMonthdays">當月天數</param>
        /// <param name="fWorkdays">在職天數</param>
        /// <param name="fHIAMT">健保保額</param>
        /// <param name="fHIKind">健保類別</param>
        /// <returns></returns>
        protected double CalHICompany(double fMonthdays, double fWorkdays, double fHIAMT, string fHIKind)
        {
            double vResultAMT = 0.0;

            return vResultAMT;
        }

        /// <summary>
        /// 計算勞退公司提撥
        /// </summary>
        /// <param name="fMonthdays">當月天數</param>
        /// <param name="fWorkdays">在職天數</param>
        /// <param name="fPension">勞退保額</param>
        /// <returns></returns>
        protected double CalCash74(double fMonthdays, double fWorkdays, double fPension)
        {
            double vResultAMT = 0.0;

            return vResultAMT;
        }

        protected void bbSearch_Click(object sender, EventArgs e)
        {
            gvDataList.DataSourceID = "";
            gvDataList.DataSource = null;
            OpenData();
        }

        protected void bbExit_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/default.aspx");
        }

        /// <summary>
        /// 選擇駕駛員
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSelectTo_Click(object sender, EventArgs e)
        {
            if (eDriverList.Items.Count > 0)
            {
                for (int i = eDriverList.Items.Count - 1; i >= 0; i--)
                {
                    if (eDriverList.Items[i].Selected)
                    {
                        ListItem liTemp = eDriverList.Items[i];
                        eSelectedList.Items.Insert(0, liTemp);
                        eDriverList.Items.RemoveAt(i);
                    }
                }
            }
        }

        /// <summary>
        /// 全選
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbSelectAll_Click(object sender, EventArgs e)
        {
            string vTempStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                vTempStr = "select EmpNo, [Name] from Employee where Title = '271' and isnull(Leaveday, '') = '' order by EmpNo";
                SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    eDriverList.Items.Clear();
                    eSelectedList.Items.Clear();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["Name"].ToString().Trim(), drTemp["EmpNo"].ToString().Trim());
                        eSelectedList.Items.Add(liTemp);
                    }
                }
            }
        }

        /// <summary>
        /// 取消選擇駕駛員
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbUnselectTo_Click(object sender, EventArgs e)
        {
            if (eSelectedList.Items.Count > 0)
            {
                for (int i = eSelectedList.Items.Count - 1; i >= 0; i--)
                {
                    if (eSelectedList.Items[i].Selected)
                    {
                        ListItem liTemp = eSelectedList.Items[i];
                        eDriverList.Items.Insert(0, liTemp);
                        eSelectedList.Items.RemoveAt(i);
                    }
                }
            }
        }

        /// <summary>
        /// 取消全選
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbUnselectAll_Click(object sender, EventArgs e)
        {
            string vTempStr = "";
            if (vConnStr == "")
            {
                vConnStr = PF.GetConnectionStr(Request.ApplicationPath);
            }
            using (SqlConnection connTemp = new SqlConnection(vConnStr))
            {
                vTempStr = "select EmpNo, [Name] from Employee where Title = '271' and isnull(Leaveday, '') = '' order by EmpNo";
                SqlCommand cmdTemp = new SqlCommand(vTempStr, connTemp);
                connTemp.Open();
                SqlDataReader drTemp = cmdTemp.ExecuteReader();
                if (drTemp.HasRows)
                {
                    eDriverList.Items.Clear();
                    eSelectedList.Items.Clear();
                    while (drTemp.Read())
                    {
                        ListItem liTemp = new ListItem(drTemp["Name"].ToString().Trim(), drTemp["EmpNo"].ToString().Trim());
                        eDriverList.Items.Add(liTemp);
                    }
                }
            }
        }

        /// <summary>
        /// 計算里程碑獎金
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void bbCalGreatBounds_Click(object sender, EventArgs e)
        {

        }
    }
}